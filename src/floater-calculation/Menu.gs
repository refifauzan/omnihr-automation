/**
 * Menu functions for Floater Calculation
 */

/**
 * Create the Floater menu when the spreadsheet opens
 */
function onOpen() {
	const ui = SpreadsheetApp.getUi();
	ui.createMenu('Floater')
		.addItem('Generate Floater View', 'generateFloaterViewMenu')
		.addItem('Update Current View', 'updateFloaterView')
		.addSeparator()
		.addSubMenu(
			ui
				.createMenu('Schedule')
				.addItem('Enable Weekly Updates', 'setupWeeklyTrigger')
				.addItem('Disable Weekly Updates', 'removeWeeklyTrigger')
				.addSeparator()
				.addItem('View Current Schedule', 'viewTriggers'),
		)
		.addSeparator()
		.addItem('Setup API Credentials', 'setupCredentials')
		.addToUi();
}

/**
 * Menu handler - Generate Floater View with month/year prompts
 */
function generateFloaterViewMenu() {
	const ui = SpreadsheetApp.getUi();

	const monthResponse = ui.prompt(
		'Generate Floater View',
		'Enter month (1-12):',
		ui.ButtonSet.OK_CANCEL,
	);

	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	const yearResponse = ui.prompt(
		'Generate Floater View',
		'Enter year (e.g., 2026):',
		ui.ButtonSet.OK_CANCEL,
	);

	if (yearResponse.getSelectedButton() !== ui.Button.OK) return;

	const month = parseInt(monthResponse.getResponseText()) - 1;
	const year = parseInt(yearResponse.getResponseText());

	if (isNaN(month) || month < 0 || month > 11 || isNaN(year)) {
		ui.alert('Invalid month or year');
		return;
	}

	generateFloaterView(month, year);
}

/**
 * Setup API credentials in Script Properties
 */
function setupCredentials() {
	const ui = SpreadsheetApp.getUi();
	const props = PropertiesService.getScriptProperties();

	const baseUrlResponse = ui.prompt(
		'API Base URL',
		'Enter OmniHR API base URL (default: https://api.omnihr.co/api/v1):',
		ui.ButtonSet.OK_CANCEL,
	);

	if (baseUrlResponse.getSelectedButton() !== ui.Button.OK) return;

	const subdomainResponse = ui.prompt(
		'Subdomain',
		'Enter OmniHR subdomain (e.g., snappymob):',
		ui.ButtonSet.OK_CANCEL,
	);

	if (subdomainResponse.getSelectedButton() !== ui.Button.OK) return;

	const usernameResponse = ui.prompt(
		'Username',
		'Enter OmniHR username (email):',
		ui.ButtonSet.OK_CANCEL,
	);

	if (usernameResponse.getSelectedButton() !== ui.Button.OK) return;

	const passwordResponse = ui.prompt(
		'Password',
		'Enter OmniHR password:',
		ui.ButtonSet.OK_CANCEL,
	);

	if (passwordResponse.getSelectedButton() !== ui.Button.OK) return;

	const baseUrl =
		baseUrlResponse.getResponseText().trim() || 'https://api.omnihr.co/api/v1';

	props.setProperty('OMNIHR_BASE_URL', baseUrl);
	props.setProperty(
		'OMNIHR_SUBDOMAIN',
		subdomainResponse.getResponseText().trim(),
	);
	props.setProperty(
		'OMNIHR_USERNAME',
		usernameResponse.getResponseText().trim(),
	);
	props.setProperty(
		'OMNIHR_PASSWORD',
		passwordResponse.getResponseText().trim(),
	);

	ui.alert('API credentials saved successfully!');
}

/**
 * Setup weekly trigger to auto-update floater view every Monday at 8 AM
 */
function setupWeeklyTrigger() {
	const ui = SpreadsheetApp.getUi();

	// Remove existing weekly triggers first
	removeWeeklyTriggerSilent();

	ScriptApp.newTrigger('scheduledWeeklyFloaterUpdate')
		.timeBased()
		.onWeekDay(ScriptApp.WeekDay.MONDAY)
		.atHour(8)
		.create();

	ui.alert(
		'Weekly update enabled!\n\nThe floater view for the current month will be automatically updated every Monday at 8 AM.',
	);
}

/**
 * Remove weekly trigger
 */
function removeWeeklyTrigger() {
	const ui = SpreadsheetApp.getUi();
	const removed = removeWeeklyTriggerSilent();

	if (removed > 0) {
		ui.alert(`Removed ${removed} weekly trigger(s).`);
	} else {
		ui.alert('No weekly triggers found.');
	}
}

/**
 * Remove weekly triggers silently (no UI alert)
 * @returns {number} Number of triggers removed
 */
function removeWeeklyTriggerSilent() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const triggers = ScriptApp.getUserTriggers(ss);
	let removed = 0;

	for (const trigger of triggers) {
		if (trigger.getHandlerFunction() === 'scheduledWeeklyFloaterUpdate') {
			ScriptApp.deleteTrigger(trigger);
			removed++;
		}
	}

	return removed;
}

/**
 * View current triggers
 */
function viewTriggers() {
	const ui = SpreadsheetApp.getUi();
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const triggers = ScriptApp.getUserTriggers(ss);

	if (triggers.length === 0) {
		ui.alert('No active triggers found.');
		return;
	}

	let message = 'Active Triggers:\n\n';
	for (const trigger of triggers) {
		message += `• ${trigger.getHandlerFunction()} - ${trigger.getEventType()}\n`;
	}

	ui.alert(message);
}

/**
 * Scheduled function called by the weekly trigger
 * Updates the floater view for the current month
 */
function scheduledWeeklyFloaterUpdate() {
	const now = new Date();
	const month = now.getMonth();
	const year = now.getFullYear();

	Logger.log(`Scheduled weekly floater update for ${month + 1}/${year}`);

	try {
		const ss = SpreadsheetApp.getActiveSpreadsheet();
		const monthNames = [
			'January',
			'February',
			'March',
			'April',
			'May',
			'June',
			'July',
			'August',
			'September',
			'October',
			'November',
			'December',
		];
		const sheetName = `Floaters ${monthNames[month]} ${year}`;

		// Delete existing sheet if it exists
		const existingSheet = ss.getSheetByName(sheetName);
		if (existingSheet) {
			ss.deleteSheet(existingSheet);
		}

		const sheet = ss.insertSheet(sheetName);

		const token = getAccessToken();
		if (!token) {
			Logger.log('Failed to get API token for scheduled update');
			return;
		}

		const employeesWithDetails = fetchAllEmployeesWithDetails(token);
		const holidays = fetchHolidaysForMonth(token, month, year);
		const holidayDays = new Set(holidays.map((h) => h.date));
		const workingDays = countWorkingDays(month, year, holidayDays);
		const cvData = readCapacityViewData(month, year);

		const floaterData = buildFloaterData(
			employeesWithDetails,
			cvData,
			month,
			year,
			workingDays,
		);

		floaterData.sort((a, b) => {
			if (a.isLeaver && !b.isLeaver) return 1;
			if (!a.isLeaver && b.isLeaver) return -1;
			return b.floaterPct - a.floaterPct;
		});

		writeFloaterSheet(sheet, floaterData, monthNames[month], year);
		SpreadsheetApp.flush();

		Logger.log(
			`Weekly floater update completed: ${floaterData.length} employees`,
		);
	} catch (error) {
		Logger.log(
			'Error in scheduled weekly floater update: ' +
				error.message +
				'\n' +
				error.stack,
		);
	}
}
