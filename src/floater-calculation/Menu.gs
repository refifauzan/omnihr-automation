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
		.addItem('Setup Token Service', 'showTokenServiceDialog')
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
 * Show dialog to configure the remote Token Service
 */
function showTokenServiceDialog() {
	const props = PropertiesService.getScriptProperties();
	const currentUrl = props.getProperty('TOKEN_SERVICE_URL') || '';
	const currentKey = props.getProperty('TOKEN_SERVICE_API_KEY') || '';

	const html = HtmlService.createHtmlOutput(
		'<style>' +
			'  body { font-family: Arial, sans-serif; padding: 16px; max-width: 420px; }' +
			'  label { display: block; margin-top: 12px; font-weight: bold; }' +
			'  input { width: 100%; padding: 8px; margin-top: 5px; box-sizing: border-box; }' +
			'  .hint { font-size: 12px; color: #666; margin-top: 4px; }' +
			'  button { margin-top: 16px; padding: 10px 24px; background: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; }' +
			'  button:hover { background: #1557b0; }' +
			'  .status { margin-top: 12px; padding: 8px; border-radius: 4px; display: none; }' +
			'  .status.ok { display: block; background: #e6f4ea; color: #137333; }' +
			'  .status.err { display: block; background: #fce8e6; color: #c5221f; }' +
			'</style>' +
			'<p class="hint">Connect to the OmniHR Token Service running on your VPS.</p>' +
			'<label>Token Service URL</label>' +
			'<input type="text" id="serviceUrl" placeholder="https://snappy-omnihr.djebreds.com" value="' +
			currentUrl +
			'">' +
			'<label>API Secret Key</label>' +
			'<input type="password" id="apiKey" placeholder="Shared secret from Token Service .env" value="' +
			currentKey +
			'">' +
			'<label>OmniHR Base URL</label>' +
			'<input type="text" id="baseUrl" value="' +
			(props.getProperty('OMNIHR_BASE_URL') || 'https://api.omnihr.co/api/v1') +
			'">' +
			'<label>OmniHR Subdomain</label>' +
			'<input type="text" id="subdomain" value="' +
			(props.getProperty('OMNIHR_SUBDOMAIN') || '') +
			'">' +
			'<button onclick="save()">Save & Test Connection</button>' +
			'<div id="status" class="status"></div>' +
			'<script>' +
			'  function save() {' +
			'    const url = document.getElementById("serviceUrl").value.trim();' +
			'    const key = document.getElementById("apiKey").value.trim();' +
			'    const baseUrl = document.getElementById("baseUrl").value.trim();' +
			'    const subdomain = document.getElementById("subdomain").value.trim();' +
			'    if (!url || !key || !baseUrl || !subdomain) { showStatus("Fill in all fields.", true); return; }' +
			'    document.getElementById("status").style.display = "block";' +
			'    document.getElementById("status").innerText = "Testing connection...";' +
			'    google.script.run' +
			'      .withSuccessHandler(function(msg) { showStatus(msg, false); })' +
			'      .withFailureHandler(function(err) { showStatus(err.message || String(err), true); })' +
			'      .saveTokenServiceConfig(url, key, baseUrl, subdomain);' +
			'  }' +
			'  function showStatus(msg, isError) {' +
			'    const el = document.getElementById("status");' +
			'    el.className = "status " + (isError ? "err" : "ok");' +
			'    el.innerText = msg;' +
			'  }' +
			'</script>',
	)
		.setWidth(460)
		.setHeight(480);

	SpreadsheetApp.getUi().showModalDialog(
		html,
		'Floater \u2014 Token Service Setup',
	);
}

/**
 * Save Token Service configuration and test the connection.
 */
function saveTokenServiceConfig(serviceUrl, apiKey, baseUrl, subdomain) {
	const testUrl = serviceUrl.replace(/\/+$/, '') + '/api/health';
	try {
		const healthResponse = UrlFetchApp.fetch(testUrl, {
			method: 'get',
			muteHttpExceptions: true,
		});
		if (healthResponse.getResponseCode() !== 200) {
			throw new Error(
				'Health check returned ' + healthResponse.getResponseCode(),
			);
		}
	} catch (e) {
		throw new Error('Cannot reach Token Service: ' + e.message);
	}

	const statusUrl = serviceUrl.replace(/\/+$/, '') + '/api/status';
	let statusResponse;
	try {
		statusResponse = UrlFetchApp.fetch(statusUrl, {
			method: 'get',
			headers: { 'X-API-Key': apiKey },
			muteHttpExceptions: true,
		});
		if (statusResponse.getResponseCode() === 403) {
			throw new Error('API key is invalid (403 Forbidden)');
		}
		if (statusResponse.getResponseCode() !== 200) {
			throw new Error(
				'Status endpoint returned ' + statusResponse.getResponseCode(),
			);
		}
	} catch (e) {
		throw new Error('Authentication failed: ' + e.message);
	}

	const props = PropertiesService.getScriptProperties();
	props.setProperty('TOKEN_SERVICE_URL', serviceUrl.replace(/\/+$/, ''));
	props.setProperty('TOKEN_SERVICE_API_KEY', apiKey);
	props.setProperty('OMNIHR_BASE_URL', baseUrl);
	props.setProperty('OMNIHR_SUBDOMAIN', subdomain);

	const statusData = JSON.parse(statusResponse.getContentText());
	let msg = 'Connected successfully!';
	if (statusData.has_access_token) {
		msg +=
			' Token available (age: ' +
			Math.round(statusData.access_token_age_seconds) +
			's).';
	} else {
		msg += ' Service is running but no token cached yet.';
	}
	return msg;
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
		const leaveData = fetchLeaveDataForMonth(
			token,
			employeesWithDetails,
			month,
			year,
		);
		const holidays = fetchHolidaysForMonth(token, month, year);
		const holidayDays = new Set(holidays.map((h) => h.date));
		const workingDays = countWorkingDays(month, year, holidayDays);
		const allocationData = readProjectSheetAllocation(
			ss,
			month,
			year,
			holidayDays,
			workingDays,
		);

		const floaterData = buildFloaterData(
			employeesWithDetails,
			allocationData,
			leaveData,
			month,
			year,
			holidayDays,
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
