/**
 * OmniHR Leave Data Integration for Google Sheets
 *
 * Main sync functions - entry points for leave data synchronization
 *
 * Files in this project:
 * - Config.gs: Configuration constants
 * - Menu.gs: Menu and trigger functions
 * - Api.gs: OmniHR API functions
 * - Utils.gs: Utility functions
 * - LeaveService.gs: Leave data processing
 * - Attendance.gs: Attendance sheet functions
 * - Code.gs: Main sync functions (this file)
 */

/**
 * Sync leave data for current month to the active sheet
 */
function syncCurrentMonth() {
	const now = new Date();
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getActiveSheet();

	Logger.log(`Syncing current month to active sheet: ${sheet.getName()}`);
	syncLeaveDataToSheet(sheet, now.getMonth(), now.getFullYear());
}

/**
 * Sync leave data - prompts for month/year
 */
function syncLeaveData() {
	const ui = SpreadsheetApp.getUi();

	const monthResponse = ui.prompt(
		'Enter Month',
		'Enter month number (1-12):',
		ui.ButtonSet.OK_CANCEL
	);

	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	const yearResponse = ui.prompt(
		'Enter Year',
		'Enter year (e.g., 2025):',
		ui.ButtonSet.OK_CANCEL
	);

	if (yearResponse.getSelectedButton() !== ui.Button.OK) return;

	const month = parseInt(monthResponse.getResponseText()) - 1;
	const year = parseInt(yearResponse.getResponseText());

	if (isNaN(month) || month < 0 || month > 11 || isNaN(year)) {
		ui.alert('Invalid month or year');
		return;
	}

	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getActiveSheet();
	const sheetName = sheet.getName();

	Logger.log(`Syncing to active sheet: ${sheetName}`);
	syncLeaveDataToSheet(sheet, month, year);
}

/**
 * Sync leave only - applies leave colors/values without reformatting attendance hours
 */
function syncLeaveOnly() {
	const ui = SpreadsheetApp.getUi();
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

	const monthResponse = ui.prompt(
		'Sync Leave Only',
		'Enter month number (1-12):',
		ui.ButtonSet.OK_CANCEL
	);

	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	const yearResponse = ui.prompt(
		'Sync Leave Only',
		'Enter year (e.g., 2025):',
		ui.ButtonSet.OK_CANCEL
	);

	if (yearResponse.getSelectedButton() !== ui.Button.OK) return;

	const month = parseInt(monthResponse.getResponseText()) - 1;
	const year = parseInt(yearResponse.getResponseText());

	if (isNaN(month) || month < 0 || month > 11 || isNaN(year)) {
		ui.alert('Invalid month or year');
		return;
	}

	try {
		Logger.log(
			`Syncing leave only for ${
				month + 1
			}/${year} to active sheet (keeping hours)`
		);

		const token = getAccessToken();
		if (!token) {
			ui.alert('Failed to get API token. Check your credentials.');
			return;
		}

		const employees = fetchAllEmployees(token);
		if (!employees || employees.length === 0) {
			ui.alert('No employees found');
			return;
		}

		const leaveData = fetchLeaveDataForMonth(token, employees, month, year);
		if (!leaveData || Object.keys(leaveData).length === 0) {
			ui.alert('No leave data found for this month');
			return;
		}

		updateSheetWithLeaveData(sheet, leaveData, month, year, true);

		SpreadsheetApp.flush();
		ui.alert(
			`Leave synced successfully!\n\n` +
				`Processed ${Object.keys(leaveData).length} employees with leave.\n` +
				`Employee working hours were NOT changed.`
		);
	} catch (error) {
		Logger.log('Error syncing leave only: ' + error.message);
		ui.alert('Error: ' + error.message);
	}
}

/**
 * Main sync function for a specific month
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 */
function syncLeaveDataForMonth(month, year) {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheetName = getMonthSheetName(month, year);

	let sheet = ss.getSheetByName(sheetName);
	if (!sheet) {
		sheet = ss.insertSheet(sheetName);
		Logger.log(`Created new sheet: ${sheetName}`);
	} else {
		Logger.log(`Using existing sheet: ${sheetName}`);
	}

	syncLeaveDataToSheet(sheet, month, year);
}

/**
 * Sync leave data to a specific sheet
 * @param {Sheet} sheet - The sheet to sync to
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 */
function syncLeaveDataToSheet(sheet, month, year) {
	const sheetName = sheet.getName();

	Logger.log(
		`Syncing leave data for ${month + 1}/${year} to sheet "${sheetName}"`
	);

	try {
		// Apply attendance data first
		const attendanceList = getAttendanceData();
		if (attendanceList && attendanceList.length > 0) {
			Logger.log('Applying attendance data...');
			applyAttendanceHours(sheet, attendanceList, month, year);
			Logger.log('Attendance data applied');
		} else {
			Logger.log('No attendance sheet found, using default hours (8)');
		}

		// Get access token
		const token = getAccessToken();
		if (!token) {
			SpreadsheetApp.getUi().alert(
				'Failed to get API token. Check your credentials.'
			);
			return;
		}

		// Fetch all employees
		Logger.log('Fetching employees...');
		const employees = fetchAllEmployees(token);
		Logger.log(`Found ${employees.length} employees`);

		// Fetch leave data
		Logger.log('Fetching leave data...');
		const leaveData = fetchLeaveDataForMonth(token, employees, month, year);
		Logger.log(`Found ${Object.keys(leaveData).length} employees with leave`);

		// Apply leave data
		Logger.log('Applying leave data...');
		updateSheetWithLeaveData(sheet, leaveData, month, year);

		SpreadsheetApp.flush();
		Logger.log('Sync complete');

		SpreadsheetApp.getUi().alert(
			`Sync complete!\n\n` +
				`• Attendance rows: ${attendanceList ? attendanceList.length : 0}\n` +
				`• Leave requests: ${Object.keys(leaveData).length} employees\n` +
				`• Total employees processed: ${employees.length}`
		);
	} catch (error) {
		Logger.log('Error: ' + error.message);
		Logger.log('Stack: ' + error.stack);
		SpreadsheetApp.getUi().alert(
			'Error: ' + error.message + '\n\nCheck View > Execution Log for details.'
		);
	} finally {
		SpreadsheetApp.flush();
	}
}

/**
 * Apply leave colors only to the active sheet without syncing attendance
 */
function applyLeaveColorsOnly() {
	const ui = SpreadsheetApp.getUi();
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

	const monthResponse = ui.prompt(
		'Apply Leave Colors',
		'Enter month (1-12):',
		ui.ButtonSet.OK_CANCEL
	);
	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	const yearResponse = ui.prompt(
		'Apply Leave Colors',
		'Enter year (e.g., 2025):',
		ui.ButtonSet.OK_CANCEL
	);
	if (yearResponse.getSelectedButton() !== ui.Button.OK) return;

	const month = parseInt(monthResponse.getResponseText()) - 1;
	const year = parseInt(yearResponse.getResponseText());

	if (isNaN(month) || month < 0 || month > 11 || isNaN(year)) {
		ui.alert('Invalid month or year');
		return;
	}

	try {
		Logger.log(
			`Applying leave colors for ${month + 1}/${year} to active sheet`
		);

		const token = getAccessToken();
		if (!token) {
			ui.alert('Failed to get API token. Check your credentials.');
			return;
		}

		const employees = fetchAllEmployees(token);
		const leaveData = fetchLeaveDataForMonth(token, employees, month, year);

		if (!leaveData || Object.keys(leaveData).length === 0) {
			ui.alert('No leave data found for this month');
			return;
		}

		applyLeaveColorsToSheet(sheet, leaveData, month, year);

		SpreadsheetApp.flush();
		ui.alert(
			`Leave colors applied successfully!\n\nProcessed ${
				Object.keys(leaveData).length
			} employees with leave.`
		);
	} catch (error) {
		Logger.log('Error applying leave colors: ' + error.message);
		ui.alert('Error: ' + error.message);
	} finally {
		SpreadsheetApp.flush();
	}
}
