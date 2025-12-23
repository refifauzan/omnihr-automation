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

	// Auto-enable protection and daily sync
	installOnEditTriggerSilent();
	setupDailyLeaveOnlyTrigger(month, year, sheetName);

	ui.alert(
		`Sync complete!\n\n` +
			`• Edit protection: Enabled\n` +
			`• Daily sync: Enabled for ${month + 1}/${year} at 6 AM`
	);
}

/**
 * Sync leave only - applies leave colors/values for current month
 */
function syncLeaveOnly() {
	const ui = SpreadsheetApp.getUi();
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getActiveSheet();
	const sheetName = sheet.getName();

	// Use current month/year automatically
	const now = new Date();
	const month = now.getMonth();
	const year = now.getFullYear();

	Logger.log(
		`Syncing leave only for current month ${
			month + 1
		}/${year} to sheet "${sheetName}"`
	);

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

		// Auto-enable protection and daily sync
		installOnEditTriggerSilent();
		setupDailyLeaveOnlyTrigger(month, year, sheetName);

		ui.alert(
			`Leave synced successfully!\n\n` +
				`Processed ${Object.keys(leaveData).length} employees with leave.\n` +
				`Employee working hours were NOT changed.\n\n` +
				`• Edit protection: Enabled\n` +
				`• Daily sync: Enabled for ${month + 1}/${year} at 6 AM`
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
		updateSheetWithLeaveData(sheet, leaveData, month, year, true);

		SpreadsheetApp.flush();
		Logger.log('Sync complete');

		SpreadsheetApp.getUi().alert(
			`Sync complete!\n\n` +
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
 * Create empty table structure from column K onwards
 * Sets all hours to 0 and Validated checkboxes to FALSE
 */
function createEmptyTableStructure() {
	const ui = SpreadsheetApp.getUi();
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

	const monthResponse = ui.prompt(
		'Create Empty Table',
		'Enter month (1-12):',
		ui.ButtonSet.OK_CANCEL
	);
	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	const yearResponse = ui.prompt(
		'Create Empty Table',
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
		const daysInMonth = new Date(year, month + 1, 0).getDate();
		const dayNames = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];

		// Calculate day columns
		const { dayColumns, validatedColumns, weekOverrideColumns } =
			calculateDayColumns(month, year);
		const lastDayCol = Math.max(
			...Object.values(dayColumns),
			...validatedColumns,
			...weekOverrideColumns
		);
		const totalCols = lastDayCol - CONFIG.FIRST_DAY_COL + 1;

		// Get number of employee rows
		const lastRow = sheet.getLastRow();
		const numRows = Math.max(lastRow - CONFIG.FIRST_DATA_ROW + 1, 0);

		if (numRows === 0) {
			ui.alert(
				'No employee data found. Please ensure employee data exists in columns A-D.'
			);
			return;
		}

		// Delete existing columns from K onwards
		if (sheet.getLastColumn() >= CONFIG.FIRST_DAY_COL) {
			const numColsToDelete = sheet.getLastColumn() - CONFIG.FIRST_DAY_COL + 1;
			sheet.deleteColumns(CONFIG.FIRST_DAY_COL, numColsToDelete);
		}

		// Insert new columns
		sheet.insertColumnsAfter(CONFIG.FIRST_DAY_COL - 1, totalCols);

		// Build headers
		const headerRow1 = [];
		const headerRow2 = [];
		const headerBgColors = [];

		for (let col = CONFIG.FIRST_DAY_COL; col <= lastDayCol; col++) {
			const dayForCol = Object.keys(dayColumns).find(
				(d) => dayColumns[d] === col
			);
			if (dayForCol) {
				const date = new Date(year, month, parseInt(dayForCol));
				const dayOfWeek = date.getDay();
				headerRow1.push(dayNames[dayOfWeek]);
				headerRow2.push(parseInt(dayForCol));
				headerBgColors.push(
					dayOfWeek >= 1 && dayOfWeek <= 5 ? '#356854' : '#efefef'
				);
			} else if (validatedColumns.includes(col)) {
				headerRow1.push('');
				headerRow2.push('Validated');
				headerBgColors.push('#356854');
			} else if (weekOverrideColumns.includes(col)) {
				headerRow1.push('');
				headerRow2.push('Time off Override');
				headerBgColors.push('#efefef');
			} else {
				headerRow1.push('');
				headerRow2.push('');
				headerBgColors.push(null);
			}
		}

		// Apply headers
		sheet
			.getRange(CONFIG.DAY_NAME_ROW, CONFIG.FIRST_DAY_COL, 1, totalCols)
			.setValues([headerRow1]);
		sheet
			.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_DAY_COL, 1, totalCols)
			.setValues([headerRow2]);

		const headerRange2 = sheet.getRange(
			CONFIG.HEADER_ROW,
			CONFIG.FIRST_DAY_COL,
			1,
			totalCols
		);
		headerRange2.setBackgrounds([headerBgColors]);

		const fontColors = headerBgColors.map((bg) =>
			bg === '#356854' ? '#FFFFFF' : '#000000'
		);
		headerRange2.setFontColors([fontColors]);

		// Set column widths
		for (let col = CONFIG.FIRST_DAY_COL; col <= lastDayCol; col++) {
			if (validatedColumns.includes(col) || weekOverrideColumns.includes(col)) {
				sheet.setColumnWidth(col, 105);
			} else {
				sheet.setColumnWidth(col, 46);
			}
		}

		// Fill data with 0 for hours and FALSE for validated
		const dataValues = [];
		for (let i = 0; i < numRows; i++) {
			const rowValues = [];
			for (let col = CONFIG.FIRST_DAY_COL; col <= lastDayCol; col++) {
				if (validatedColumns.includes(col)) {
					rowValues.push(false); // FALSE for validated
				} else if (weekOverrideColumns.includes(col)) {
					rowValues.push(false); // FALSE for override
				} else {
					rowValues.push(0); // 0 for hours
				}
			}
			dataValues.push(rowValues);
		}

		// Apply data
		const dataRange = sheet.getRange(
			CONFIG.FIRST_DATA_ROW,
			CONFIG.FIRST_DAY_COL,
			numRows,
			totalCols
		);
		dataRange.setValues(dataValues);

		// Setup checkboxes
		for (const col of validatedColumns) {
			const checkboxRange = sheet.getRange(
				CONFIG.FIRST_DATA_ROW,
				col,
				numRows,
				1
			);
			checkboxRange.insertCheckboxes();
		}

		for (const col of weekOverrideColumns) {
			const checkboxRange = sheet.getRange(
				CONFIG.FIRST_DATA_ROW,
				col,
				numRows,
				1
			);
			checkboxRange.insertCheckboxes();
		}

		// Apply formulas for totals
		const formulas = [];
		for (let i = 0; i < numRows; i++) {
			const row = CONFIG.FIRST_DATA_ROW + i;
			const firstDayColLetter = columnToLetter(CONFIG.FIRST_DAY_COL);
			const lastDayColLetter = columnToLetter(lastDayCol);
			const rangeStr = `${firstDayColLetter}${row}:${lastDayColLetter}${row}`;
			formulas.push([
				`=SUM(${rangeStr})`,
				`=G${row}/8`,
				`=COUNTIF(${rangeStr},"=0")`,
			]);
		}

		sheet.getRange(CONFIG.FIRST_DATA_ROW, 7, numRows, 3).setFormulas(formulas);

		SpreadsheetApp.flush();

		ui.alert(
			`Empty table created successfully!\n\n` +
				`• Month: ${month + 1}/${year}\n` +
				`• Days: ${daysInMonth}\n` +
				`• Employee rows: ${numRows}\n` +
				`• All hours set to 0\n` +
				`• All Validated checkboxes set to FALSE`
		);
	} catch (error) {
		Logger.log('Error creating empty table: ' + error.message);
		ui.alert('Error: ' + error.message);
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
