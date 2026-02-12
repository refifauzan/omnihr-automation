/**
 * OmniHR Leave Data Integration for Google Sheets
 *
 * WHY: Automates leave data synchronization to eliminate manual data entry
 * and reduce errors in timesheet management. This integration ensures that
 * employee availability is always accurate for project planning and payroll.
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
 * WHY: Quick sync for current month without user input
 * This function exists because managers frequently need to update the current
 * month's data and shouldn't have to manually enter dates each time.
 */
function syncCurrentMonth() {
	const now = new Date();
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getActiveSheet();

	Logger.log(`Syncing current month to active sheet: ${sheet.getName()}`);
	syncLeaveDataToSheet(sheet, now.getMonth(), now.getFullYear());
}

/**
 * WHY: Flexible sync for any month/year combination
 * This function provides flexibility for historical data updates, future planning,
 * and correcting errors in specific months without affecting other periods.
 */
function syncLeaveData() {
	const ui = SpreadsheetApp.getUi();

	const monthResponse = ui.prompt(
		'Enter Month',
		'Enter month number (1-12):',
		ui.ButtonSet.OK_CANCEL,
	);

	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	const yearResponse = ui.prompt(
		'Enter Year',
		'Enter year (e.g., 2025):',
		ui.ButtonSet.OK_CANCEL,
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
			`• Daily sync: Enabled for ${month + 1}/${year} at 6 AM`,
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
		}/${year} to sheet "${sheetName}"`,
	);

	try {
		Logger.log(
			`Syncing leave only for ${
				month + 1
			}/${year} to active sheet (keeping hours)`,
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

		// Fetch holidays to exclude from Total Days Off
		const holidays = fetchHolidaysForMonth(token, month, year);
		const holidayDays = new Set(holidays.map((h) => h.date));

		updateSheetWithLeaveData(sheet, leaveData, month, year, true, holidayDays);

		// Apply grey-out for employees who joined/left mid-month
		Logger.log('Applying grey-out for hire/termination dates...');
		applyEmployeeDateGreyOut(month, year);

		SpreadsheetApp.flush();

		// Auto-enable protection and daily sync
		installOnEditTriggerSilent();
		setupDailyLeaveOnlyTrigger(month, year, sheetName);

		ui.alert(
			`Leave synced successfully!\n\n` +
				`Processed ${Object.keys(leaveData).length} employees with leave.\n` +
				`Employee working hours were NOT changed.\n\n` +
				`• Edit protection: Enabled\n` +
				`• Daily sync: Enabled for ${month + 1}/${year} at 6 AM`,
		);
	} catch (error) {
		Logger.log('Error syncing leave only: ' + error.message);
		ui.alert('Error: ' + error.message);
	}
}

/**
 * Populate default hours for empty cells while preserving existing values and leave
 * This function fills empty working day cells with 8 hours for default teams
 */
function populateDefaultHours() {
	const ui = SpreadsheetApp.getUi();
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getActiveSheet();
	const sheetName = sheet.getName();

	const monthResponse = ui.prompt(
		'Populate Default Hours',
		'Enter month (1-12):',
		ui.ButtonSet.OK_CANCEL,
	);

	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	const yearResponse = ui.prompt(
		'Populate Default Hours',
		'Enter year (e.g., 2025):',
		ui.ButtonSet.OK_CANCEL,
	);

	if (yearResponse.getSelectedButton() !== ui.Button.OK) return;

	const month = parseInt(monthResponse.getResponseText()) - 1;
	const year = parseInt(yearResponse.getResponseText());

	if (isNaN(month) || month < 0 || month > 11 || isNaN(year)) {
		ui.alert('Invalid month or year');
		return;
	}

	Logger.log(
		`Populating default hours for ${month + 1}/${year} to sheet "${sheetName}"`,
	);

	try {
		// Fetch holidays to skip them
		let holidays = [];
		try {
			const token = getAccessToken();
			if (token) {
				holidays = fetchHolidaysForMonth(token, month, year);
				Logger.log(`Found ${holidays.length} holidays`);
			}
		} catch (e) {
			Logger.log('Could not fetch holidays: ' + e.message);
		}
		const holidayDays = new Set(holidays.map((h) => h.date));

		// Call the existing function to set default hours
		const { dayColumns } = calculateDayColumns(month, year);
		setOperationsDefaultHours(sheet, dayColumns, month, year, holidayDays);

		SpreadsheetApp.flush();

		ui.alert(
			`Default hours populated successfully!\n\n` +
				`Empty cells for default teams (${CONFIG.DEFAULT_HOUR_TEAMS.join(', ')}) ` +
				`have been set to ${CONFIG.DEFAULT_HOURS} hours.\n\n` +
				`Existing values and leave markings were preserved.`,
		);
	} catch (error) {
		Logger.log('Error populating default hours: ' + error.message);
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
		`Syncing leave data for ${month + 1}/${year} to sheet "${sheetName}"`,
	);

	try {
		// Get access token
		const token = getAccessToken();
		if (!token) {
			SpreadsheetApp.getUi().alert(
				'Failed to get API token. Check your credentials.',
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

		// Fetch holidays to exclude from Total Days Off
		Logger.log('Fetching holidays...');
		const holidays = fetchHolidaysForMonth(token, month, year);
		const holidayDays = new Set(holidays.map((h) => h.date));
		Logger.log(`Found ${holidays.length} public holidays`);

		// Apply leave data
		Logger.log('Applying leave data...');
		updateSheetWithLeaveData(sheet, leaveData, month, year, true, holidayDays);

		// Apply grey-out for employees who joined/left mid-month
		Logger.log('Applying grey-out for hire/termination dates...');
		applyEmployeeDateGreyOut(month, year);

		SpreadsheetApp.flush();
		Logger.log('Sync complete');

		SpreadsheetApp.getUi().alert(
			`Sync complete!\n\n` +
				`• Leave requests: ${Object.keys(leaveData).length} employees\n` +
				`• Total employees processed: ${employees.length}`,
		);
	} catch (error) {
		Logger.log('Error: ' + error.message);
		Logger.log('Stack: ' + error.stack);
		SpreadsheetApp.getUi().alert(
			'Error: ' + error.message + '\n\nCheck View > Execution Log for details.',
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
		ui.ButtonSet.OK_CANCEL,
	);
	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	const yearResponse = ui.prompt(
		'Create Empty Table',
		'Enter year (e.g., 2025):',
		ui.ButtonSet.OK_CANCEL,
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

		// Fetch holidays from OmniHR API
		let holidays = [];
		try {
			const token = getAccessToken();
			if (token) {
				holidays = fetchHolidaysForMonth(token, month, year);
				Logger.log(
					`Found ${holidays.length} public holidays for ${month + 1}/${year}`,
				);
			}
		} catch (e) {
			Logger.log('Could not fetch holidays: ' + e.message);
			// Continue without holidays - not a critical error
		}

		// Build a set of holiday day numbers for quick lookup
		const holidayDays = new Set(holidays.map((h) => h.date));

		// Calculate day columns
		const { dayColumns, validatedColumns, weekOverrideColumns } =
			calculateDayColumns(month, year);
		const lastDayCol = Math.max(
			...Object.values(dayColumns),
			...validatedColumns,
			...weekOverrideColumns,
		);
		const totalCols = lastDayCol - CONFIG.FIRST_DAY_COL + 1;

		// Get number of employee rows
		const lastRow = sheet.getLastRow();
		const numRows = Math.max(lastRow - CONFIG.FIRST_DATA_ROW + 1, 0);

		if (numRows === 0) {
			ui.alert(
				'No employee data found. Please ensure employee data exists in columns A-D.',
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
				(d) => dayColumns[d] === col,
			);
			if (dayForCol) {
				const date = new Date(year, month, parseInt(dayForCol));
				const dayOfWeek = date.getDay();
				headerRow1.push(dayNames[dayOfWeek]);
				headerRow2.push(parseInt(dayForCol));
				headerBgColors.push(
					dayOfWeek >= 1 && dayOfWeek <= 5 ? '#356854' : '#efefef',
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
			totalCols,
		);
		headerRange2.setBackgrounds([headerBgColors]);

		const fontColors = headerBgColors.map((bg) =>
			bg === '#356854' ? '#FFFFFF' : '#000000',
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

		// Get team data from column C to check for Operations team
		const teamData = sheet
			.getRange(CONFIG.FIRST_DATA_ROW, CONFIG.PROJECT_COL, numRows, 1)
			.getValues()
			.map((row) => row[0]);

		// Fill data with 0 for hours (8 for Operations) and FALSE for validated
		const dataValues = [];
		const dataBackgrounds = [];

		for (let i = 0; i < numRows; i++) {
			const rowValues = [];
			const rowBackgrounds = [];
			const team = teamData[i] ? teamData[i].toString().toLowerCase() : '';
			const hasDefaultHours = CONFIG.DEFAULT_HOUR_TEAMS.some((t) =>
				team.includes(t),
			);

			for (let col = CONFIG.FIRST_DAY_COL; col <= lastDayCol; col++) {
				const dayForCol = Object.keys(dayColumns).find(
					(d) => dayColumns[d] === col,
				);

				if (dayForCol) {
					const date = new Date(year, month, parseInt(dayForCol));
					const dayOfWeek = date.getDay();

					const dayNum = parseInt(dayForCol);
					const isHoliday = holidayDays.has(dayNum);

					if (dayOfWeek === 0 || dayOfWeek === 6) {
						// Sunday or Saturday
						rowValues.push(''); // Empty value
						rowBackgrounds.push('#efefef'); // Gray background
					} else if (isHoliday) {
						// Public holiday - treat like weekend but with pastel red
						rowValues.push(''); // Empty value (like weekend)
						rowBackgrounds.push('#FFCCCB'); // Pastel red background
					} else {
						// Weekday - 8 hours for default teams, 0 for others
						rowValues.push(hasDefaultHours ? CONFIG.DEFAULT_HOURS : 0);
						rowBackgrounds.push(null); // Default background
					}
				} else if (validatedColumns.includes(col)) {
					rowValues.push(false); // FALSE for validated
					rowBackgrounds.push(null);
				} else if (weekOverrideColumns.includes(col)) {
					rowValues.push(false); // FALSE for override
					rowBackgrounds.push(null);
				} else {
					rowValues.push(0);
					rowBackgrounds.push(null);
				}
			}
			dataValues.push(rowValues);
			dataBackgrounds.push(rowBackgrounds);
		}

		// Apply data
		const dataRange = sheet.getRange(
			CONFIG.FIRST_DATA_ROW,
			CONFIG.FIRST_DAY_COL,
			numRows,
			totalCols,
		);
		dataRange.setValues(dataValues);
		dataRange.setBackgrounds(dataBackgrounds);

		// Setup checkboxes
		for (const col of validatedColumns) {
			const checkboxRange = sheet.getRange(
				CONFIG.FIRST_DATA_ROW,
				col,
				numRows,
				1,
			);
			checkboxRange.insertCheckboxes();
		}

		for (const col of weekOverrideColumns) {
			const checkboxRange = sheet.getRange(
				CONFIG.FIRST_DATA_ROW,
				col,
				numRows,
				1,
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
				'', // Column I will be updated by leave sync with actual working days off
			]);
		}

		sheet.getRange(CONFIG.FIRST_DATA_ROW, 7, numRows, 3).setFormulas(formulas);

		// Add conditional formatting (inactive since validated is FALSE)
		addValidatedConditionalFormatting(sheet, month, year);

		SpreadsheetApp.flush();

		const holidayInfo =
			holidays.length > 0
				? `• Public holidays: ${holidays.length} (marked in red)\n`
				: '';

		ui.alert(
			`Empty table created successfully!\n\n` +
				`• Month: ${month + 1}/${year}\n` +
				`• Days: ${daysInMonth}\n` +
				`• Employee rows: ${numRows}\n` +
				holidayInfo +
				`• All hours set to 0\n` +
				`• All Validated checkboxes set to FALSE`,
		);
	} catch (error) {
		Logger.log('Error creating empty table: ' + error.message);
		ui.alert('Error: ' + error.message);
	}
}

/**
 * Sync holidays from OmniHR - refreshes holiday formatting on the active sheet
 * Applies pastel red background to holiday columns for all employee rows
 */
function syncHolidays() {
	const ui = SpreadsheetApp.getUi();
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

	const monthResponse = ui.prompt(
		'Sync Holidays',
		'Enter month (1-12):',
		ui.ButtonSet.OK_CANCEL,
	);
	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	const yearResponse = ui.prompt(
		'Sync Holidays',
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

	try {
		Logger.log(`Syncing holidays for ${month + 1}/${year} to active sheet`);

		const token = getAccessToken();
		if (!token) {
			ui.alert('Failed to get API token. Check your credentials.');
			return;
		}

		// Fetch holidays from OmniHR
		const holidays = fetchHolidaysForMonth(token, month, year);
		Logger.log(`Found ${holidays.length} public holidays from OmniHR`);

		if (holidays.length === 0) {
			ui.alert(`No public holidays found for ${month + 1}/${year}`);
			return;
		}

		// Log each holiday for verification
		for (const holiday of holidays) {
			Logger.log(`Holiday: Day ${holiday.date} - ${holiday.name}`);
		}

		// Apply holiday formatting to the sheet
		applyHolidayFormatting(sheet, holidays, month, year);

		SpreadsheetApp.flush();

		const holidayList = holidays
			.map((h) => `• Day ${h.date}: ${h.name}`)
			.join('\n');

		ui.alert(
			`Holidays synced successfully!\n\n` +
				`Found ${holidays.length} public holidays for ${
					month + 1
				}/${year}:\n\n` +
				holidayList,
		);
	} catch (error) {
		Logger.log('Error syncing holidays: ' + error.message);
		ui.alert('Error: ' + error.message);
	}
}

/**
 * Test function to verify holidays from OmniHR API
 * Run this to check what holidays are returned for a specific month
 */
function testFetchHolidays() {
	const ui = SpreadsheetApp.getUi();

	const monthResponse = ui.prompt(
		'Test Fetch Holidays',
		'Enter month (1-12):',
		ui.ButtonSet.OK_CANCEL,
	);
	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	const yearResponse = ui.prompt(
		'Test Fetch Holidays',
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

	try {
		const token = getAccessToken();
		if (!token) {
			ui.alert('Failed to get API token. Check your credentials.');
			return;
		}

		Logger.log(`Testing holiday fetch for ${month + 1}/${year}`);
		const holidays = fetchHolidaysForMonth(token, month, year);

		if (holidays.length === 0) {
			ui.alert(`No holidays found for ${month + 1}/${year}`);
			return;
		}

		let message = `Holidays for ${month + 1}/${year}:\n\n`;
		for (const h of holidays) {
			message += `• Day ${h.date}: ${h.name}\n`;
			Logger.log(`Holiday: Day ${h.date} - ${h.name}`);
		}

		ui.alert(message);
	} catch (error) {
		Logger.log('Error testing holidays: ' + error.message);
		ui.alert('Error: ' + error.message);
	}
}

/**
 * Apply holiday formatting to the sheet
 * Sets pastel red background and clears values for holiday columns
 * @param {Sheet} sheet - The sheet
 * @param {Array} holidays - Array of { date: dayNumber, name: holidayName }
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 */
function applyHolidayFormatting(sheet, holidays, month, year) {
	const { dayColumns } = calculateDayColumns(month, year);
	const holidayDays = new Set(holidays.map((h) => h.date));

	const lastRow = sheet.getLastRow();
	const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;

	if (numRows <= 0) {
		Logger.log('No data rows to format');
		return;
	}

	const holidayColor = '#FFCCCB'; // Pastel red
	let formattedCells = 0;

	for (const [dayStr, col] of Object.entries(dayColumns)) {
		const dayNum = parseInt(dayStr);

		if (!holidayDays.has(dayNum)) continue;

		// Check if it's a weekday (holidays on weekends are already grey)
		const date = new Date(year, month, dayNum);
		const dayOfWeek = date.getDay();

		if (dayOfWeek === 0 || dayOfWeek === 6) {
			Logger.log(`Day ${dayNum} is a weekend, skipping holiday formatting`);
			continue;
		}

		// Apply holiday formatting to all rows in this column
		const range = sheet.getRange(CONFIG.FIRST_DATA_ROW, col, numRows, 1);
		range.setBackground(holidayColor);
		range.setValue(''); // Clear values (no work on holidays)

		formattedCells += numRows;
		Logger.log(`Applied holiday formatting to day ${dayNum}, column ${col}`);
	}

	Logger.log(`Formatted ${formattedCells} cells as holidays`);
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
		ui.ButtonSet.OK_CANCEL,
	);
	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	const yearResponse = ui.prompt(
		'Apply Leave Colors',
		'Enter year (e.g., 2025):',
		ui.ButtonSet.OK_CANCEL,
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
			`Applying leave colors for ${month + 1}/${year} to active sheet`,
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
			} employees with leave.`,
		);
	} catch (error) {
		Logger.log('Error applying leave colors: ' + error.message);
		ui.alert('Error: ' + error.message);
	} finally {
		SpreadsheetApp.flush();
	}
}
