/**
 * OmniHR Leave Data Integration for Google Sheets
 *
 * This script fetches leave data from OmniHR API and updates the spreadsheet
 * with leave information, matching the template format.
 *
 * Setup:
 * 1. Open your Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Paste this code
 * 4. Click OmniHR > Setup API Credentials and enter:
 *    - Base URL (e.g., https://api.omnihr.co/api/v1)
 *    - Subdomain
 *    - Username
 *    - Password
 * 5. Run setupTrigger() once to enable automatic updates
 */

const CONFIG = {
	// Sheet configuration
	DAY_NAME_ROW: 1, // Row 1 containing day names (S, M, T, W, T, F, S)
	HEADER_ROW: 2, // Row 2 containing day numbers
	FIRST_DATA_ROW: 3, // Row 3 = First row with employee data
	EMPLOYEE_ID_COL: 1, // Column A = Employee ID
	EMPLOYEE_NAME_COL: 2, // Column B = Employee Name
	PROJECT_COL: 3, // Column C = Project (for matching with attendance)
	FIRST_DAY_COL: 11, // Column K = Day 1 (same as Excel template)

	// Attendance sheet configuration
	ATTENDANCE_SHEET_NAME: 'Attendance',
	ATTENDANCE_ID_COL: 1, // Column A = Employee ID
	ATTENDANCE_NAME_COL: 2, // Column B = Employee Name
	ATTENDANCE_PROJECT_COL: 3, // Column C = Project
	ATTENDANCE_TYPE_COL: 4, // Column D = Type (Fulltime, Parttime, Custom)
	ATTENDANCE_HOURS_COL: 5, // Column E = Daily Hours
	ATTENDANCE_FIRST_DATA_ROW: 2, // First row with data (after header)

	// Colors (in hex without #)
	COLORS: {
		FULL_DAY: '#FF0000', // Red for full day leave
		HALF_DAY: '#FFA500', // Orange for half day leave
		WEEKEND: '#D3D3D3', // Light grey for weekend
	},

	// Hours
	DEFAULT_HOURS: 8,
	HALF_DAY_HOURS: 4,
};

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
	try {
		const ui = SpreadsheetApp.getUi();
		ui.createMenu('OmniHR')
			.addItem('Sync Leave Data', 'syncLeaveData')
			.addItem('Sync Current Month', 'syncCurrentMonth')
			.addSeparator()
			.addSubMenu(
				ui
					.createMenu('Attendance')
					.addItem('Create Attendance Sheet', 'createAttendanceSheet')
					.addItem(
						'Apply Attendance to Current Sheet',
						'applyAttendanceToCurrentSheet'
					)
			)
			.addSubMenu(
				ui
					.createMenu('Schedule')
					.addItem('Enable Monthly Sync (1st of month)', 'setupMonthlyTrigger')
					.addItem('Enable Daily Sync', 'setupDailyTrigger')
					.addSeparator()
					.addItem('View Current Schedule', 'viewTriggers')
					.addItem('Disable Automation', 'removeTriggers')
			)
			.addSeparator()
			.addItem('Setup API Credentials', 'showCredentialsDialog')
			.addToUi();
	} catch (e) {
		// UI not available (e.g., running from trigger or API context)
		Logger.log('onOpen: UI not available - ' + e.message);
	}
}

/**
 * Setup monthly trigger - runs on 1st of each month at 6 AM
 * Run this function ONCE to enable automatic monthly sync
 */
function setupMonthlyTrigger() {
	// Remove existing triggers
	const triggers = ScriptApp.getProjectTriggers();
	triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));

	// Create monthly trigger - runs on day 1 of each month at 6 AM
	ScriptApp.newTrigger('syncCurrentMonth')
		.timeBased()
		.onMonthDay(1)
		.atHour(6)
		.create();

	Logger.log(
		'Monthly sync trigger created - will run on 1st of each month at 6 AM'
	);
	SpreadsheetApp.getUi().alert(
		'Monthly sync enabled!\n\nThe script will automatically sync on the 1st of each month at 6 AM.'
	);
}

/**
 * Setup daily trigger - runs every day at 6 AM
 * Only syncs the CURRENT month - past months are not re-synced
 * When the month changes, it automatically starts syncing the new month
 */
function setupDailyTrigger() {
	// Remove existing triggers
	const triggers = ScriptApp.getProjectTriggers();
	triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));

	// Create daily trigger at 6 AM
	ScriptApp.newTrigger('syncCurrentMonth')
		.timeBased()
		.everyDays(1)
		.atHour(6)
		.create();

	Logger.log('Daily sync trigger created for 6 AM');
	SpreadsheetApp.getUi().alert(
		'Daily sync enabled!\n\n' +
			'The script will automatically sync every day at 6 AM.\n\n' +
			'• Only the CURRENT month is synced\n' +
			'• Past months will NOT be re-synced\n' +
			'• When the month changes, the new month is synced automatically'
	);
}

/**
 * Remove all triggers (disable automation)
 */
function removeTriggers() {
	const triggers = ScriptApp.getProjectTriggers();
	triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));
	Logger.log('All triggers removed');
	SpreadsheetApp.getUi().alert(
		'Automation disabled.\n\nAll scheduled syncs have been removed.'
	);
}

/**
 * View current triggers
 */
function viewTriggers() {
	const triggers = ScriptApp.getProjectTriggers();
	if (triggers.length === 0) {
		SpreadsheetApp.getUi().alert(
			'No scheduled syncs.\n\nUse OmniHR > Schedule to set up automation.'
		);
		return;
	}

	let info = 'Current scheduled syncs:\n\n';
	triggers.forEach((trigger, i) => {
		info += `${
			i + 1
		}. ${trigger.getHandlerFunction()} - ${trigger.getEventType()}\n`;
	});
	SpreadsheetApp.getUi().alert(info);
}

/**
 * Get the sheet name for a specific month
 * Format: "January 2025", "February 2025", etc.
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @returns {string} Sheet name
 */
function getMonthSheetName(month, year) {
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
	return `${monthNames[month]} ${year}`;
}

/**
 * Create the Attendance sheet and fetch all employees from OmniHR
 */
function createAttendanceSheet() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const ui = SpreadsheetApp.getUi();

	// Check if sheet already exists
	let sheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME);
	if (sheet) {
		const response = ui.alert(
			'Attendance sheet already exists!',
			'Do you want to refresh employee data from OmniHR?',
			ui.ButtonSet.YES_NO
		);
		if (response !== ui.Button.YES) return;

		// Clear existing data but keep headers
		const lastRow = sheet.getLastRow();
		if (lastRow > 1) {
			sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
		}
	} else {
		// Create new sheet
		sheet = ss.insertSheet(CONFIG.ATTENDANCE_SHEET_NAME);

		// Set headers
		const headers = ['ID', 'Employee Name', 'Project', 'Type', 'Daily Hours'];
		sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
		sheet
			.getRange(1, 1, 1, headers.length)
			.setFontWeight('bold')
			.setBackground('#4285f4')
			.setFontColor('#FFFFFF');

		// Add data validation for Type column (column D)
		const typeRule = SpreadsheetApp.newDataValidation()
			.requireValueInList(['Full-time', 'Part-time', 'Custom'], true)
			.setAllowInvalid(false)
			.build();
		sheet.getRange('D2:D1000').setDataValidation(typeRule);

		// Add instructions
		sheet.getRange('G1').setValue('Instructions:');
		sheet
			.getRange('G2')
			.setValue('1. Employee ID and Name are fetched from OmniHR');
		sheet.getRange('G3').setValue('2. Enter Project name in column C');
		sheet
			.getRange('G4')
			.setValue('3. Select Type: Full-time (8h), Part-time (4h), or Custom');
		sheet
			.getRange('G5')
			.setValue('4. For Custom, enter the daily hours in column E');
		sheet.getRange('G6').setValue('5. Sync will automatically use these hours');
		sheet.getRange('G1:G6').setFontStyle('italic').setFontColor('#666666');
	}

	// Fetch employees from OmniHR
	try {
		ui.alert('Fetching employees from OmniHR...\n\nThis may take a moment.');

		const token = getAccessToken();
		if (!token) {
			ui.alert(
				'Failed to get API token. Please setup credentials first.\n\nUse OmniHR > Setup API Credentials'
			);
			return;
		}

		Logger.log('Fetching employees for attendance sheet...');
		const employees = fetchAllEmployees(token);
		Logger.log(`Found ${employees.length} employees`);

		// Fetch employee IDs (SM0068 format) in batches
		const employeeData = [];
		const BATCH_SIZE = 50;

		for (let i = 0; i < employees.length; i += BATCH_SIZE) {
			const batch = employees.slice(i, i + BATCH_SIZE);
			Logger.log(
				`Fetching base data batch ${Math.floor(i / BATCH_SIZE) + 1}/${Math.ceil(
					employees.length / BATCH_SIZE
				)}`
			);

			const requests = buildBaseDataRequests(token, batch);
			const responses = UrlFetchApp.fetchAll(requests.requests);

			for (let j = 0; j < responses.length; j++) {
				try {
					const data = JSON.parse(responses[j].getContentText());
					const baseData = data.data || data;
					const emp = batch[j];

					employeeData.push([
						baseData.employee_id || '',
						emp.full_name || emp.name || `User ${emp.id}`,
						'', // Project (user fills in)
						'Full-time', // Default to full-time
						8, // Default 8 hours
					]);
				} catch (e) {
					const emp = batch[j];
					employeeData.push([
						'',
						emp.full_name || emp.name || `User ${emp.id}`,
						'', // Project
						'Full-time',
						8,
					]);
				}
			}
		}

		// Write employee data to sheet
		if (employeeData.length > 0) {
			sheet.getRange(2, 1, employeeData.length, 5).setValues(employeeData);
		}

		// Auto-resize columns
		sheet.autoResizeColumns(1, 5);

		SpreadsheetApp.flush();

		ui.alert(
			`Attendance sheet ready!\n\n` +
				`• ${employeeData.length} employees loaded from OmniHR\n` +
				`• Default: Full-time (8 hours)\n\n` +
				`Edit the Type column to set Part-time (4h) or Custom hours.`
		);
	} catch (error) {
		Logger.log('Error creating attendance sheet: ' + error.message);
		ui.alert('Error: ' + error.message);
	}
}

/**
 * Build batch requests for fetching employee base data only
 */
function buildBaseDataRequests(token, employees) {
	const props = PropertiesService.getScriptProperties();
	const baseUrl = props.getProperty('OMNIHR_BASE_URL');
	const subdomain = props.getProperty('OMNIHR_SUBDOMAIN');

	const headers = {
		Authorization: `Bearer ${token}`,
		'x-subdomain': subdomain,
		'Content-Type': 'application/json',
	};

	const requests = [];

	for (const emp of employees) {
		const userId = emp.id || emp.user_id;
		requests.push({
			url: `${baseUrl}/employee/2.0/users/${userId}/base-data/`,
			method: 'get',
			headers: headers,
			muteHttpExceptions: true,
		});
	}

	return { requests };
}

/**
 * Get attendance data as an array (preserves row order)
 * Each row maps: { empId, empName, project, hours }
 * Used to match main sheet rows by Employee ID + Project
 */
function getAttendanceData() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME);

	if (!sheet) {
		return null;
	}

	const lastRow = sheet.getLastRow();
	if (lastRow < CONFIG.ATTENDANCE_FIRST_DATA_ROW) {
		return [];
	}

	const data = sheet
		.getRange(
			CONFIG.ATTENDANCE_FIRST_DATA_ROW,
			1,
			lastRow - CONFIG.ATTENDANCE_FIRST_DATA_ROW + 1,
			5 // 5 columns: ID, Name, Project, Type, Hours
		)
		.getValues();

	const attendanceList = [];

	for (let i = 0; i < data.length; i++) {
		const row = data[i];
		const empId = String(row[0]).trim().toUpperCase();
		const empName = String(row[1]).trim();
		const project = String(row[2]).trim();
		const type = String(row[3]).trim();
		const rawHours = parseFloat(row[4]);
		// Use raw hours if it's a valid number (including 0), otherwise use default
		let hours = !isNaN(rawHours) ? rawHours : CONFIG.DEFAULT_HOURS;

		if (!empId && !empName) continue;

		// Determine hours based on type (only if not Custom)
		if (type === 'Full-time') {
			hours = 8;
		} else if (type === 'Part-time') {
			hours = 4;
		}

		attendanceList.push({
			empId,
			empName,
			project,
			hours,
			type,
		});

		Logger.log(
			`Attendance: ${empId} - ${empName} (${project}) = ${hours} hours`
		);
	}

	Logger.log(`Loaded ${attendanceList.length} attendance rows`);
	return attendanceList;
}

/**
 * Apply attendance data to the current sheet (fill base hours for weekdays)
 */
function applyAttendanceToCurrentSheet() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getActiveSheet();

	// Get attendance data
	const attendanceList = getAttendanceData();
	if (!attendanceList) {
		SpreadsheetApp.getUi().alert(
			'Attendance sheet not found!\n\nUse OmniHR > Attendance > Create Attendance Sheet first.'
		);
		return;
	}

	if (attendanceList.length === 0) {
		SpreadsheetApp.getUi().alert(
			'No attendance data found in the Attendance sheet.'
		);
		return;
	}

	// Get month/year from user
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

	// Apply attendance to sheet
	try {
		const daysInMonth = new Date(year, month + 1, 0).getDate();
		applyAttendanceHours(sheet, attendanceList, month, year);

		SpreadsheetApp.getUi().alert(
			`Attendance hours applied!\n\n` +
				`• Month: ${month + 1}/${year}\n` +
				`• Days in month: ${daysInMonth}\n` +
				`• Attendance rows: ${attendanceList.length}\n\n` +
				`Check the Execution Log for details.`
		);
	} catch (error) {
		Logger.log('Error applying attendance: ' + error.message);
		SpreadsheetApp.getUi().alert('Error: ' + error.message);
	}
}

/**
 * Apply attendance hours to a sheet (OPTIMIZED with batch operations)
 * Matches by Employee ID + Project name from main sheet to attendance sheet
 * Only applies hours when both ID and Project match
 */
function applyAttendanceHours(sheet, attendanceList, month, year) {
	const daysInMonth = new Date(year, month + 1, 0).getDate();
	const lastRow = sheet.getLastRow();
	const dayNames = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];

	// Column configuration for totals
	const TOTAL_HOURS_COL = 7; // Column G
	const TOTAL_DAYS_COL = 8; // Column H
	const TOTAL_DAYS_OFF_COL = 9; // Column I

	Logger.log(
		`Applying attendance for ${
			month + 1
		}/${year}, days in month: ${daysInMonth}`
	);
	Logger.log(`Main sheet last row: ${lastRow}`);
	Logger.log(`Attendance list has ${attendanceList.length} rows`);

	// Build a lookup map from attendance list: key = "EMPID|PROJECT"
	const attendanceLookup = {};
	for (let i = 0; i < attendanceList.length; i++) {
		const att = attendanceList[i];
		const key = `${att.empId}|${att.project.toUpperCase()}`;
		attendanceLookup[key] = att;
	}

	// Build day columns mapping with "Validated" checkbox columns after each Friday
	const dayColumns = {};
	const validatedColumns = [];
	let currentCol = CONFIG.FIRST_DAY_COL;

	for (let day = 1; day <= daysInMonth; day++) {
		const date = new Date(year, month, day);
		const dayOfWeek = date.getDay();
		dayColumns[day] = currentCol;
		currentCol++;
		if (dayOfWeek === 5) {
			// Friday
			validatedColumns.push(currentCol);
			currentCol++;
		}
	}

	const lastDayCol = Math.max(
		...Object.values(dayColumns),
		...validatedColumns
	);
	const totalCols = lastDayCol - CONFIG.FIRST_DAY_COL + 1;

	// Read employee data
	const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;
	if (numRows <= 0) {
		Logger.log('No data rows found in sheet');
		return;
	}

	const employeeData = sheet
		.getRange(CONFIG.FIRST_DATA_ROW, 1, numRows, CONFIG.PROJECT_COL)
		.getValues();

	// STEP 1: Delete ALL columns from K onwards (completely removes checkboxes and old data)
	Logger.log('Deleting existing columns from K onwards...');
	const sheetLastCol = sheet.getLastColumn();

	if (sheetLastCol >= CONFIG.FIRST_DAY_COL) {
		const numColsToDelete = sheetLastCol - CONFIG.FIRST_DAY_COL + 1;
		Logger.log(
			`Deleting ${numColsToDelete} columns starting from column ${CONFIG.FIRST_DAY_COL}`
		);
		sheet.deleteColumns(CONFIG.FIRST_DAY_COL, numColsToDelete);
	}

	// Insert fresh columns for the new data
	Logger.log(`Inserting ${totalCols} fresh columns...`);
	sheet.insertColumnsAfter(CONFIG.FIRST_DAY_COL - 1, totalCols);
	SpreadsheetApp.flush();

	// Also clear any existing conditional formatting rules for leaves
	const existingRules = sheet.getConditionalFormatRules();
	const filteredRules = existingRules.filter((rule) => {
		try {
			const bg = rule.getBooleanCondition();
			if (bg) {
				const condition = bg.getCriteriaType();
				if (condition === SpreadsheetApp.BooleanCriteria.NUMBER_EQUAL_TO) {
					const values = bg.getCriteriaValues();
					if (values && (values[0] === 0 || values[0] === 4)) {
						return false; // Remove leave rules
					}
				}
			}
		} catch (e) {}
		return true;
	});
	sheet.setConditionalFormatRules(filteredRules);

	// Build header arrays for batch update
	const headerRow1 = [];
	const headerRow2 = [];
	const headerBgColors = []; // Background colors for header
	for (let col = CONFIG.FIRST_DAY_COL; col <= lastDayCol; col++) {
		const dayForCol = Object.keys(dayColumns).find(
			(d) => dayColumns[d] === col
		);
		if (dayForCol) {
			const date = new Date(year, month, parseInt(dayForCol));
			const dayOfWeek = date.getDay();
			headerRow1.push(dayNames[dayOfWeek]);
			headerRow2.push(parseInt(dayForCol));
			// Weekday = #356854, Weekend = #efefef
			headerBgColors.push(
				dayOfWeek >= 1 && dayOfWeek <= 5 ? '#356854' : '#efefef'
			);
		} else if (validatedColumns.includes(col)) {
			headerRow1.push('');
			headerRow2.push('Validated');
			headerBgColors.push('#356854'); // Validated column same as weekday
		} else {
			headerRow1.push('');
			headerRow2.push('');
			headerBgColors.push(null);
		}
	}

	// Batch update headers
	Logger.log('Setting up headers...');
	sheet
		.getRange(CONFIG.DAY_NAME_ROW, CONFIG.FIRST_DAY_COL, 1, totalCols)
		.setValues([headerRow1]);
	sheet
		.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_DAY_COL, 1, totalCols)
		.setValues([headerRow2]);

	// Apply header background colors only to dates row (row 2), not day abbreviation row (row 1)
	const headerRange2 = sheet.getRange(
		CONFIG.HEADER_ROW,
		CONFIG.FIRST_DAY_COL,
		1,
		totalCols
	);
	headerRange2.setBackgrounds([headerBgColors]);

	// Set font color to white for weekday headers (#356854), black for weekend (#efefef)
	const fontColors = headerBgColors.map((bg) =>
		bg === '#356854' ? '#FFFFFF' : '#000000'
	);
	headerRange2.setFontColors([fontColors]);

	// Set column widths: day columns = 46, validated columns = 105
	Logger.log('Setting column widths...');
	for (let col = CONFIG.FIRST_DAY_COL; col <= lastDayCol; col++) {
		if (validatedColumns.includes(col)) {
			sheet.setColumnWidth(col, 105);
		} else {
			sheet.setColumnWidth(col, 46);
		}
	}

	// Build data arrays for batch update
	Logger.log('Building data arrays...');
	const dataValues = [];
	const backgroundColors = [];
	const formulas = []; // For totals columns G, H, I
	let matchedCount = 0;

	for (let i = 0; i < employeeData.length; i++) {
		const empId = String(employeeData[i][0]).trim().toUpperCase();
		const project = String(employeeData[i][CONFIG.PROJECT_COL - 1])
			.trim()
			.toUpperCase();

		if (!empId) {
			// Empty row - fill with empty values
			dataValues.push(new Array(totalCols).fill(''));
			backgroundColors.push(new Array(totalCols).fill(null));
			formulas.push(['', '', '']);
			continue;
		}

		const key = `${empId}|${project}`;
		const record = attendanceLookup[key];
		const hours = record ? record.hours : 0;
		if (record) matchedCount++;

		const rowValues = [];
		const rowColors = [];

		for (let col = CONFIG.FIRST_DAY_COL; col <= lastDayCol; col++) {
			const dayForCol = Object.keys(dayColumns).find(
				(d) => dayColumns[d] === col
			);

			if (dayForCol) {
				const date = new Date(year, month, parseInt(dayForCol));
				const dayOfWeek = date.getDay();
				const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

				if (isWeekend) {
					rowValues.push('');
					rowColors.push(CONFIG.COLORS.WEEKEND);
				} else {
					rowValues.push(hours);
					rowColors.push(null);
				}
			} else if (validatedColumns.includes(col)) {
				rowValues.push(true); // Checkbox checked
				rowColors.push(null);
			} else {
				rowValues.push('');
				rowColors.push(null);
			}
		}

		dataValues.push(rowValues);
		backgroundColors.push(rowColors);

		// Formulas for this row
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

	// Batch update data values
	Logger.log('Applying data values...');
	const dataRange = sheet.getRange(
		CONFIG.FIRST_DATA_ROW,
		CONFIG.FIRST_DAY_COL,
		numRows,
		totalCols
	);
	dataRange.setValues(dataValues);

	// Batch update background colors
	Logger.log('Applying background colors...');
	dataRange.setBackgrounds(backgroundColors);

	// Insert checkboxes in validated columns (batch)
	Logger.log('Setting up checkboxes...');
	for (const col of validatedColumns) {
		try {
			const checkboxRange = sheet.getRange(
				CONFIG.FIRST_DATA_ROW,
				col,
				numRows,
				1
			);
			checkboxRange.insertCheckboxes();
			checkboxRange.setValue(true);
		} catch (e) {
			// Checkboxes might already exist
			Logger.log(`Checkbox column ${col}: ${e.message}`);
		}
	}

	// Batch update formulas
	Logger.log('Applying formulas...');
	sheet
		.getRange(CONFIG.FIRST_DATA_ROW, TOTAL_HOURS_COL, numRows, 3)
		.setFormulas(formulas);

	// Add conditional formatting for validated weeks only (green)
	// Leave colors are applied directly when leave data is synced
	Logger.log('Adding conditional formatting...');
	addValidatedConditionalFormatting(sheet, month, year);

	Logger.log(`Matched ${matchedCount} rows with attendance data`);
	SpreadsheetApp.flush();
	Logger.log('Attendance hours applied successfully');
}

/**
 * Sync leave data for current month
 */
function syncCurrentMonth() {
	const now = new Date();
	syncLeaveDataForMonth(now.getMonth(), now.getFullYear());
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

	const month = parseInt(monthResponse.getResponseText()) - 1; // Convert to 0-indexed
	const year = parseInt(yearResponse.getResponseText());

	if (isNaN(month) || month < 0 || month > 11 || isNaN(year)) {
		ui.alert('Invalid month or year');
		return;
	}

	syncLeaveDataForMonth(month, year);
}

/**
 * Main sync function for a specific month
 * Creates or uses a sheet named after the month (e.g., "January 2025")
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 */
function syncLeaveDataForMonth(month, year) {
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	// Create sheet name based on month/year (e.g., "January 2025")
	const sheetName = getMonthSheetName(month, year);

	// Get or create the sheet for this month
	let sheet = ss.getSheetByName(sheetName);
	if (!sheet) {
		// Create new sheet for this month
		sheet = ss.insertSheet(sheetName);
		Logger.log(`Created new sheet: ${sheetName}`);
	} else {
		Logger.log(`Using existing sheet: ${sheetName}`);
	}

	Logger.log(
		`Syncing leave data for ${month + 1}/${year} to sheet "${sheetName}"`
	);

	try {
		// Apply attendance data first (base hours for each employee)
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

		// Fetch leave data for each employee
		Logger.log('Fetching leave data...');
		const leaveData = fetchLeaveDataForMonth(token, employees, month, year);
		Logger.log(`Found ${Object.keys(leaveData).length} employees with leave`);

		// Overlay leave data on top of attendance
		Logger.log('Applying leave data...');
		updateSheetWithLeaveData(sheet, leaveData, month, year);

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
	}
}

/**
 * Get access token from OmniHR using username/password
 */
function getAccessToken() {
	const props = PropertiesService.getScriptProperties();
	const baseUrl = props.getProperty('OMNIHR_BASE_URL');
	const subdomain = props.getProperty('OMNIHR_SUBDOMAIN');
	const username = props.getProperty('OMNIHR_USERNAME');
	const password = props.getProperty('OMNIHR_PASSWORD');

	if (!baseUrl || !subdomain || !username || !password) {
		throw new Error(
			'API credentials not configured. Use OmniHR > Setup API Credentials'
		);
	}

	const response = UrlFetchApp.fetch(`${baseUrl}/auth/token/`, {
		method: 'post',
		contentType: 'application/x-www-form-urlencoded',
		payload: `username=${encodeURIComponent(
			username
		)}&password=${encodeURIComponent(password)}`,
		headers: {
			'x-subdomain': subdomain,
		},
		muteHttpExceptions: true,
	});

	const responseText = response.getContentText();
	const data = JSON.parse(responseText);

	const token = data.access || data.token || data.access_token;
	if (token) {
		return token;
	}

	throw new Error('Failed to get access token: ' + responseText);
}

/**
 * Make authenticated API request
 */
function apiRequest(token, endpoint, params = {}) {
	const props = PropertiesService.getScriptProperties();
	const baseUrl = props.getProperty('OMNIHR_BASE_URL');
	const subdomain = props.getProperty('OMNIHR_SUBDOMAIN');

	let url = `${baseUrl}${endpoint}`;

	if (Object.keys(params).length > 0) {
		const queryString = Object.entries(params)
			.map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
			.join('&');
		url += '?' + queryString;
	}

	const response = UrlFetchApp.fetch(url, {
		method: 'get',
		headers: {
			Authorization: `Bearer ${token}`,
			'x-subdomain': subdomain,
			'Content-Type': 'application/json',
		},
		muteHttpExceptions: true,
	});

	return JSON.parse(response.getContentText());
}

/**
 * Fetch all employees with pagination
 */
function fetchAllEmployees(token) {
	let allEmployees = [];
	let page = 1;
	let hasMore = true;

	while (hasMore) {
		const response = apiRequest(token, '/employee/list/', {
			page,
			page_size: 100,
		});
		const results = response.results || response;
		allEmployees = allEmployees.concat(results);
		hasMore = response.next !== null && response.next !== undefined;
		page++;
	}

	return allEmployees;
}

/**
 * Format date as DD/MM/YYYY
 */
function formatDateDMY(d) {
	const day = String(d.getDate()).padStart(2, '0');
	const month = String(d.getMonth() + 1).padStart(2, '0');
	const year = d.getFullYear();
	return `${day}/${month}/${year}`;
}

/**
 * Build batch request objects for UrlFetchApp.fetchAll()
 */
function buildBatchRequests(token, employees, startDate, endDate) {
	const props = PropertiesService.getScriptProperties();
	const baseUrl = props.getProperty('OMNIHR_BASE_URL');
	const subdomain = props.getProperty('OMNIHR_SUBDOMAIN');

	const headers = {
		Authorization: `Bearer ${token}`,
		'x-subdomain': subdomain,
		'Content-Type': 'application/json',
	};

	const requests = [];
	const requestMeta = []; // Track which request belongs to which employee

	for (const emp of employees) {
		const userId = emp.id || emp.user_id;
		const empName = emp.full_name || emp.name || `User ${userId}`;

		// Base data request
		requests.push({
			url: `${baseUrl}/employee/2.0/users/${userId}/base-data/`,
			method: 'get',
			headers: headers,
			muteHttpExceptions: true,
		});
		requestMeta.push({ userId, empName, type: 'base' });

		// Time-off calendar request
		const calendarUrl = `${baseUrl}/employee/1.1/${userId}/time-off-calendar/?start_date=${formatDateDMY(
			startDate
		)}&end_date=${formatDateDMY(endDate)}`;
		requests.push({
			url: calendarUrl,
			method: 'get',
			headers: headers,
			muteHttpExceptions: true,
		});
		requestMeta.push({ userId, empName, type: 'calendar' });
	}

	return { requests, requestMeta };
}

/**
 * Fetch leave data for all employees using batch requests (FAST)
 */
function fetchLeaveDataForMonth(token, employees, month, year) {
	const startDate = new Date(year, month, 1);
	const endDate = new Date(year, month + 1, 0); // Last day of month

	const leaveData = {};
	const BATCH_SIZE = 50; // Process 50 employees at a time (100 requests)

	// Process in batches to avoid hitting limits
	for (let i = 0; i < employees.length; i += BATCH_SIZE) {
		const batch = employees.slice(i, i + BATCH_SIZE);
		Logger.log(
			`Processing batch ${Math.floor(i / BATCH_SIZE) + 1}/${Math.ceil(
				employees.length / BATCH_SIZE
			)} (${batch.length} employees)`
		);

		const { requests, requestMeta } = buildBatchRequests(
			token,
			batch,
			startDate,
			endDate
		);

		// Execute all requests in parallel!
		const responses = UrlFetchApp.fetchAll(requests);

		// Group responses by employee
		const employeeData = {};
		for (let j = 0; j < responses.length; j++) {
			const meta = requestMeta[j];
			const response = responses[j];

			try {
				const responseCode = response.getResponseCode();
				const responseText = response.getContentText();

				// Check for HTTP errors
				if (responseCode !== 200) {
					Logger.log(
						`API error for ${meta.empName} (${
							meta.type
						}): HTTP ${responseCode} - ${responseText.substring(0, 200)}`
					);
					continue;
				}

				const data = JSON.parse(responseText);

				// Check for API error response
				if (data.error || data.detail || data.message) {
					Logger.log(
						`API error for ${meta.empName}: ${
							data.error || data.detail || data.message
						}`
					);
					continue;
				}

				if (!employeeData[meta.userId]) {
					employeeData[meta.userId] = { empName: meta.empName };
				}

				if (meta.type === 'base') {
					employeeData[meta.userId].baseData = data.data || data;
				} else {
					employeeData[meta.userId].calendar = data;
				}
			} catch (e) {
				Logger.log(`Error parsing response for ${meta.empName}: ${e.message}`);
			}
		}

		// Process the collected data
		for (const [userId, data] of Object.entries(employeeData)) {
			const employeeId = data.baseData?.employee_id;
			const empName = data.empName;
			const calendar = data.calendar || {};

			// Filter approved requests (status === 3)
			const approvedRequests = (calendar.time_off_request || []).filter(
				(r) => r.status === 3
			);

			if (approvedRequests.length === 0) continue;

			// Process leave days
			const leaveDays = [];

			for (const request of approvedRequests) {
				const leaveStart = parseDateDMY(request.effective_date);
				const leaveEnd = request.end_date
					? parseDateDMY(request.end_date)
					: leaveStart;

				if (!leaveStart) continue;

				// Iterate through each day in the leave range
				const currentDate = new Date(leaveStart);
				while (currentDate <= leaveEnd) {
					const dayOfWeek = currentDate.getDay();
					const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

					if (
						!isWeekend &&
						currentDate.getMonth() === month &&
						currentDate.getFullYear() === year
					) {
						leaveDays.push({
							date: currentDate.getDate(),
							leave_type: request.time_off?.name,
							is_half_day:
								request.effective_date_duration === 2 ||
								request.effective_date_duration === 3,
						});
					}

					currentDate.setDate(currentDate.getDate() + 1);
				}
			}

			if (leaveDays.length > 0) {
				leaveData[employeeId || empName] = {
					employee_id: employeeId,
					employee_name: empName,
					leave_requests: leaveDays,
				};
			}
		}
	}

	return leaveData;
}

/**
 * Parse date in DD/MM/YYYY format
 */
function parseDateDMY(dateStr) {
	if (!dateStr) return null;
	const parts = dateStr.split('/');
	if (parts.length !== 3) return null;
	return new Date(
		parseInt(parts[2]),
		parseInt(parts[1]) - 1,
		parseInt(parts[0])
	);
}

/**
 * Update sheet with leave data
 * Uses conditional formatting for leave colors (red for full day, orange for half day)
 * Applies leave to ALL rows for an employee (multiple projects)
 * Half-day leave is divided proportionally based on each project's hours
 */
function updateSheetWithLeaveData(sheet, leaveData, month, year) {
	const daysInMonth = new Date(year, month + 1, 0).getDate();

	// Build employee lookup from sheet (now includes hours per row)
	const employeeLookup = buildEmployeeLookup(sheet);

	// Get attendance data to get hours per project
	const attendanceList = getAttendanceData() || [];
	// Build lookup: empId -> [{row, project, hours}, ...]
	const attendanceByEmployee = {};
	for (const att of attendanceList) {
		const key = att.empId;
		if (!attendanceByEmployee[key]) {
			attendanceByEmployee[key] = [];
		}
		attendanceByEmployee[key].push({
			project: att.project,
			hours: att.hours,
		});
	}

	// Get day columns mapping (calculated based on month/year)
	const dayColumns = getDayColumns(sheet, daysInMonth, month, year);

	// Collect cells to update - store cell and value pairs for half-day
	const fullDayCells = []; // Will have value 0
	const halfDayCellsMap = {}; // Map of cellA1 -> value (divided among projects)
	let matchedEmployees = 0;

	for (const [key, empData] of Object.entries(leaveData)) {
		const { employee_id, employee_name, leave_requests } = empData;

		// Find ALL employee rows (employee may have multiple projects)
		const rows = findEmployeeRows(employeeLookup, employee_id, employee_name);
		if (rows.length === 0) {
			Logger.log(`Employee not found: ${employee_name} (${employee_id})`);
			continue;
		}

		matchedEmployees++;
		Logger.log(
			`Found ${rows.length} rows for ${employee_name} (${employee_id})`
		);

		// Get attendance info for this employee to calculate proportional hours
		const empAttendance = attendanceByEmployee[employee_id.toUpperCase()] || [];
		const totalHours =
			empAttendance.reduce((sum, att) => sum + att.hours, 0) || 8;

		// Build row -> hours mapping by reading from sheet
		const rowHoursMap = {};
		for (const row of rows) {
			// Get the hours from the first weekday column for this row
			const firstWeekdayCol = Object.values(dayColumns)[0];
			if (firstWeekdayCol) {
				const cellValue = sheet.getRange(row, firstWeekdayCol).getValue();
				rowHoursMap[row] = parseFloat(cellValue) || 8;
			} else {
				rowHoursMap[row] = 8;
			}
		}
		const rowTotalHours =
			Object.values(rowHoursMap).reduce((sum, h) => sum + h, 0) || 8;

		// Collect cells by leave type - apply to ALL rows for this employee
		for (const leave of leave_requests) {
			const col = dayColumns[leave.date];
			if (!col) continue;

			// Apply to each row (each project) for this employee
			for (const row of rows) {
				const cellA1 = columnToLetter(col) + row;

				if (leave.is_half_day) {
					// Half-day leave: divide 4 hours proportionally based on project hours
					// e.g., Project A=1hr, B=1hr, C=6hr (total 8hr)
					// Half-day = 4hr -> A gets 0.5, B gets 0.5, C gets 3
					const projectHours = rowHoursMap[row] || 8;
					const proportion = projectHours / rowTotalHours;
					const halfDayHoursForProject = 4 * proportion;
					halfDayCellsMap[cellA1] = halfDayHoursForProject;
					Logger.log(
						`Half-day for row ${row}: ${projectHours}/${rowTotalHours} * 4 = ${halfDayHoursForProject}`
					);
				} else {
					fullDayCells.push(cellA1);
				}
			}
		}
	}

	// Apply full day leave (value 0 + red background)
	if (fullDayCells.length > 0) {
		Logger.log(
			`Applying full-day leave to ${fullDayCells.length} cells (value 0, red)`
		);
		const fullDayRanges = sheet.getRangeList(fullDayCells);
		fullDayRanges.setValue(0);
		fullDayRanges.setBackground(CONFIG.COLORS.FULL_DAY);
		fullDayRanges.setFontColor('#FFFFFF');
		fullDayRanges.setFontWeight('bold');
	}

	// Apply half day leave with divided values (value + orange background)
	const halfDayCellEntries = Object.entries(halfDayCellsMap);
	if (halfDayCellEntries.length > 0) {
		Logger.log(
			`Applying half-day leave to ${halfDayCellEntries.length} cells (orange)`
		);
		// Apply values and colors directly to each cell
		for (const [cellA1, value] of halfDayCellEntries) {
			const range = sheet.getRange(cellA1);
			range.setValue(parseFloat(value));
			range.setBackground(CONFIG.COLORS.HALF_DAY);
			range.setFontColor('#000000');
			range.setFontWeight('bold');
		}
	}

	// Add conditional formatting for validated weeks only (green)
	addValidatedConditionalFormatting(sheet, month, year);

	// Force flush all pending changes
	SpreadsheetApp.flush();

	Logger.log(
		`Matched ${matchedEmployees} employees, updated ${
			fullDayCells.length + halfDayCellEntries.length
		} cells`
	);
}

/**
 * Add conditional formatting for validated weeks only (green)
 * Leave colors (red/orange) are applied directly to cells, not via conditional formatting
 */
function addValidatedConditionalFormatting(sheet, month, year) {
	// Calculate the range based on month/year
	const { dayColumns, validatedColumns } = calculateDayColumns(month, year);

	const lastRow = sheet.getLastRow();
	const firstCol = CONFIG.FIRST_DAY_COL;

	// Make sure we have valid dimensions
	const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;

	if (numRows <= 0) {
		Logger.log('addValidatedConditionalFormatting: Invalid range dimensions');
		return;
	}

	// Clear ALL existing conditional formatting rules (start fresh)
	sheet.setConditionalFormatRules([]);

	const rules = [];

	// Green for validated weeks (when checkbox is TRUE)
	// Only apply to weekdays (not weekends) + the checkbox column itself
	let weekStartCol = firstCol;
	for (const validatedCol of validatedColumns) {
		if (validatedCol >= weekStartCol) {
			const checkboxColLetter = columnToLetter(validatedCol);
			const weekdayRanges = [];

			// Find weekday columns in this week (exclude weekends)
			for (const [dayStr, col] of Object.entries(dayColumns)) {
				if (col >= weekStartCol && col < validatedCol) {
					const day = parseInt(dayStr);
					const date = new Date(year, month, day);
					const dayOfWeek = date.getDay();
					// Only include weekdays (Mon-Fri: 1-5)
					if (dayOfWeek >= 1 && dayOfWeek <= 5) {
						weekdayRanges.push(
							sheet.getRange(CONFIG.FIRST_DATA_ROW, col, numRows, 1)
						);
					}
				}
			}

			// Also include the checkbox column itself
			weekdayRanges.push(
				sheet.getRange(CONFIG.FIRST_DATA_ROW, validatedCol, numRows, 1)
			);

			if (weekdayRanges.length > 0) {
				// Custom formula: apply green when checkbox is TRUE
				const greenRule = SpreadsheetApp.newConditionalFormatRule()
					.whenFormulaSatisfied(
						`=$${checkboxColLetter}${CONFIG.FIRST_DATA_ROW}=TRUE`
					)
					.setBackground('#B8E1CD')
					.setRanges(weekdayRanges)
					.build();
				rules.push(greenRule);
			}
		}
		weekStartCol = validatedCol + 1;
	}

	sheet.setConditionalFormatRules(rules);
	Logger.log(
		`Applied ${rules.length} conditional formatting rules (validated weeks only)`
	);
}

/**
 * Convert column number to letter (1=A, 2=B, 27=AA, etc..)
 */
function columnToLetter(column) {
	let letter = '';
	while (column > 0) {
		const mod = (column - 1) % 26;
		letter = String.fromCharCode(65 + mod) + letter;
		column = Math.floor((column - mod) / 26);
	}
	return letter;
}

/**
 * Build lookup of employees from sheet
 * Stores arrays of rows since an employee can have multiple projects
 */
function buildEmployeeLookup(sheet) {
	const lastRow = sheet.getLastRow();
	const lookup = {
		byId: {},
		byName: {},
	};

	for (let row = CONFIG.FIRST_DATA_ROW; row <= lastRow; row++) {
		const id = sheet.getRange(row, CONFIG.EMPLOYEE_ID_COL).getValue();
		const name = sheet.getRange(row, CONFIG.EMPLOYEE_NAME_COL).getValue();

		if (id) {
			const idUpper = String(id).trim().toUpperCase();
			if (!lookup.byId[idUpper]) {
				lookup.byId[idUpper] = [];
			}
			lookup.byId[idUpper].push(row);
		}

		if (name) {
			const nameLower = String(name).trim().toLowerCase();
			if (!lookup.byName[nameLower]) {
				lookup.byName[nameLower] = [];
			}
			lookup.byName[nameLower].push(row);
		}
	}

	return lookup;
}

/**
 * Find all employee rows by ID or name (returns array of rows)
 * An employee can have multiple rows if they work on multiple projects
 */
function findEmployeeRows(lookup, employeeId, employeeName) {
	// Try by ID first
	if (employeeId) {
		const rows = lookup.byId[employeeId.toUpperCase()];
		if (rows && rows.length > 0) return rows;
	}

	// Try by exact name
	if (employeeName) {
		const nameLower = employeeName.trim().toLowerCase();
		const rows = lookup.byName[nameLower];
		if (rows && rows.length > 0) return rows;

		// Try partial match
		for (const [sheetName, rows] of Object.entries(lookup.byName)) {
			if (sheetName.includes(nameLower) || nameLower.includes(sheetName)) {
				return rows;
			}
		}
	}

	return [];
}

/**
 * Calculate day columns mapping based on month/year
 * Accounts for "Validated" checkbox columns after each Friday
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @returns {Object} { dayColumns: {day: col}, validatedColumns: [cols] }
 */
function calculateDayColumns(month, year) {
	const daysInMonth = new Date(year, month + 1, 0).getDate();
	const dayColumns = {};
	const validatedColumns = [];
	let currentCol = CONFIG.FIRST_DAY_COL;

	for (let day = 1; day <= daysInMonth; day++) {
		const date = new Date(year, month, day);
		const dayOfWeek = date.getDay();
		dayColumns[day] = currentCol;
		currentCol++;
		if (dayOfWeek === 5) {
			// Friday - add validated column after
			validatedColumns.push(currentCol);
			currentCol++;
		}
	}

	// Add validated column after the last day if there are weekdays after the last Friday
	const lastDayOfMonth = new Date(year, month, daysInMonth);
	const lastDayOfWeek = lastDayOfMonth.getDay();

	// Only add if last day is Mon-Thu (1-4), meaning there are weekdays not yet validated
	// If last day is Friday (5), it was already handled in the loop
	// If last day is Sat (6) or Sun (0), no weekdays to validate after last Friday
	if (lastDayOfWeek >= 1 && lastDayOfWeek <= 4) {
		validatedColumns.push(currentCol);
	}

	return { dayColumns, validatedColumns };
}

/**
 * Get mapping of day numbers to column indices for a specific month/year
 * Uses calculated positions based on weekday layout
 */
function getDayColumns(sheet, daysInMonth, month, year) {
	// If month/year provided, calculate columns based on date logic
	if (month !== undefined && year !== undefined) {
		const result = calculateDayColumns(month, year);
		Logger.log(
			`getDayColumns calculated ${
				Object.keys(result.dayColumns).length
			} day columns for ${month + 1}/${year}`
		);
		return result.dayColumns;
	}

	// Fallback: read from header row
	const dayColumns = {};
	const lastCol = sheet.getLastColumn();
	if (lastCol >= CONFIG.FIRST_DAY_COL) {
		const numColsToRead = Math.min(50, lastCol - CONFIG.FIRST_DAY_COL + 1);
		const headerRow = sheet
			.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_DAY_COL, 1, numColsToRead)
			.getValues()[0];

		for (let i = 0; i < headerRow.length; i++) {
			const value = headerRow[i];
			const dayNum = parseInt(value);
			if (!isNaN(dayNum) && dayNum >= 1 && dayNum <= 31) {
				dayColumns[dayNum] = CONFIG.FIRST_DAY_COL + i;
			}
		}
	}

	Logger.log(
		`getDayColumns found ${
			Object.keys(dayColumns).length
		} day columns from header`
	);
	return dayColumns;
}

/**
 * Show dialog to set API credentials
 */
function showCredentialsDialog() {
	const html = HtmlService.createHtmlOutput(
		`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 10px; font-weight: bold; }
      input { width: 100%; padding: 8px; margin-top: 5px; box-sizing: border-box; }
      button { margin-top: 20px; padding: 10px 20px; background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #357abd; }
      .hint { font-size: 12px; color: #666; margin-top: 2px; }
    </style>
    <form onsubmit="saveCredentials(event)">
      <label>Base URL:</label>
      <input type="text" id="baseUrl" placeholder="https://api.omnihr.co/api/v1" required>
      <div class="hint">e.g., https://api.omnihr.co/api/v1</div>

      <label>Subdomain:</label>
      <input type="text" id="subdomain" required>

      <label>Username:</label>
      <input type="text" id="username" required>

      <label>Password:</label>
      <input type="password" id="password" required>

      <button type="submit">Save Credentials</button>
    </form>
    <script>
      function saveCredentials(e) {
        e.preventDefault();
        google.script.run
          .withSuccessHandler(() => {
            alert('Credentials saved!');
            google.script.host.close();
          })
          .withFailureHandler((err) => alert('Error: ' + err))
          .saveCredentials(
            document.getElementById('baseUrl').value,
            document.getElementById('subdomain').value,
            document.getElementById('username').value,
            document.getElementById('password').value
          );
      }
    </script>
  `
	)
		.setWidth(400)
		.setHeight(380);

	SpreadsheetApp.getUi().showModalDialog(html, 'OmniHR API Credentials');
}

/**
 * Save credentials to script properties
 */
function saveCredentials(baseUrl, subdomain, username, password) {
	const props = PropertiesService.getScriptProperties();
	props.setProperty('OMNIHR_BASE_URL', baseUrl);
	props.setProperty('OMNIHR_SUBDOMAIN', subdomain);
	props.setProperty('OMNIHR_USERNAME', username);
	props.setProperty('OMNIHR_PASSWORD', password);
}
