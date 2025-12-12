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
			.addItem('Sync Leave Only (Keep Hours)', 'syncLeaveOnly')
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
					.addItem('Enable Monthly Sync', 'setupMonthlyTrigger')
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
 * Creates a new sheet for each month (e.g., "December 2025")
 */
function setupMonthlyTrigger() {
	// Remove existing triggers
	const triggers = ScriptApp.getProjectTriggers();
	triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));

	// Create monthly trigger - runs on day 1 of each month at 6 AM
	ScriptApp.newTrigger('scheduledSync')
		.timeBased()
		.onMonthDay(1)
		.atHour(6)
		.create();

	Logger.log(
		'Monthly sync trigger created - will run on 1st of each month at 6 AM'
	);
	SpreadsheetApp.getUi().alert(
		'Monthly sync enabled!\n\n' +
			'The script will automatically sync the active sheet on the 1st of each month at 6 AM.'
	);
}

/**
 * Setup daily trigger - runs every day at 6 AM
 * Creates/updates a sheet for the current month (e.g., "December 2025")
 */
function setupDailyTrigger() {
	// Remove existing triggers
	const triggers = ScriptApp.getProjectTriggers();
	triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));

	// Create daily trigger at 6 AM
	ScriptApp.newTrigger('scheduledSync')
		.timeBased()
		.everyDays(1)
		.atHour(6)
		.create();

	Logger.log('Daily sync trigger created for 6 AM');
	SpreadsheetApp.getUi().alert(
		'Daily sync enabled!\n\n' +
			'The script will automatically sync the active sheet every day at 6 AM.'
	);
}

/**
 * Scheduled sync function (called by trigger)
 * Syncs the active sheet with current month's leave data
 */
function scheduledSync() {
	const now = new Date();
	const month = now.getMonth();
	const year = now.getFullYear();

	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getActiveSheet();

	Logger.log(`Scheduled sync to active sheet: ${sheet.getName()}`);
	syncLeaveDataToSheet(sheet, month, year);
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

	// Build day columns mapping using shared function (now includes week override columns)
	const { dayColumns, validatedColumns, weekOverrideColumns } =
		calculateDayColumns(month, year);

	const lastDayCol = Math.max(
		...Object.values(dayColumns),
		...validatedColumns,
		...weekOverrideColumns
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

	// STEP 0: Save Override checkbox states AND hour values before reformatting
	// This preserves manual corrections across syncs
	const savedOverrideStates = {};
	const savedHourValues = {}; // Save hour values for overridden weeks
	const sheetLastCol = sheet.getLastColumn();

	if (sheetLastCol >= CONFIG.FIRST_DAY_COL) {
		// Find existing Override columns and day columns by looking at headers
		const headerRange = sheet.getRange(
			CONFIG.HEADER_ROW,
			CONFIG.FIRST_DAY_COL,
			1,
			sheetLastCol - CONFIG.FIRST_DAY_COL + 1
		);
		const headerValues = headerRange.getValues()[0];

		// First pass: find Override columns and which rows have them checked
		// Check for both old "Override" and new "Time off Override" headers
		const overrideColIndices = [];
		for (let i = 0; i < headerValues.length; i++) {
			if (
				headerValues[i] === 'Override' ||
				headerValues[i] === 'Time off Override'
			) {
				overrideColIndices.push(i);
				const overrideCol = CONFIG.FIRST_DAY_COL + i;

				// Save override states for each row
				for (let rowIdx = 0; rowIdx < numRows; rowIdx++) {
					const empId = String(employeeData[rowIdx][0] || '')
						.trim()
						.toUpperCase();
					const project = String(
						employeeData[rowIdx][CONFIG.PROJECT_COL - 1] || ''
					)
						.trim()
						.toUpperCase();
					if (empId) {
						const overrideValue = sheet
							.getRange(CONFIG.FIRST_DATA_ROW + rowIdx, overrideCol)
							.getValue();
						if (overrideValue === true) {
							const weekKey = `${empId}|${project}|${i}`;
							savedOverrideStates[weekKey] = true;
							Logger.log(`Saved override state for ${weekKey}`);

							// Save hour values for this week (find the week's day columns)
							// Look backwards from Override column to find day values
							for (let dayOffset = 1; dayOffset <= 7; dayOffset++) {
								const dayColIdx = i - dayOffset;
								if (dayColIdx >= 0) {
									const dayHeader = headerValues[dayColIdx];
									// Check if it's a day number (1-31) or Validated column
									if (
										typeof dayHeader === 'number' ||
										(typeof dayHeader === 'string' && /^\d+$/.test(dayHeader))
									) {
										const dayCol = CONFIG.FIRST_DAY_COL + dayColIdx;
										const cellValue = sheet
											.getRange(CONFIG.FIRST_DATA_ROW + rowIdx, dayCol)
											.getValue();
										const cellBg = sheet
											.getRange(CONFIG.FIRST_DATA_ROW + rowIdx, dayCol)
											.getBackground();
										const hourKey = `${empId}|${project}|${dayColIdx}`;
										savedHourValues[hourKey] = {
											value: cellValue,
											background: cellBg,
										};
									} else if (dayHeader === 'Validated') {
										break; // Stop at previous week's Validated column
									}
								}
							}
						}
					}
				}
			}
		}
		Logger.log(
			`Saved ${Object.keys(savedOverrideStates).length} override states`
		);
		Logger.log(
			`Saved ${
				Object.keys(savedHourValues).length
			} hour values for overridden weeks`
		);
	}

	// STEP 1: Delete ALL columns from K onwards (completely removes checkboxes and old data)
	Logger.log('Deleting existing columns from K onwards...');

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
		} else if (weekOverrideColumns.includes(col)) {
			headerRow1.push('');
			headerRow2.push('Time off Override');
			headerBgColors.push('#efefef'); // Same as weekend background
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

	// Set font color to white for dark headers (weekday #356854), black for light backgrounds (weekend #efefef)
	const fontColors = headerBgColors.map((bg) =>
		bg === '#356854' ? '#FFFFFF' : '#000000'
	);
	headerRange2.setFontColors([fontColors]);

	// Also set row 1 (day abbreviations) font color to white for dark backgrounds
	const headerRange1 = sheet.getRange(
		CONFIG.DAY_NAME_ROW,
		CONFIG.FIRST_DAY_COL,
		1,
		totalCols
	);
	headerRange1.setFontColors([fontColors]);

	// Set column widths: day columns = 46, validated columns = 105
	Logger.log('Setting column widths...');
	for (let col = CONFIG.FIRST_DAY_COL; col <= lastDayCol; col++) {
		if (validatedColumns.includes(col) || weekOverrideColumns.includes(col)) {
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
				rowValues.push(true); // Validated checkbox checked by default
				rowColors.push(null);
			} else if (weekOverrideColumns.includes(col)) {
				rowValues.push(false); // Override checkbox unchecked by default
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
	Logger.log('Setting up checkboxes for Validated columns...');
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
			Logger.log(`Validated checkbox column ${col}: ${e.message}`);
		}
	}

	// Insert checkboxes in week override columns (batch) - default to false (unchecked)
	// Then restore any previously saved override states
	Logger.log('Setting up checkboxes for Week Override columns...');
	for (let weekIdx = 0; weekIdx < weekOverrideColumns.length; weekIdx++) {
		const col = weekOverrideColumns[weekIdx];
		try {
			const checkboxRange = sheet.getRange(
				CONFIG.FIRST_DATA_ROW,
				col,
				numRows,
				1
			);
			checkboxRange.insertCheckboxes();
			checkboxRange.setValue(false); // Default to unchecked

			// Restore saved override states for this week
			// Calculate the column index relative to FIRST_DAY_COL for matching saved keys
			const colIndex = col - CONFIG.FIRST_DAY_COL;
			for (let rowIdx = 0; rowIdx < numRows; rowIdx++) {
				const empId = String(employeeData[rowIdx][0] || '')
					.trim()
					.toUpperCase();
				const project = String(
					employeeData[rowIdx][CONFIG.PROJECT_COL - 1] || ''
				)
					.trim()
					.toUpperCase();
				if (empId) {
					const weekKey = `${empId}|${project}|${colIndex}`;
					if (savedOverrideStates[weekKey]) {
						sheet.getRange(CONFIG.FIRST_DATA_ROW + rowIdx, col).setValue(true);
						Logger.log(`Restored override state for ${weekKey}`);
					}
				}
			}
		} catch (e) {
			Logger.log(`Override checkbox column ${col}: ${e.message}`);
		}
	}

	// Restore saved hour values for overridden weeks
	if (Object.keys(savedHourValues).length > 0) {
		Logger.log('Restoring hour values for overridden weeks...');
		for (let rowIdx = 0; rowIdx < numRows; rowIdx++) {
			const empId = String(employeeData[rowIdx][0] || '')
				.trim()
				.toUpperCase();
			const project = String(employeeData[rowIdx][CONFIG.PROJECT_COL - 1] || '')
				.trim()
				.toUpperCase();
			if (empId) {
				// Check each day column
				for (let colIdx = 0; colIdx < totalCols; colIdx++) {
					const hourKey = `${empId}|${project}|${colIdx}`;
					if (savedHourValues[hourKey]) {
						const col = CONFIG.FIRST_DAY_COL + colIdx;
						const cell = sheet.getRange(CONFIG.FIRST_DATA_ROW + rowIdx, col);
						cell.setValue(savedHourValues[hourKey].value);
						if (savedHourValues[hourKey].background) {
							cell.setBackground(savedHourValues[hourKey].background);
						}
						Logger.log(
							`Restored hour value for ${hourKey}: ${savedHourValues[hourKey].value}`
						);
					}
				}
			}
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

	const month = parseInt(monthResponse.getResponseText()) - 1; // Convert to 0-indexed
	const year = parseInt(yearResponse.getResponseText());

	if (isNaN(month) || month < 0 || month > 11 || isNaN(year)) {
		ui.alert('Invalid month or year');
		return;
	}

	// Use active sheet
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getActiveSheet();
	const sheetName = sheet.getName();

	Logger.log(`Syncing to active sheet: ${sheetName}`);
	syncLeaveDataToSheet(sheet, month, year);
}

/**
 * Sync leave only - applies leave colors/values without reformatting attendance hours
 * Keeps existing employee working hours intact
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

	const month = parseInt(monthResponse.getResponseText()) - 1; // 0-indexed
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

		// Get access token
		const token = getAccessToken();
		if (!token) {
			ui.alert('Failed to get API token. Check your credentials.');
			return;
		}

		// Fetch employees
		const employees = fetchAllEmployees(token);
		if (!employees || employees.length === 0) {
			ui.alert('No employees found');
			return;
		}

		// Fetch leave data
		const leaveData = fetchLeaveDataForMonth(token, employees, month, year);
		if (!leaveData || Object.keys(leaveData).length === 0) {
			ui.alert('No leave data found for this month');
			return;
		}

		// Apply leave colors and values only (no attendance reformatting)
		// Pass useSheetHours=true to read hours from current sheet values instead of attendance data
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

		// Ensure all changes are written
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
		// Always flush at the end to clear the "Working..." spinner
		SpreadsheetApp.flush();
	}
}

/**
 * Apply leave colors only to the active sheet without syncing attendance
 * This finds matching employees and dates, then applies red/orange colors
 * Does not reformat or touch attendance data
 * Splits leave evenly across multiple project rows for same employee
 */
function applyLeaveColorsOnly() {
	const ui = SpreadsheetApp.getUi();
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

	// Prompt for month/year
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

	const month = parseInt(monthResponse.getResponseText()) - 1; // 0-indexed
	const year = parseInt(yearResponse.getResponseText());

	if (isNaN(month) || month < 0 || month > 11 || isNaN(year)) {
		ui.alert('Invalid month or year');
		return;
	}

	try {
		Logger.log(
			`Applying leave colors for ${month + 1}/${year} to active sheet`
		);

		// Fetch leave data from API
		const leaveData = fetchLeaveDataFromAPI(month, year);
		if (!leaveData || Object.keys(leaveData).length === 0) {
			ui.alert('No leave data found for this month');
			return;
		}

		// Apply leave colors only (no attendance sync, no reformatting)
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

/**
 * Apply leave colors to sheet without modifying attendance data
 * Only colors cells for employees/dates that have leave
 * Half-day: deduct 4 hours total, starting from smallest projects first
 * Respects Week Override column - skips updates for rows with override checked
 */
function applyLeaveColorsToSheet(sheet, leaveData, month, year) {
	const { dayColumns, weekRanges } = calculateDayColumns(month, year);

	// Build a map of day -> week override column for quick lookup
	const dayToOverrideCol = {};
	for (const weekRange of weekRanges) {
		for (const [dayStr, col] of Object.entries(dayColumns)) {
			if (col >= weekRange.startCol && col <= weekRange.endCol) {
				dayToOverrideCol[dayStr] = weekRange.overrideCol;
			}
		}
	}

	// Build employee lookup from sheet
	const lastRow = sheet.getLastRow();
	const employeeLookup = {};

	// Read employee data (columns A-C: Employee ID, Name, Project)
	if (lastRow >= CONFIG.FIRST_DATA_ROW) {
		const employeeRange = sheet.getRange(
			CONFIG.FIRST_DATA_ROW,
			1,
			lastRow - CONFIG.FIRST_DATA_ROW + 1,
			3
		);
		const employeeValues = employeeRange.getValues();

		for (let i = 0; i < employeeValues.length; i++) {
			const row = CONFIG.FIRST_DATA_ROW + i;
			const empId = String(employeeValues[i][0] || '')
				.trim()
				.toUpperCase();
			const empName = String(employeeValues[i][1] || '')
				.trim()
				.toUpperCase();

			if (empId) {
				if (!employeeLookup[empId]) employeeLookup[empId] = [];
				employeeLookup[empId].push(row);
			}
			if (empName) {
				if (!employeeLookup[empName]) employeeLookup[empName] = [];
				if (!employeeLookup[empName].includes(row)) {
					employeeLookup[empName].push(row);
				}
			}
		}
	}

	// Collect cells to color
	const fullDayCells = [];
	const halfDayCells = [];
	let matchedEmployees = 0;

	for (const [key, empData] of Object.entries(leaveData)) {
		const { employee_id, employee_name, leave_requests } = empData;

		// Find employee rows
		const rows = findEmployeeRows(employeeLookup, employee_id, employee_name);
		if (rows.length === 0) {
			Logger.log(`Employee not found: ${employee_name} (${employee_id})`);
			continue;
		}

		matchedEmployees++;
		const numRows = rows.length;
		Logger.log(`Found ${numRows} rows for ${employee_name}`);

		// Build row -> hours mapping
		// Try to find a column with a valid hour value (not 0, not empty)
		// This avoids reading from columns that already have leave applied
		const rowHoursMap = {};
		const allDayCols = Object.values(dayColumns);

		for (const row of rows) {
			let foundHours = null;

			// Scan through day columns to find a non-zero, non-leave value
			for (const col of allDayCols) {
				const cellValue = sheet.getRange(row, col).getValue();
				const numValue = parseFloat(cellValue);

				// Valid hours should be > 0 and typically <= 8
				// Skip 0 (full-day leave) and values that look like half-day leave
				if (!isNaN(numValue) && numValue > 0) {
					foundHours = numValue;
					break;
				}
			}

			// If no hours found or 0, default to 8 (leave booked before hours allocated)
			rowHoursMap[row] = foundHours || CONFIG.DEFAULT_HOURS;
			Logger.log(`Row ${row}: found ${rowHoursMap[row]} hours`);
		}

		// Apply leaves to all rows for this employee
		// Full-day: all rows get 0
		// Half-day: deduct 4 hours total, starting from smallest projects first
		for (const leave of leave_requests) {
			const col = dayColumns[leave.date];
			if (!col) continue;

			// Check Week Override for each row - skip rows where override is checked
			const overrideCol = dayToOverrideCol[leave.date];

			// Filter rows that don't have Week Override checked
			const activeRows = rows.filter((row) => {
				if (!overrideCol) return true; // No override column for this week
				const overrideValue = sheet.getRange(row, overrideCol).getValue();
				if (overrideValue === true) {
					Logger.log(
						`Skipping row ${row} for day ${leave.date} - Week Override is checked`
					);
					return false;
				}
				return true;
			});

			if (activeRows.length === 0) {
				Logger.log(
					`All rows have Week Override checked for day ${leave.date}, skipping`
				);
				continue;
			}

			if (leave.is_half_day) {
				// Half-day leave = 4 hours of work remaining
				// Divide the REMAINING 4 hours EQUALLY across all projects
				const totalRemainingHours = CONFIG.HALF_DAY_HOURS; // 4 hours
				const hoursPerProject = totalRemainingHours / activeRows.length;

				Logger.log(
					`Half-day leave: ${activeRows.length} projects, ${hoursPerProject}hr each (${totalRemainingHours}hr total remaining)`
				);

				// Apply equal remaining hours to each row - all half-day = ORANGE
				for (const row of activeRows) {
					const cellA1 = columnToLetter(col) + row;

					// Half-day leave always gets ORANGE color with equal hours
					halfDayCells.push({ cell: cellA1, value: hoursPerProject });
					Logger.log(`Row ${row}: ${hoursPerProject}hr (ORANGE)`);
				}
			} else {
				// Full-day: all active rows get 0
				for (const row of activeRows) {
					const cellA1 = columnToLetter(col) + row;
					fullDayCells.push(cellA1);
				}
			}
		}
	}

	// Apply colors to full-day leave cells (red)
	if (fullDayCells.length > 0) {
		Logger.log(`Applying red to ${fullDayCells.length} full-day leave cells`);
		const fullDayRanges = sheet.getRangeList(fullDayCells);
		fullDayRanges.setValue(0);
		fullDayRanges.setBackground(CONFIG.COLORS.FULL_DAY);
		fullDayRanges.setFontColor('#FFFFFF');
		fullDayRanges.setFontWeight('bold');
	}

	// Apply colors to half-day leave cells (orange)
	if (halfDayCells.length > 0) {
		Logger.log(
			`Applying orange to ${halfDayCells.length} half-day leave cells`
		);
		for (const { cell, value } of halfDayCells) {
			const range = sheet.getRange(cell);
			range.setValue(value);
			range.setBackground(CONFIG.COLORS.HALF_DAY);
			range.setFontColor('#000000');
			range.setFontWeight('bold');
		}
	}

	Logger.log(
		`Applied leave colors: ${matchedEmployees} employees, ${fullDayCells.length} full-day, ${halfDayCells.length} half-day`
	);
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
						// Determine if this specific day is half-day
						// effective_date_duration: 1=full, 2=AM half, 3=PM half (for first day)
						// end_date_duration: 1=full, 2=AM half, 3=PM half (for last day)
						const isFirstDay = currentDate.getTime() === leaveStart.getTime();
						const isLastDay = currentDate.getTime() === leaveEnd.getTime();
						const isSingleDay = isFirstDay && isLastDay;

						// Parse duration values (API may return string or number)
						const effectiveDuration =
							parseInt(request.effective_date_duration) || 1;
						const endDuration = parseInt(request.end_date_duration) || 1;

						let isHalfDay = false;
						if (isSingleDay) {
							// Single day leave - check effective_date_duration
							// 1 = full day, 2 = AM half, 3 = PM half
							isHalfDay = effectiveDuration === 2 || effectiveDuration === 3;
						} else if (isFirstDay) {
							// First day of multi-day leave
							isHalfDay = effectiveDuration === 2 || effectiveDuration === 3;
						} else if (isLastDay) {
							// Last day of multi-day leave
							isHalfDay = endDuration === 2 || endDuration === 3;
						}
						// Middle days are always full days (isHalfDay = false)

						Logger.log(
							`Leave for ${empName} on day ${currentDate.getDate()}: effectiveDuration=${effectiveDuration}, endDuration=${endDuration}, isHalfDay=${isHalfDay}`
						);

						leaveDays.push({
							date: currentDate.getDate(),
							leave_type: request.time_off?.name,
							is_half_day: isHalfDay,
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
 * Half-day leave: deduct 4 hours total, starting from smallest projects first
 * Respects Week Override column - skips updates for rows with override checked
 * @param {boolean} useSheetHours - If true, read hours from sheet cells instead of attendance data
 */
function updateSheetWithLeaveData(
	sheet,
	leaveData,
	month,
	year,
	useSheetHours = false
) {
	const daysInMonth = new Date(year, month + 1, 0).getDate();

	// Build employee lookup from sheet (now includes hours per row)
	const employeeLookup = buildEmployeeLookup(sheet);

	// Get attendance data to get hours per project (only if not using sheet hours)
	let attendanceByEmployee = {};
	if (!useSheetHours) {
		const attendanceList = getAttendanceData() || [];
		// Build lookup: empId -> [{row, project, hours}, ...]
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
	}

	// Get day columns, validated columns, and week override columns
	const { dayColumns, weekRanges } = calculateDayColumns(month, year);

	// Build a map of day -> week override column for quick lookup
	const dayToOverrideCol = {};
	for (const weekRange of weekRanges) {
		for (const [dayStr, col] of Object.entries(dayColumns)) {
			if (col >= weekRange.startCol && col <= weekRange.endCol) {
				dayToOverrideCol[dayStr] = weekRange.overrideCol;
			}
		}
	}

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
			`Found ${
				rows.length
			} rows for ${employee_name} (ID: ${employee_id}) at rows: ${rows.join(
				', '
			)}`
		);

		// Build row -> hours mapping
		// If useSheetHours is true, read from sheet cells (for "Sync Leave Only")
		// Otherwise, read from attendance data
		const rowHoursMap = {};

		if (useSheetHours) {
			// Read hours from the first weekday column in the sheet for each row
			// Find first weekday column (Monday = day 1 or first day of month that's a weekday)
			const firstDayCol = Object.values(dayColumns)[0];

			for (const row of rows) {
				// Scan the first few day columns to find a valid hour value
				let foundHours = null;
				for (const [dayStr, col] of Object.entries(dayColumns)) {
					const cellValue = sheet.getRange(row, col).getValue();
					// Check if it's a number and not 0 (which could be leave)
					if (typeof cellValue === 'number' && cellValue > 0) {
						foundHours = cellValue;
						break;
					}
				}

				if (foundHours) {
					rowHoursMap[row] = foundHours;
					Logger.log(`Row ${row}: ${foundHours} hours from sheet`);
				} else {
					rowHoursMap[row] = CONFIG.DEFAULT_HOURS;
					Logger.log(
						`Row ${row}: defaulting to ${CONFIG.DEFAULT_HOURS} hours (no valid hours found in sheet)`
					);
				}
			}
		} else {
			// Read from attendance data
			const empAttendance =
				attendanceByEmployee[employee_id.toUpperCase()] || [];

			for (const row of rows) {
				// Get project name from column C (index 3)
				const projectName = String(sheet.getRange(row, 3).getValue() || '')
					.trim()
					.toUpperCase();

				// Find matching attendance entry
				const matchingAtt = empAttendance.find(
					(att) =>
						String(att.project || '')
							.trim()
							.toUpperCase() === projectName
				);

				if (matchingAtt && matchingAtt.hours > 0) {
					rowHoursMap[row] = matchingAtt.hours;
					Logger.log(
						`Row ${row} (${projectName}): ${matchingAtt.hours} hours from attendance`
					);
				} else {
					// If hours is 0 or no attendance match, default to 8 hours
					// (leave is usually booked in advance before hours are allocated)
					rowHoursMap[row] = CONFIG.DEFAULT_HOURS;
					Logger.log(
						`Row ${row} (${projectName}): defaulting to ${CONFIG.DEFAULT_HOURS} hours (no hours allocated yet)`
					);
				}
			}
		}

		// Apply leaves to all rows for this employee
		// Full-day: all rows get 0
		// Half-day: deduct 4 hours total, divided equally across projects
		for (const leave of leave_requests) {
			const col = dayColumns[leave.date];
			if (!col) continue;

			Logger.log(
				`Processing leave for ${employee_name} on day ${leave.date}, is_half_day: ${leave.is_half_day}`
			);

			// Check Week Override for each row - skip rows where override is checked
			const overrideCol = dayToOverrideCol[leave.date];

			// Filter rows that don't have Week Override checked
			const activeRows = rows.filter((row) => {
				if (!overrideCol) {
					Logger.log(
						`Row ${row}: No override column for day ${leave.date}, will update`
					);
					return true;
				}
				const overrideValue = sheet.getRange(row, overrideCol).getValue();
				Logger.log(
					`Row ${row}: Override col ${overrideCol}, value="${overrideValue}" (type: ${typeof overrideValue})`
				);
				if (overrideValue === true) {
					Logger.log(
						`Skipping row ${row} for day ${leave.date} - Week Override is checked`
					);
					return false;
				}
				return true;
			});

			if (activeRows.length === 0) {
				Logger.log(
					`All rows have Week Override checked for day ${leave.date}, skipping`
				);
				continue;
			}

			if (leave.is_half_day) {
				// Half-day leave = 4 hours of work remaining
				// Divide the REMAINING 4 hours EQUALLY across all projects
				const totalRemainingHours = CONFIG.HALF_DAY_HOURS; // 4 hours
				const hoursPerProject = totalRemainingHours / activeRows.length;

				Logger.log(
					`Half-day leave: ${activeRows.length} projects, ${hoursPerProject}hr each (${totalRemainingHours}hr total remaining)`
				);

				// Apply equal remaining hours to each row - all half-day = ORANGE
				for (const row of activeRows) {
					const cellA1 = columnToLetter(col) + row;

					// Half-day leave always gets ORANGE color with equal hours
					halfDayCellsMap[cellA1] = hoursPerProject;
					Logger.log(`Row ${row}: ${hoursPerProject}hr (ORANGE)`);
				}
			} else {
				// Full-day: all active rows get 0
				for (const row of activeRows) {
					const cellA1 = columnToLetter(col) + row;
					fullDayCells.push(cellA1);
				}
			}
		}
	}

	// Apply leave values and colors
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

	const halfDayCellEntries = Object.entries(halfDayCellsMap);
	if (halfDayCellEntries.length > 0) {
		Logger.log(
			`Applying half-day leave to ${halfDayCellEntries.length} cells (orange)`
		);
		for (const [cellA1, value] of halfDayCellEntries) {
			const range = sheet.getRange(cellA1);
			range.setValue(parseFloat(value));
			range.setBackground(CONFIG.COLORS.HALF_DAY);
			range.setFontColor('#000000');
			range.setFontWeight('bold');
		}
	}

	// Add conditional formatting for validated weeks (green)
	// Pass leave cells so they are EXCLUDED from green formatting
	// Also scan sheet for existing leave cells (red/orange backgrounds only) to preserve them
	const newLeaveCells = [...fullDayCells, ...Object.keys(halfDayCellsMap)];
	const existingLeaveCells = scanForLeaveCells(sheet, dayColumns);
	const allLeaveCells = [...new Set([...newLeaveCells, ...existingLeaveCells])];
	Logger.log(
		`Total leave cells to exclude: ${allLeaveCells.length} (${newLeaveCells.length} new, ${existingLeaveCells.length} existing)`
	);
	addValidatedConditionalFormatting(sheet, month, year, allLeaveCells);

	// Final flush
	SpreadsheetApp.flush();

	Logger.log(
		`Matched ${matchedEmployees} employees, updated ${
			fullDayCells.length + halfDayCellEntries.length
		} cells`
	);
}

/**
 * Add conditional formatting for validated weeks only (green)
 * Leave cells are excluded by building ranges that skip them
 * @param {Sheet} sheet - The sheet
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {string[]} leaveCells - Array of cell A1 notations to exclude from green
 */
function addValidatedConditionalFormatting(
	sheet,
	month,
	year,
	leaveCells = []
) {
	// Calculate the range based on month/year (use weekRanges for proper column tracking)
	const { dayColumns, weekRanges } = calculateDayColumns(month, year);

	const lastRow = sheet.getLastRow();

	// Make sure we have valid dimensions
	const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;

	if (numRows <= 0) {
		Logger.log('addValidatedConditionalFormatting: Invalid range dimensions');
		return;
	}

	// Build a Set of leave cells for fast lookup
	const leaveCellSet = new Set(leaveCells.map((c) => c.toUpperCase()));
	Logger.log(
		`Excluding ${leaveCellSet.size} leave cells from green formatting`
	);
	if (leaveCellSet.size > 0) {
		Logger.log(
			`Leave cells: ${Array.from(leaveCellSet).slice(0, 10).join(', ')}${
				leaveCellSet.size > 10 ? '...' : ''
			}`
		);
	}

	// Clear ALL existing conditional formatting rules (start fresh)
	sheet.setConditionalFormatRules([]);

	const rules = [];

	// Green for validated weeks (when checkbox is TRUE)
	// Build ranges that exclude leave cells by grouping consecutive non-leave rows
	for (const weekRange of weekRanges) {
		const { startCol, endCol, validatedCol } = weekRange;
		const checkboxColLetter = columnToLetter(validatedCol);
		const weekdayRanges = [];

		// Find weekday columns in this week (exclude weekends)
		for (const [dayStr, col] of Object.entries(dayColumns)) {
			if (col >= startCol && col <= endCol) {
				const day = parseInt(dayStr);
				const date = new Date(year, month, day);
				const dayOfWeek = date.getDay();
				// Only include weekdays (Mon-Fri: 1-5)
				if (dayOfWeek >= 1 && dayOfWeek <= 5) {
					const colLetter = columnToLetter(col);
					// Group consecutive rows that are NOT leave cells
					let rangeStart = null;
					for (let row = CONFIG.FIRST_DATA_ROW; row <= lastRow + 1; row++) {
						const cellA1 = `${colLetter}${row}`;
						const isLeave = leaveCellSet.has(cellA1.toUpperCase());
						const isLastRow = row > lastRow;

						if (!isLeave && !isLastRow) {
							if (rangeStart === null) rangeStart = row;
						} else {
							if (rangeStart !== null) {
								// End of consecutive range, add it
								const rangeEnd = row - 1;
								weekdayRanges.push(
									sheet.getRange(rangeStart, col, rangeEnd - rangeStart + 1, 1)
								);
								rangeStart = null;
							}
						}
					}
				}
			}
		}

		// Also include the checkbox column itself (all rows)
		weekdayRanges.push(
			sheet.getRange(CONFIG.FIRST_DATA_ROW, validatedCol, numRows, 1)
		);

		if (weekdayRanges.length > 0) {
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

	sheet.setConditionalFormatRules(rules);
	Logger.log(
		`Applied ${rules.length} conditional formatting rules (validated weeks only)`
	);
}

/**
 * Scan sheet for existing leave cells
 * Detects leave ONLY by red/orange background color (not by value)
 * This ensures only API-sourced leave is detected, not manual 0 values
 * Returns array of cell A1 notations
 */
function scanForLeaveCells(sheet, dayColumns) {
	const leaveCells = [];
	const lastRow = sheet.getLastRow();
	const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;

	if (numRows <= 0) return leaveCells;

	// Get all day columns
	const cols = Object.values(dayColumns);
	if (cols.length === 0) return leaveCells;

	const minCol = Math.min(...cols);
	const maxCol = Math.max(...cols);
	const numCols = maxCol - minCol + 1;

	// Read all backgrounds in one batch
	const range = sheet.getRange(CONFIG.FIRST_DATA_ROW, minCol, numRows, numCols);
	const backgrounds = range.getBackgrounds();

	// Leave colors (normalize to uppercase for comparison)
	const fullDayColor = CONFIG.COLORS.FULL_DAY.toUpperCase();
	const halfDayColor = CONFIG.COLORS.HALF_DAY.toUpperCase();

	for (let rowIdx = 0; rowIdx < numRows; rowIdx++) {
		for (let colIdx = 0; colIdx < numCols; colIdx++) {
			const bg = backgrounds[rowIdx][colIdx].toUpperCase();

			// Detect leave ONLY by background color (red or orange)
			if (bg === fullDayColor || bg === halfDayColor) {
				const col = minCol + colIdx;
				const row = CONFIG.FIRST_DATA_ROW + rowIdx;
				const cellA1 = columnToLetter(col) + row;
				leaveCells.push(cellA1);
			}
		}
	}

	Logger.log(`Found ${leaveCells.length} existing leave cells in sheet`);
	return leaveCells;
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
 * Calculate day columns, validated columns, and week override columns for a given month/year
 * Layout: [days...] [Validated] [Week Override] [days...] [Validated] [Week Override] ...
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @returns {Object} { dayColumns: {day: col}, validatedColumns: [cols], weekOverrideColumns: [cols], weekRanges: [{startCol, endCol, validatedCol, overrideCol}] }
 */
function calculateDayColumns(month, year) {
	const daysInMonth = new Date(year, month + 1, 0).getDate();
	const dayColumns = {};
	const validatedColumns = [];
	const weekOverrideColumns = [];
	const weekRanges = [];
	let currentCol = CONFIG.FIRST_DAY_COL;
	let weekStartCol = CONFIG.FIRST_DAY_COL;

	for (let day = 1; day <= daysInMonth; day++) {
		const date = new Date(year, month, day);
		const dayOfWeek = date.getDay();
		dayColumns[day] = currentCol;
		currentCol++;
		if (dayOfWeek === 5) {
			// Friday - add validated column and week override column after
			const validatedCol = currentCol;
			validatedColumns.push(validatedCol);
			currentCol++;
			const overrideCol = currentCol;
			weekOverrideColumns.push(overrideCol);
			currentCol++;

			// Track week range
			weekRanges.push({
				startCol: weekStartCol,
				endCol: validatedCol - 1,
				validatedCol: validatedCol,
				overrideCol: overrideCol,
			});
			weekStartCol = currentCol;
		}
	}

	// Add validated and override columns after the last day if there are weekdays after the last Friday
	const lastDayOfMonth = new Date(year, month, daysInMonth);
	const lastDayOfWeek = lastDayOfMonth.getDay();

	// Only add if last day is Mon-Thu (1-4), meaning there are weekdays not yet validated
	if (lastDayOfWeek >= 1 && lastDayOfWeek <= 4) {
		const validatedCol = currentCol;
		validatedColumns.push(validatedCol);
		currentCol++;
		const overrideCol = currentCol;
		weekOverrideColumns.push(overrideCol);

		weekRanges.push({
			startCol: weekStartCol,
			endCol: validatedCol - 1,
			validatedCol: validatedCol,
			overrideCol: overrideCol,
		});
	}

	return { dayColumns, validatedColumns, weekOverrideColumns, weekRanges };
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
