/**
 * Attendance sheet functions
 */

/**
 * Create the Attendance sheet and fetch all employees from OmniHR
 */
function createAttendanceSheet() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const ui = SpreadsheetApp.getUi();

	let sheet = ss.getSheetByName(CONFIG.ATTENDANCE_SHEET_NAME);
	if (sheet) {
		const response = ui.alert(
			'Attendance sheet already exists!',
			'Do you want to refresh employee data from OmniHR?',
			ui.ButtonSet.YES_NO
		);
		if (response !== ui.Button.YES) return;

		const lastRow = sheet.getLastRow();
		if (lastRow > 1) {
			sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
		}
	} else {
		sheet = ss.insertSheet(CONFIG.ATTENDANCE_SHEET_NAME);

		const headers = ['ID', 'Employee Name', 'Project', 'Type', 'Daily Hours'];
		sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
		sheet
			.getRange(1, 1, 1, headers.length)
			.setFontWeight('bold')
			.setBackground('#4285f4')
			.setFontColor('#FFFFFF');

		const typeRule = SpreadsheetApp.newDataValidation()
			.requireValueInList(['Full-time', 'Part-time', 'Custom'], true)
			.setAllowInvalid(false)
			.build();
		sheet.getRange('D2:D1000').setDataValidation(typeRule);

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
						'',
						'Full-time',
						8,
					]);
				} catch (e) {
					const emp = batch[j];
					employeeData.push([
						'',
						emp.full_name || emp.name || `User ${emp.id}`,
						'',
						'Full-time',
						8,
					]);
				}
			}
		}

		if (employeeData.length > 0) {
			sheet.getRange(2, 1, employeeData.length, 5).setValues(employeeData);
		}

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
 * Get attendance data as an array
 * @returns {Array|null} Attendance list or null if sheet not found
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
			5
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
		let hours = !isNaN(rawHours) ? rawHours : CONFIG.DEFAULT_HOURS;

		if (!empId && !empName) continue;

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
 * Apply attendance data to the current sheet
 */
function applyAttendanceToCurrentSheet() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getActiveSheet();

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
 * @param {Sheet} sheet - The sheet
 * @param {Array} attendanceList - Attendance data
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {boolean} autoValidate - Whether to automatically set Validated checkboxes to true (default: true)
 */
function applyAttendanceHours(
	sheet,
	attendanceList,
	month,
	year,
	autoValidate = true
) {
	const daysInMonth = new Date(year, month + 1, 0).getDate();
	const lastRow = sheet.getLastRow();
	const dayNames = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];

	const TOTAL_HOURS_COL = 7;
	const TOTAL_DAYS_COL = 8;
	const TOTAL_DAYS_OFF_COL = 9;

	Logger.log(
		`Applying attendance for ${
			month + 1
		}/${year}, days in month: ${daysInMonth}`
	);
	Logger.log(`Main sheet last row: ${lastRow}`);
	Logger.log(`Attendance list has ${attendanceList.length} rows`);

	// Build lookup map
	const attendanceLookup = {};
	for (let i = 0; i < attendanceList.length; i++) {
		const att = attendanceList[i];
		const key = `${att.empId}|${att.project.toUpperCase()}`;
		attendanceLookup[key] = att;
	}

	const { dayColumns, validatedColumns, weekOverrideColumns } =
		calculateDayColumns(month, year);

	const lastDayCol = Math.max(
		...Object.values(dayColumns),
		...validatedColumns,
		...weekOverrideColumns
	);
	const totalCols = lastDayCol - CONFIG.FIRST_DAY_COL + 1;

	const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;
	if (numRows <= 0) {
		Logger.log('No data rows found in sheet');
		return;
	}

	const employeeData = sheet
		.getRange(CONFIG.FIRST_DATA_ROW, 1, numRows, CONFIG.PROJECT_COL)
		.getValues();

	// Save override states
	const savedOverrideStates = {};
	const savedHourValues = {};
	const sheetLastCol = sheet.getLastColumn();

	if (sheetLastCol >= CONFIG.FIRST_DAY_COL) {
		const headerRange = sheet.getRange(
			CONFIG.HEADER_ROW,
			CONFIG.FIRST_DAY_COL,
			1,
			sheetLastCol - CONFIG.FIRST_DAY_COL + 1
		);
		const headerValues = headerRange.getValues()[0];

		const overrideColIndices = [];
		for (let i = 0; i < headerValues.length; i++) {
			if (
				headerValues[i] === 'Override' ||
				headerValues[i] === 'Time off Override'
			) {
				overrideColIndices.push(i);
				const overrideCol = CONFIG.FIRST_DAY_COL + i;

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

							for (let dayOffset = 1; dayOffset <= 7; dayOffset++) {
								const dayColIdx = i - dayOffset;
								if (dayColIdx >= 0) {
									const dayHeader = headerValues[dayColIdx];
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
										break;
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

	// Delete and recreate columns
	Logger.log('Deleting existing columns from K onwards...');

	if (sheetLastCol >= CONFIG.FIRST_DAY_COL) {
		const numColsToDelete = sheetLastCol - CONFIG.FIRST_DAY_COL + 1;
		Logger.log(
			`Deleting ${numColsToDelete} columns starting from column ${CONFIG.FIRST_DAY_COL}`
		);
		sheet.deleteColumns(CONFIG.FIRST_DAY_COL, numColsToDelete);
	}

	Logger.log(`Inserting ${totalCols} fresh columns...`);
	sheet.insertColumnsAfter(CONFIG.FIRST_DAY_COL - 1, totalCols);
	SpreadsheetApp.flush();

	// Clear conditional formatting
	const existingRules = sheet.getConditionalFormatRules();
	const filteredRules = existingRules.filter((rule) => {
		try {
			const bg = rule.getBooleanCondition();
			if (bg) {
				const condition = bg.getCriteriaType();
				if (condition === SpreadsheetApp.BooleanCriteria.NUMBER_EQUAL_TO) {
					const values = bg.getCriteriaValues();
					if (values && (values[0] === 0 || values[0] === 4)) {
						return false;
					}
				}
			}
		} catch (e) {}
		return true;
	});
	sheet.setConditionalFormatRules(filteredRules);

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
	Logger.log('Setting up headers...');
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

	const headerRange1 = sheet.getRange(
		CONFIG.DAY_NAME_ROW,
		CONFIG.FIRST_DAY_COL,
		1,
		totalCols
	);
	headerRange1.setFontColors([fontColors]);

	// Set column widths
	Logger.log('Setting column widths...');
	for (let col = CONFIG.FIRST_DAY_COL; col <= lastDayCol; col++) {
		if (validatedColumns.includes(col) || weekOverrideColumns.includes(col)) {
			sheet.setColumnWidth(col, 105);
		} else {
			sheet.setColumnWidth(col, 46);
		}
	}

	// Build data arrays
	Logger.log('Building data arrays...');
	const dataValues = [];
	const backgroundColors = [];
	const formulas = [];
	let matchedCount = 0;

	for (let i = 0; i < employeeData.length; i++) {
		const empId = String(employeeData[i][0]).trim().toUpperCase();
		const project = String(employeeData[i][CONFIG.PROJECT_COL - 1])
			.trim()
			.toUpperCase();

		if (!empId) {
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
				rowValues.push(true);
				rowColors.push(null);
			} else if (weekOverrideColumns.includes(col)) {
				rowValues.push(false);
				rowColors.push(null);
			} else {
				rowValues.push('');
				rowColors.push(null);
			}
		}

		dataValues.push(rowValues);
		backgroundColors.push(rowColors);

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

	// Apply data
	Logger.log('Applying data values...');
	const dataRange = sheet.getRange(
		CONFIG.FIRST_DATA_ROW,
		CONFIG.FIRST_DAY_COL,
		numRows,
		totalCols
	);
	dataRange.setValues(dataValues);

	Logger.log('Applying background colors...');
	dataRange.setBackgrounds(backgroundColors);

	// Setup checkboxes
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
			if (autoValidate) {
				checkboxRange.setValue(true);
			} else {
				checkboxRange.setValue(false);
			}
		} catch (e) {
			Logger.log(`Validated checkbox column ${col}: ${e.message}`);
		}
	}

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
			checkboxRange.setValue(false);

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

	// Restore saved hour values
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

	// Apply formulas
	Logger.log('Applying formulas...');
	sheet
		.getRange(CONFIG.FIRST_DATA_ROW, TOTAL_HOURS_COL, numRows, 3)
		.setFormulas(formulas);

	// Add conditional formatting
	Logger.log('Adding conditional formatting...');
	addValidatedConditionalFormatting(sheet, month, year);

	Logger.log(`Matched ${matchedCount} rows with attendance data`);
	SpreadsheetApp.flush();
	Logger.log('Attendance hours applied successfully');
}
