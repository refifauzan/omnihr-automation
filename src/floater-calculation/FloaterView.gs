/**
 * Floater View - Monthly Floaters & Floater Cost Breakdown
 *
 * Calculates floater percentage for each employee based on their
 * project allocation from the capacity/project sheets.
 * Floater % = percentage of working hours NOT allocated to any project.
 */

/**
 * Generate the Monthly Floaters & Floater Cost Breakdown sheet
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 */
function generateFloaterView(month, year) {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const ui = SpreadsheetApp.getUi();

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

	// Check if sheet already exists
	let sheet = ss.getSheetByName(sheetName);
	if (sheet) {
		const response = ui.alert(
			'Sheet Exists',
			`Sheet "${sheetName}" already exists. Overwrite?`,
			ui.ButtonSet.YES_NO,
		);
		if (response !== ui.Button.YES) return;
		ss.deleteSheet(sheet);
	}

	sheet = ss.insertSheet(sheetName);

	Logger.log(`Generating Floater View for ${monthNames[month]} ${year}`);

	try {
		const token = getAccessToken();
		if (!token) {
			ui.alert('Failed to get API token. Check your credentials.');
			return;
		}

		// Fetch employee data
		const employeesWithDetails = fetchAllEmployeesWithDetails(token);
		Logger.log(`Fetched ${employeesWithDetails.length} employees`);

		// Fetch leave data for this month
		const leaveData = fetchLeaveDataForMonth(
			token,
			employeesWithDetails,
			month,
			year,
		);

		// Fetch holidays
		const holidays = fetchHolidaysForMonth(token, month, year);
		const holidayDays = new Set(holidays.map((h) => h.date));

		// Calculate working days in the month
		const daysInMonth = new Date(year, month + 1, 0).getDate();
		const workingDays = countWorkingDays(month, year, holidayDays);
		Logger.log(`Working days in ${monthNames[month]} ${year}: ${workingDays}`);

		// Read project sheet data to determine allocation
		const allocationData = readProjectSheetAllocation(
			ss,
			month,
			year,
			holidayDays,
			workingDays,
		);

		// Build floater data for each employee
		const floaterData = buildFloaterData(
			employeesWithDetails,
			allocationData,
			leaveData,
			month,
			year,
			holidayDays,
			workingDays,
		);

		// Sort: leavers at bottom, then by floater % descending
		floaterData.sort((a, b) => {
			if (a.isLeaver && !b.isLeaver) return 1;
			if (!a.isLeaver && b.isLeaver) return -1;
			return b.floaterPct - a.floaterPct;
		});

		// Write to sheet
		writeFloaterSheet(sheet, floaterData, monthNames[month], year);

		SpreadsheetApp.flush();

		ui.alert(
			`Floater View generated!\n\n` +
				`Sheet: "${sheetName}"\n` +
				`Employees: ${floaterData.length}\n` +
				`Working days: ${workingDays}`,
		);
	} catch (error) {
		Logger.log(
			'Error generating Floater View: ' + error.message + '\n' + error.stack,
		);
		ui.alert('Error: ' + error.message);
	}
}

/**
 * Count working days in a month (excluding weekends and holidays)
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {Set} holidayDays - Set of holiday day numbers
 * @returns {number} Number of working days
 */
function countWorkingDays(month, year, holidayDays) {
	const daysInMonth = new Date(year, month + 1, 0).getDate();
	let workingDays = 0;

	for (let day = 1; day <= daysInMonth; day++) {
		const date = new Date(year, month, day);
		const dayOfWeek = date.getDay();
		if (dayOfWeek >= 1 && dayOfWeek <= 5 && !holidayDays.has(day)) {
			workingDays++;
		}
	}

	return workingDays;
}

/**
 * Read project sheet data to determine each employee's allocation
 * Looks for sheets that match the month/year pattern and reads hours data
 * @param {Spreadsheet} ss - Spreadsheet
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {Set} holidayDays - Set of holiday day numbers
 * @param {number} workingDays - Total working days
 * @returns {Map} Map of employeeName (lowercase) -> { totalHours, empId, empName, projects }
 */
function readProjectSheetAllocation(ss, month, year, holidayDays, workingDays) {
	const allocationData = new Map();
	const allSheets = ss.getSheets();
	const daysInMonth = new Date(year, month + 1, 0).getDate();

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
	const monthAbbrevs = [
		'Jan',
		'Feb',
		'Mar',
		'Apr',
		'May',
		'Jun',
		'Jul',
		'Aug',
		'Sep',
		'Oct',
		'Nov',
		'Dec',
	];

	// Find project sheets for this month
	const projectSheets = allSheets.filter((s) => {
		const name = s.getName().toLowerCase();
		const monthFull = monthNames[month].toLowerCase();
		const monthAbbrev = monthAbbrevs[month].toLowerCase();
		const yearStr = String(year);

		// Skip floater sheets and capacity view sheets
		if (name.startsWith('floater')) return false;
		if (name.startsWith('capacity')) return false;
		if (name === 'attendance') return false;

		return (
			(name.includes(monthFull) || name.includes(monthAbbrev)) &&
			name.includes(yearStr)
		);
	});

	Logger.log(
		`Found ${projectSheets.length} project sheets for ${monthNames[month]} ${year}`,
	);

	for (const projectSheet of projectSheets) {
		Logger.log(`Reading project sheet: ${projectSheet.getName()}`);

		try {
			const lastRow = projectSheet.getLastRow();
			const lastCol = projectSheet.getLastColumn();
			if (lastRow < 3 || lastCol < 11) continue;

			const numRows = lastRow - 2;

			// Batch read columns A (ID), B (Name), C (Project)
			const metaData = projectSheet.getRange(3, 1, numRows, 3).getValues();

			// Find day columns by reading header row (row 2)
			const headerValues = projectSheet
				.getRange(2, 11, 1, lastCol - 10)
				.getValues()[0];

			const dayColIndices = []; // Array of { day, colIdx } for working days only
			for (let colIdx = 0; colIdx < headerValues.length; colIdx++) {
				const val = headerValues[colIdx];
				if (typeof val === 'number' && val >= 1 && val <= daysInMonth) {
					const date = new Date(year, month, val);
					const dayOfWeek = date.getDay();
					// Only include working days (not weekends or holidays)
					if (dayOfWeek >= 1 && dayOfWeek <= 5 && !holidayDays.has(val)) {
						dayColIndices.push({ day: val, colIdx: colIdx });
					}
				}
			}

			// Batch read all hour values from column K onwards
			const numDataCols = lastCol - 10;
			const allHoursData = projectSheet
				.getRange(3, 11, numRows, numDataCols)
				.getValues();

			for (let i = 0; i < numRows; i++) {
				const empId = String(metaData[i][0] || '').trim();
				const empName = String(metaData[i][1] || '').trim();
				const project = String(metaData[i][2] || '').trim();
				if (!empName) continue;

				const key = empName.toLowerCase();
				if (!allocationData.has(key)) {
					allocationData.set(key, {
						totalHours: 0,
						empId: empId,
						empName: empName,
						projects: new Set(),
					});
				}

				// Track project name from Column C
				if (project) {
					allocationData.get(key).projects.add(project);
				}

				// Sum hours for working days only
				for (const { colIdx } of dayColIndices) {
					const cellValue = allHoursData[i][colIdx];
					const hours = typeof cellValue === 'number' ? cellValue : 0;
					allocationData.get(key).totalHours += hours;
				}
			}
		} catch (e) {
			Logger.log(
				`Error reading project sheet ${projectSheet.getName()}: ${e.message}`,
			);
		}
	}

	Logger.log(`Read allocation data for ${allocationData.size} employees`);
	return allocationData;
}

/**
 * Build floater data for each employee
 * @param {Array} employees - Employee details
 * @param {Map} allocationData - Allocation data from project sheets
 * @param {Map} leaveData - Leave data
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {Set} holidayDays - Holiday day numbers
 * @param {number} workingDays - Total working days in the month
 * @returns {Array} Array of floater data objects
 */
function buildFloaterData(
	employees,
	allocationData,
	leaveData,
	month,
	year,
	holidayDays,
	workingDays,
) {
	const floaterData = [];
	const daysInMonth = new Date(year, month + 1, 0).getDate();
	const maxHours = workingDays * 8;

	for (const emp of employees) {
		const empName = (emp.full_name || '').trim();
		const empNameLower = empName.toLowerCase();
		const empId = String(emp.employee_id || '')
			.trim()
			.toUpperCase();

		// Check if leaver (has termination date in this month or before)
		let isLeaver = false;
		if (emp.termination_date) {
			const termDate = parseDateDMY(emp.termination_date);
			if (termDate) {
				const monthEnd = new Date(year, month + 1, 0);
				if (termDate <= monthEnd) {
					isLeaver = true;
				}
			}
		}

		// Get allocation from project sheets
		const allocation = allocationData.get(empNameLower);
		const totalAllocatedHours = allocation ? allocation.totalHours : 0;

		// Get leave days count
		let leaveDays = 0;
		const empLeave = leaveData.get(empId) || leaveData.get(empNameLower);
		if (empLeave) {
			for (const [day, info] of empLeave) {
				leaveDays += info.is_half_day ? 0.5 : 1;
			}
		}

		// Calculate effective working days (exclude leave)
		const effectiveWorkingDays = workingDays - leaveDays;
		const effectiveMaxHours = effectiveWorkingDays * 8;

		// Calculate floater percentage
		// Floater % = (unallocated hours / effective max hours) * 100
		let floaterPct = 0;
		if (effectiveMaxHours > 0) {
			const unallocatedHours = Math.max(
				0,
				effectiveMaxHours - totalAllocatedHours,
			);
			floaterPct = (unallocatedHours / effectiveMaxHours) * 100;
		}

		// If leaver, set floater to 100%
		if (isLeaver) {
			floaterPct = 100;
		}

		// Calculate floater cost
		const floaterCost = (floaterPct / 100) * CONFIG.AVERAGE_SALARY;

		// Get department/team
		const department = emp.team || '';

		// Get current project from allocation data (Column C of project sheets)
		let currentProject = '';
		if (allocation && allocation.projects && allocation.projects.size > 0) {
			currentProject = [...allocation.projects].join(', ');
		}

		floaterData.push({
			name: empName,
			department: department,
			floaterPct: Math.round(floaterPct * 100) / 100,
			floaterCost: Math.round(floaterCost),
			currentProject: currentProject,
			isLeaver: isLeaver,
			totalAllocatedHours: totalAllocatedHours,
			effectiveMaxHours: effectiveMaxHours,
			leaveDays: leaveDays,
		});
	}

	return floaterData;
}

/**
 * Write floater data to the sheet
 * @param {Sheet} sheet - Target sheet
 * @param {Array} floaterData - Array of floater data objects
 * @param {string} monthName - Month name
 * @param {number} year - Year
 */
function writeFloaterSheet(sheet, floaterData, monthName, year) {
	// Title row
	sheet.getRange(CONFIG.TITLE_ROW, 1, 1, 4).merge();
	sheet
		.getRange(CONFIG.TITLE_ROW, 1)
		.setValue('Monthly Floaters & Floater Cost Breakdown');
	sheet.getRange(CONFIG.TITLE_ROW, 1).setFontSize(14).setFontWeight('bold');

	// Month row
	sheet.getRange(CONFIG.MONTH_ROW, 1).setValue(monthName);
	sheet.getRange(CONFIG.MONTH_ROW, 1).setFontSize(11).setFontWeight('bold');

	// Headers
	const headers = ['Name', 'Department', 'Floater %', 'Current Project'];
	sheet.getRange(CONFIG.HEADER_ROW, 1, 1, headers.length).setValues([headers]);
	sheet
		.getRange(CONFIG.HEADER_ROW, 1, 1, headers.length)
		.setBackground(CONFIG.HEADER_BG)
		.setFontColor(CONFIG.HEADER_FONT_COLOR)
		.setFontWeight('bold')
		.setHorizontalAlignment('center')
		.setBorder(
			true,
			true,
			true,
			true,
			true,
			true,
			'#000000',
			SpreadsheetApp.BorderStyle.SOLID,
		);

	// Write data rows
	if (floaterData.length > 0) {
		const dataRows = floaterData.map((emp) => [
			emp.name,
			emp.department,
			emp.floaterPct / 100, // Store as decimal for percentage formatting
			emp.currentProject,
		]);

		const dataRange = sheet.getRange(
			CONFIG.FIRST_DATA_ROW,
			1,
			dataRows.length,
			headers.length,
		);
		dataRange.setValues(dataRows);

		// Format Floater % column as percentage
		sheet
			.getRange(
				CONFIG.FIRST_DATA_ROW,
				CONFIG.FLOATER_PCT_COL,
				dataRows.length,
				1,
			)
			.setNumberFormat('0.0%')
			.setHorizontalAlignment('center');

		// Apply conditional scale colors based on floater cost
		for (let i = 0; i < floaterData.length; i++) {
			const row = CONFIG.FIRST_DATA_ROW + i;
			const emp = floaterData[i];
			const rowRange = sheet.getRange(row, 1, 1, headers.length);

			let scale;
			if (emp.isLeaver) {
				scale = CONFIG.SCALES.LEAVERS;
			} else if (emp.floaterCost >= 10000) {
				scale = CONFIG.SCALES.ABOVE_10K;
			} else if (emp.floaterCost >= 7000) {
				scale = CONFIG.SCALES.FROM_7K_TO_10K;
			} else if (emp.floaterCost >= 4000) {
				scale = CONFIG.SCALES.FROM_4K_TO_7K;
			} else {
				scale = CONFIG.SCALES.BELOW_4K;
			}

			rowRange.setBackground(scale.color);
			rowRange.setFontColor(scale.fontColor);
		}

		// Add borders to data area
		sheet
			.getRange(CONFIG.FIRST_DATA_ROW, 1, dataRows.length, headers.length)
			.setBorder(
				true,
				true,
				true,
				true,
				true,
				true,
				'#000000',
				SpreadsheetApp.BorderStyle.SOLID,
			);
	}

	// Write Conditional Scales legend
	writeLegend(sheet);

	// Set column widths
	sheet.setColumnWidth(CONFIG.NAME_COL, 200);
	sheet.setColumnWidth(CONFIG.DEPARTMENT_COL, 150);
	sheet.setColumnWidth(CONFIG.FLOATER_PCT_COL, 100);
	sheet.setColumnWidth(CONFIG.CURRENT_PROJECT_COL, 200);

	// Freeze header rows
	sheet.setFrozenRows(CONFIG.HEADER_ROW);
}

/**
 * Write the Conditional Scales legend on the right side of the sheet
 * @param {Sheet} sheet - Target sheet
 */
function writeLegend(sheet) {
	const legendStartRow = CONFIG.MONTH_ROW;
	const labelCol = CONFIG.LEGEND_LABEL_COL;
	const colorCol = CONFIG.LEGEND_COLOR_COL;

	// Legend title
	sheet.getRange(legendStartRow, labelCol, 1, 2).merge();
	sheet
		.getRange(legendStartRow, labelCol)
		.setValue('Conditional Scales')
		.setFontWeight('bold')
		.setFontSize(11);

	// Legend items
	const scales = [
		CONFIG.SCALES.ABOVE_10K,
		CONFIG.SCALES.FROM_7K_TO_10K,
		CONFIG.SCALES.FROM_4K_TO_7K,
		CONFIG.SCALES.BELOW_4K,
		CONFIG.SCALES.LEAVERS,
	];

	for (let i = 0; i < scales.length; i++) {
		const row = legendStartRow + 1 + i;
		sheet.getRange(row, labelCol).setValue(scales[i].label);
		sheet.getRange(row, colorCol).setBackground(scales[i].color);
		sheet
			.getRange(row, labelCol, 1, 2)
			.setBorder(
				true,
				true,
				true,
				true,
				true,
				true,
				'#000000',
				SpreadsheetApp.BorderStyle.SOLID,
			);
	}

	// Set legend column widths
	sheet.setColumnWidth(labelCol, 120);
	sheet.setColumnWidth(colorCol, 80);
}

/**
 * Update existing floater view with latest data
 */
function updateFloaterView() {
	const ui = SpreadsheetApp.getUi();
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const activeSheet = ss.getActiveSheet();
	const sheetName = activeSheet.getName();

	// Check if this is a Floater sheet
	if (!sheetName.startsWith('Floaters')) {
		ui.alert('Please navigate to a Floaters sheet first.');
		return;
	}

	// Parse month/year from sheet name (e.g., "Floaters February 2026")
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

	let month = -1;
	let year = -1;

	for (let i = 0; i < monthNames.length; i++) {
		if (sheetName.includes(monthNames[i])) {
			month = i;
			break;
		}
	}

	const yearMatch = sheetName.match(/(\d{4})/);
	if (yearMatch) {
		year = parseInt(yearMatch[1]);
	}

	if (month === -1 || year === -1) {
		ui.alert('Could not determine month/year from sheet name.');
		return;
	}

	// Regenerate with the same month/year
	generateFloaterView(month, year);
}
