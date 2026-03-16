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
	let isUpdate = false;
	if (sheet) {
		const response = ui.alert(
			'Sheet Exists',
			`Sheet "${sheetName}" already exists. Update columns A-E only?`,
			ui.ButtonSet.YES_NO,
		);
		if (response !== ui.Button.YES) return;
		isUpdate = true;
	} else {
		sheet = ss.insertSheet(sheetName);
	}

	Logger.log(`Generating Floater View for ${monthNames[month]} ${year}`);

	try {
		const token = getAccessToken();
		if (!token) {
			ui.alert('Failed to get API token. Check your credentials.');
			return;
		}

		// Fetch employee data from API (for department + termination dates)
		const employeesWithDetails = fetchAllEmployeesWithDetails(token);
		Logger.log(`Fetched ${employeesWithDetails.length} employees`);

		// Fetch holidays (to calculate working days for denominator)
		const holidays = fetchHolidaysForMonth(token, month, year);
		const holidayDays = new Set(holidays.map((h) => h.date));
		const leaveData = fetchLeaveDataForMonth(
			token,
			employeesWithDetails,
			month,
			year,
		);

		// Calculate working days in the month
		const workingDays = countWorkingDays(month, year, holidayDays);
		Logger.log(`Working days in ${monthNames[month]} ${year}: ${workingDays}`);

		// Read project attendance data from source spreadsheet (read-only)
		// Uses "[Month] [Year]" sheet (e.g. "February 2026"), NOT "CV [Month] [Year]"
		const cvData = readCapacityViewData(
			month,
			year,
			employeesWithDetails,
			leaveData,
			holidayDays,
		);
		Logger.log(`Read project data for ${cvData.size} employees`);

		// Build floater data by merging API data (department, termination) with project sheet data (free hours, projects)
		const floaterData = buildFloaterData(
			employeesWithDetails,
			cvData,
			month,
			year,
			workingDays,
			holidayDays,
		);

		// Only include employees with floater % > 0
		const floatersOnly = floaterData.filter((emp) => emp.floaterPct > 0);

		// Sort: leavers at bottom, then by floater % descending
		floatersOnly.sort((a, b) => {
			if (a.isLeaver && !b.isLeaver) return 1;
			if (!a.isLeaver && b.isLeaver) return -1;
			return b.floaterPct - a.floaterPct;
		});

		// Write to sheet (update only columns A-E if sheet already existed)
		writeFloaterSheet(sheet, floatersOnly, monthNames[month], year, isUpdate);

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

function buildProjectSheetDayColumns(month, year) {
	const daysInMonth = new Date(year, month + 1, 0).getDate();
	const dayColumns = {};
	let currentCol = 11;

	for (let day = 1; day <= daysInMonth; day++) {
		const date = new Date(year, month, day);
		dayColumns[day] = currentCol;
		currentCol++;
		if (date.getDay() === 5) {
			currentCol += 2;
		}
	}

	return dayColumns;
}

function hasFloaterTag(value) {
	return String(value || '')
		.trim()
		.toLowerCase()
		.includes('floater');
}

function buildEmployeeDetailsLookup(employees) {
	const employeeDetailsLookup = new Map();

	for (const emp of employees || []) {
		const empId = String(emp.employee_id || '')
			.trim()
			.toUpperCase();
		const empName = String(emp.full_name || emp.name || '')
			.trim()
			.toLowerCase();

		if (empId) {
			employeeDetailsLookup.set(empId, emp);
		}
		if (empName) {
			employeeDetailsLookup.set(empName, emp);
		}
	}

	return employeeDetailsLookup;
}

function resolveEmployeeDetails(employeeDetailsLookup, empId, empNameLower) {
	return (
		(employeeDetailsLookup && empId && employeeDetailsLookup.get(empId)) ||
		(employeeDetailsLookup &&
			empNameLower &&
			employeeDetailsLookup.get(empNameLower)) ||
		null
	);
}

function normalizeDateOnly(date) {
	if (!date) return null;
	return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function isEmployeeActiveOnDay(employeeDetails, year, month, day) {
	if (!employeeDetails) return true;

	const currentDate = new Date(year, month, day);
	const hireDate = employeeDetails.hired_date
		? normalizeDateOnly(parseDateDMY(employeeDetails.hired_date))
		: null;
	const terminationDate = employeeDetails.termination_date
		? normalizeDateOnly(parseDateDMY(employeeDetails.termination_date))
		: null;

	if (hireDate && currentDate < hireDate) {
		return false;
	}

	if (terminationDate && currentDate > terminationDate) {
		return false;
	}

	return true;
}

function mergeLeaveInfo(sourceLeaveInfo, apiLeaveInfo) {
	if (!sourceLeaveInfo && !apiLeaveInfo) {
		return null;
	}

	if (
		(sourceLeaveInfo && !sourceLeaveInfo.is_half_day) ||
		(apiLeaveInfo && !apiLeaveInfo.is_half_day)
	) {
		return { is_half_day: false };
	}

	return { is_half_day: true };
}

function countEmployeeWorkingDays(employeeDetails, month, year, holidayDays) {
	const daysInMonth = new Date(year, month + 1, 0).getDate();
	let totalWorkingDays = 0;

	for (let day = 1; day <= daysInMonth; day++) {
		const date = new Date(year, month, day);
		const dayOfWeek = date.getDay();
		if (dayOfWeek === 0 || dayOfWeek === 6 || holidayDays.has(day)) {
			continue;
		}

		if (!isEmployeeActiveOnDay(employeeDetails, year, month, day)) {
			continue;
		}

		totalWorkingDays++;
	}

	return totalWorkingDays;
}

function findMonthlyProjectColumn(sheet, firstDayCol) {
	const metaColumnCount = Math.max(firstDayCol - 1, 0);
	if (metaColumnCount === 0) {
		return null;
	}

	const headerRow1 = sheet.getRange(1, 1, 1, metaColumnCount).getValues()[0];
	const headerRow2 = sheet.getRange(2, 1, 1, metaColumnCount).getValues()[0];

	for (let col = 1; col <= metaColumnCount; col++) {
		const header1 = String(headerRow1[col - 1] || '')
			.trim()
			.toLowerCase();
		const header2 = String(headerRow2[col - 1] || '')
			.trim()
			.toLowerCase();
		const combinedHeader = `${header1} ${header2}`.trim();

		if (combinedHeader.includes('current project')) {
			return col;
		}
	}

	for (let col = 1; col <= metaColumnCount; col++) {
		const header1 = String(headerRow1[col - 1] || '')
			.trim()
			.toLowerCase();
		const header2 = String(headerRow2[col - 1] || '')
			.trim()
			.toLowerCase();
		const combinedHeader = `${header1} ${header2}`.trim();

		if (
			combinedHeader === 'project' ||
			combinedHeader.endsWith(' project') ||
			combinedHeader.startsWith('project ')
		) {
			if (!combinedHeader.includes('project contribution')) {
				return col;
			}
		}
	}

	return null;
}

/**
 * Read project attendance data from the source spreadsheet (read-only)
 * Opens the spreadsheet and reads the "[Month] [Year]" sheet (e.g. "February 2026").
 * Does NOT use "CV [Month] [Year]" sheets.
 *
 * Sheet structure (project attendance):
 *   Row 1: Day names (S, M, T, W, T, F, S)
 *   Row 2: Headers - ID | Name | Team | 1 | 2 | ... | 31 | Validated | etc.
 *   Row 3+: Data rows
 *
 * Each day cell = hours allocated. Column C = Team/Project.
 *
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @returns {Map} Map of empId -> { empId, empName, projects, totalFreeHours, totalOverHours }
 */

function readCapacityViewData(month, year, employees, leaveData, holidayDays) {
	const cvData = new Map();
	const uniqueEntries = [];
	const fullDayLeaveColor = '#ff0000';
	const halfDayLeaveColor = '#ffa500';
	const sourceFirstDataRow = 3;
	const sourceFirstDayCol = 11;
	const employeeDetailsLookup = buildEmployeeDetailsLookup(employees);

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

	// Open the source spreadsheet (read-only)
	let sourceSS;
	try {
		sourceSS = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
	} catch (e) {
		Logger.log('Error opening source spreadsheet: ' + e.message);
		return cvData;
	}

	// Read from "[Month] [Year]" sheet only (e.g. "February 2026"), NOT "CV [Month] [Year]"
	const sheetName = `${monthNames[month]} ${year}`;
	const projectSheet = sourceSS.getSheetByName(sheetName);

	if (!projectSheet) {
		Logger.log(`No sheet found with name: "${sheetName}"`);
		return cvData;
	}

	Logger.log(`Reading project sheet: ${projectSheet.getName()} (read-only)`);

	try {
		const lastRow = projectSheet.getLastRow();
		if (
			lastRow < sourceFirstDataRow ||
			projectSheet.getLastColumn() < sourceFirstDayCol
		) {
			Logger.log('Project sheet has insufficient data');
			return cvData;
		}

		const daysInMonth = new Date(year, month + 1, 0).getDate();
		const dayColumns = buildProjectSheetDayColumns(month, year);
		const normalizedHolidayDays = holidayDays || new Set();
		const numRows = lastRow - sourceFirstDataRow + 1;
		const sourceProjectColumn = findMonthlyProjectColumn(
			projectSheet,
			sourceFirstDayCol,
		);
		const employeeData = projectSheet
			.getRange(sourceFirstDataRow, 1, numRows, sourceFirstDayCol - 1)
			.getValues();
		const hoursData = {};
		const backgroundData = {};

		for (let day = 1; day <= daysInMonth; day++) {
			const col = dayColumns[day];
			const range = projectSheet.getRange(sourceFirstDataRow, col, numRows, 1);
			hoursData[day] = range.getValues().map((row) => row[0]);
			backgroundData[day] = range
				.getBackgrounds()
				.map((row) => String(row[0] || '').toLowerCase());
		}

		for (let i = 0; i < numRows; i++) {
			const empId = String(employeeData[i][0] || '')
				.trim()
				.toUpperCase();
			const empName = String(employeeData[i][1] || '').trim();
			// Column C = Team, Column D = Project Contribution (google-appscript structure)
			const sourceColC = String(employeeData[i][2] || '').trim();
			const sourceColD = String(employeeData[i][3] || '').trim();
			const currentProject = sourceProjectColumn
				? String(employeeData[i][sourceProjectColumn - 1] || '').trim()
				: sourceColD || sourceColC || '';
			const empNameLower = empName.toLowerCase();

			if (!empId && !empName) continue;

			let entry = cvData.get(empId) || cvData.get(empNameLower);
			if (!entry) {
				const perDayRegularHours = {};
				for (let day = 1; day <= daysInMonth; day++) {
					perDayRegularHours[day] = 0;
				}

				entry = {
					empId: empId,
					empName: empName,
					projects: new Set(),
					totalFreeHours: 0,
					totalOverHours: 0,
					maxHours: 0,
					perDayRegularHours: perDayRegularHours,
					perDaySourceLeave: new Map(),
					employeeDetails: resolveEmployeeDetails(
						employeeDetailsLookup,
						empId,
						empNameLower,
					),
				};
				uniqueEntries.push(entry);
			}

			if (!entry.empId && empId) {
				entry.empId = empId;
			}
			if (!entry.empName && empName) {
				entry.empName = empName;
			}
			if (!entry.employeeDetails) {
				entry.employeeDetails = resolveEmployeeDetails(
					employeeDetailsLookup,
					empId,
					empNameLower,
				);
			}

			if (empId) {
				cvData.set(empId, entry);
			}
			if (empName) {
				cvData.set(empNameLower, entry);
			}
			if (currentProject) {
				entry.projects.add(currentProject);
			}

			// Only treat as floater assignment if the project is EXACTLY "Floater"
			// Mixed labels like "Atlas 2, Floater" should count their hours
			const isFloaterAssignment =
				currentProject.toLowerCase() === 'floater' ||
				sourceColC.toLowerCase() === 'floater' ||
				sourceColD.toLowerCase() === 'floater';

			for (let day = 1; day <= daysInMonth; day++) {
				const date = new Date(year, month, day);
				const dayOfWeek = date.getDay();
				if (
					dayOfWeek === 0 ||
					dayOfWeek === 6 ||
					normalizedHolidayDays.has(day)
				) {
					continue;
				}

				if (!isEmployeeActiveOnDay(entry.employeeDetails, year, month, day)) {
					continue;
				}

				const background = backgroundData[day][i];
				if (
					background === fullDayLeaveColor ||
					background === halfDayLeaveColor
				) {
					const existingLeaveInfo = entry.perDaySourceLeave.get(day);
					if (
						!existingLeaveInfo ||
						(existingLeaveInfo.is_half_day && background === fullDayLeaveColor)
					) {
						entry.perDaySourceLeave.set(day, {
							is_half_day: background === halfDayLeaveColor,
						});
					}
					continue;
				}

				if (isFloaterAssignment) {
					continue;
				}

				const rawHours = hoursData[day][i];
				const validHours = typeof rawHours === 'number' ? rawHours : 0;
				entry.perDayRegularHours[day] += validHours;
			}
		}

		for (const entry of uniqueEntries) {
			const entryLeaveDays =
				(leaveData &&
					(entry.empId
						? leaveData.get(entry.empId)
						: leaveData.get(String(entry.empName || '').toLowerCase()))) ||
				(leaveData && leaveData.get(String(entry.empName || '').toLowerCase()));

			for (let day = 1; day <= daysInMonth; day++) {
				const date = new Date(year, month, day);
				const dayOfWeek = date.getDay();
				if (
					dayOfWeek === 0 ||
					dayOfWeek === 6 ||
					normalizedHolidayDays.has(day)
				) {
					continue;
				}

				if (!isEmployeeActiveOnDay(entry.employeeDetails, year, month, day)) {
					continue;
				}

				const sourceLeaveInfo = entry.perDaySourceLeave.get(day);
				const apiLeaveInfo = entryLeaveDays && entryLeaveDays.get(day);
				const leaveInfo = mergeLeaveInfo(sourceLeaveInfo, apiLeaveInfo);
				if (leaveInfo) {
					continue;
				}
				const standardHours = 8;
				const assignedHours = entry.perDayRegularHours[day] || 0;
				entry.maxHours += standardHours;

				entry.totalFreeHours += Math.max(0, standardHours - assignedHours);
				entry.totalOverHours += Math.max(0, assignedHours - standardHours);
			}
		}
	} catch (e) {
		Logger.log(`Error reading project sheet: ${e.message}`);
	}

	Logger.log(`Read project data for ${cvData.size} employee keys`);
	return cvData;
}

/**
 * Build floater data by merging API employee details with project sheet data
 *
 * Floater % = (Total Free Hours) / (working days * 8) * 100
 * Free hours are computed from the [Month] [Year] project sheet (8 - allocated hours).
 *
 * @param {Array} employees - Employee details from API (department, termination)
 * @param {Map} cvData - Project sheet data (totalFreeHours, projects from "[Month] [Year]")
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {number} workingDays - Total working days in the month
 * @param {Set} holidayDays - Set of holiday day numbers
 * @returns {Array} Array of floater data objects
 */
function buildFloaterData(
	employees,
	cvData,
	month,
	year,
	workingDays,
	holidayDays,
) {
	const floaterData = [];

	for (const emp of employees) {
		// Exclude Operations team and other excluded teams from floater list
		const team = String(emp.team || '').trim().toLowerCase();
		const isExcludedTeam = CONFIG.EXCLUDED_TEAMS.some((excluded) =>
			team.includes(excluded.toLowerCase()),
		);
		if (isExcludedTeam) {
			Logger.log(`Excluding from floater list (team): ${emp.full_name} - ${team}`);
			continue;
		}

		const empName = (emp.full_name || '').trim();
		const empNameLower = empName.toLowerCase();
		const cvEntry =
			cvData.get(String(emp.employee_id || '').trim().toUpperCase()) ||
			cvData.get(empNameLower);

		// Only include employees that exist in the attendance sheet (source of truth)
		if (!cvEntry) {
			continue;
		}

		// Employee ID: ONLY from attendance sheet
		const empId = cvEntry.empId
			? String(cvEntry.empId).trim().toUpperCase()
			: '';

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

		const employeeMaxHours = cvEntry.maxHours;
		const totalFreeHours = cvEntry.totalFreeHours;
		const totalOverHours = cvEntry.totalOverHours || 0;

		// Floater % = max(0, (free hours - over hours) / max hours) * 100
		let floaterPct = 0;
		if (employeeMaxHours > 0) {
			const adjustedFreeHours = Math.max(0, totalFreeHours - totalOverHours);
			floaterPct = (adjustedFreeHours / employeeMaxHours) * 100;
		}

		const floaterCost = (floaterPct / 100) * CONFIG.AVERAGE_SALARY;

		const department = emp.department || '';

		let currentProject = '';
		if (cvEntry.projects && cvEntry.projects.size > 0) {
			currentProject = [...cvEntry.projects].join(', ');
		} else if (floaterPct >= 100) {
			currentProject = 'Floater';
		}

		floaterData.push({
			employeeId: empId,
			name: empName,
			department: department,
			floaterPct: Math.round(floaterPct * 100) / 100,
			floaterCost: Math.round(floaterCost),
			currentProject: currentProject,
			isLeaver: isLeaver,
			totalFreeHours: totalFreeHours,
			totalOverHours: totalOverHours,
			maxHours: employeeMaxHours,
		});
	}

	return floaterData;
}

/**
 * Write floater data to the sheet
 * When isUpdate=true, only clears and rewrites columns A-E (data area),
 * leaving any content beyond column E untouched.
 * @param {Sheet} sheet - Target sheet
 * @param {Array} floaterData - Array of floater data objects
 * @param {string} monthName - Month name
 * @param {number} year - Year
 * @param {boolean} [isUpdate=false] - If true, only update columns A-E data rows
 */
function writeFloaterSheet(sheet, floaterData, monthName, year, isUpdate) {
	const headers = [
		'Employee ID',
		'Name',
		'Department',
		'Floater %',
		'Current Project',
	];

	if (isUpdate) {
		// Clear only the data area in columns A-E (preserve everything else)
		const lastRow = sheet.getLastRow();
		if (lastRow >= CONFIG.FIRST_DATA_ROW) {
			sheet
				.getRange(
					CONFIG.FIRST_DATA_ROW,
					1,
					lastRow - CONFIG.FIRST_DATA_ROW + 1,
					CONFIG.DATA_COLS,
				)
				.clearContent()
				.clearFormat();
		}
	} else {
		// Title row
		sheet.getRange(CONFIG.TITLE_ROW, 1, 1, CONFIG.DATA_COLS).merge();
		sheet
			.getRange(CONFIG.TITLE_ROW, 1)
			.setValue('Monthly Floaters & Floater Cost Breakdown');
		sheet.getRange(CONFIG.TITLE_ROW, 1).setFontSize(14).setFontWeight('bold');

		// Month row
		sheet.getRange(CONFIG.MONTH_ROW, 1).setValue(monthName);
		sheet.getRange(CONFIG.MONTH_ROW, 1).setFontSize(11).setFontWeight('bold');

		// Headers
		sheet
			.getRange(CONFIG.HEADER_ROW, 1, 1, headers.length)
			.setValues([headers]);
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
	}

	// Write data rows (columns A-E only)
	if (floaterData.length > 0) {
		const dataRows = floaterData.map((emp) => [
			emp.employeeId,
			emp.name,
			emp.department,
			emp.floaterPct / 100, // Store as decimal for percentage formatting
			emp.currentProject,
		]);

		sheet
			.getRange(CONFIG.FIRST_DATA_ROW, 1, dataRows.length, CONFIG.DATA_COLS)
			.setValues(dataRows);

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

		// Add borders to data area (columns A-E)
		sheet
			.getRange(CONFIG.FIRST_DATA_ROW, 1, dataRows.length, CONFIG.DATA_COLS)
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

	if (!isUpdate) {
		// Set column widths (only on first create)
		sheet.setColumnWidth(CONFIG.EMP_ID_COL, 120);
		sheet.setColumnWidth(CONFIG.NAME_COL, 200);
		sheet.setColumnWidth(CONFIG.DEPARTMENT_COL, 150);
		sheet.setColumnWidth(CONFIG.FLOATER_PCT_COL, 100);
		sheet.setColumnWidth(CONFIG.CURRENT_PROJECT_COL, 200);

		// Freeze header rows
		sheet.setFrozenRows(CONFIG.HEADER_ROW);
	}
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
