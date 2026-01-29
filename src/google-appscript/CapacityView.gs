/**
 * Capacity View - Shows free capacity for each employee per day
 *
 * CAPACITY CALCULATION LOGIC:
 * - Base capacity: 8 hours per working day
 * - Formula: Capacity = 8 - (total assigned hours across all project sheets)
 * - Minimum capacity: 0 (never negative)
 *
 * EXAMPLES:
 * - Aaron assigned 8 hours on Atlas Monday = 0 capacity (8 - 8 = 0)
 * - Aaron assigned 1 hour on Atlas + 4 hours on GBG = 3 capacity (8 - 5 = 3)
 * - Aaron on full-day leave = 0 capacity (leave overrides all assignments)
 * - Aaron on half-day leave with 2 hours assigned = 2 capacity (4 - 2 = 2)
 * - Weekend/holiday = 0 capacity (greyed out)
 */

/**
 * Entry point for creating Capacity Views - automatically finds matching sheets
 * This function automatically generates capacity reports for the specified month/year
 * without requiring users to manually select sheets.
 */
function generateCapacityView() {
	const ui = SpreadsheetApp.getUi();
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	// Prompt for month
	const monthResponse = ui.prompt(
		'Generate Capacity View',
		'Enter month (1-12):',
		ui.ButtonSet.OK_CANCEL,
	);
	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	// Prompt for year
	const yearResponse = ui.prompt(
		'Generate Capacity View',
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

	// Get all sheets that could be project sheets for this month/year
	const allSheets = ss.getSheets();
	const projectSheets = [];

	// Find sheets that have the exact matching month/year name
	for (const sheet of allSheets) {
		const sheetName = sheet.getName();

		// Skip known non-project sheets
		if (
			sheetName === 'Attendance' ||
			sheetName === 'Config' ||
			sheetName === 'Template' ||
			sheetName.startsWith('CV ')
		) {
			continue;
		}

		// Check if sheet name exactly matches the target month/year
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
		const expectedSheetName = `${monthNames[month]} ${year}`;

		if (sheetName === expectedSheetName) {
			// Check if sheet has the expected structure (day columns)
			if (sheet.getLastColumn() >= CONFIG.FIRST_DAY_COL) {
				projectSheets.push(sheetName);
			}
		}
	}

	if (projectSheets.length === 0) {
		ui.alert(
			`No project sheets found for ${month + 1}/${year}.\n\nMake sure you have sheets named "[Month] [Year]" with the day column structure.\n\nExamples: "January 2026", "February 2026"`,
		);
		return;
	}

	// Automatically create capacity view with all found sheets
	try {
		createCapacityViewSheet(month, year, projectSheets);
	} catch (error) {
		ui.alert('Error generating capacity view: ' + error.message);
	}
}

/**
 * Create or update the Capacity View sheet
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {Array<string>} projectSheetNames - Names of sheets to include
 */
function createCapacityViewSheet(month, year, projectSheetNames) {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const ui = SpreadsheetApp.getUi();

	Logger.log(`Creating Capacity View for ${month + 1}/${year}`);
	Logger.log(`Including sheets: ${projectSheetNames.join(', ')}`);

	try {
		// Get or create the Capacity View sheet
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
		const capacitySheetName = `CV ${monthNames[month]} ${year}`;
		let capacitySheet = ss.getSheetByName(capacitySheetName);

		if (capacitySheet) {
			// Clear existing content
			capacitySheet.clear();
			Logger.log(`Cleared existing sheet: ${capacitySheetName}`);
		} else {
			capacitySheet = ss.insertSheet(capacitySheetName);
			Logger.log(`Created new sheet: ${capacitySheetName}`);
		}

		// Calculate day columns for this month
		const { dayColumns, validatedColumns, weekOverrideColumns } =
			calculateDayColumns(month, year);
		const daysInMonth = new Date(year, month + 1, 0).getDate();

		Logger.log(`Day columns calculated for ${month + 1}/${year}:`);
		Logger.log(`  Days in month: ${daysInMonth}`);
		Logger.log(
			`  Sample day columns: Day 1 -> Col ${dayColumns[1]}, Day 2 -> Col ${dayColumns[2]}, Day 15 -> Col ${dayColumns[15]}`,
		);

		// Fetch holidays
		let holidayDays = new Set();
		let leaveData = new Map(); // employeeId -> Set of leave days

		try {
			const token = getAccessToken();
			if (token) {
				// Fetch holidays
				const holidays = fetchHolidaysForMonth(token, month, year);
				holidayDays = new Set(holidays.map((h) => h.date));
				Logger.log(`Found ${holidays.length} holidays`);

				// Fetch leave data
				const employees = fetchAllEmployees(token);
				if (employees && employees.length > 0) {
					const monthLeaveData = fetchLeaveDataForMonth(
						token,
						employees,
						month,
						year,
					);

					// Convert leave data to Map of employeeId -> Map<day, leaveInfo>
					for (const [employeeId, empData] of Object.entries(monthLeaveData)) {
						const leaveDays = new Map(); // day -> { is_half_day }
						if (empData.leave_requests) {
							Logger.log(
								`Processing leave for employee ${empData.employee_name} (${employeeId}): ${empData.leave_requests.length} requests`,
							);
							for (const leave of empData.leave_requests) {
								// Add all days in the leave period
								leaveDays.set(leave.date, { is_half_day: leave.is_half_day });
								Logger.log(
									`  - Day ${leave.date}: half_day=${leave.is_half_day}, type=${leave.leave_type}`,
								);
							}
						}
						leaveData.set(employeeId, leaveDays);
					}
					Logger.log(`Found leave data for ${leaveData.size} employees`);
				}
			}
		} catch (e) {
			Logger.log('Could not fetch holidays/leave: ' + e.message);
		}

		// Collect all employees from all project sheets - aggregate hours across projects
		const employeeMap = new Map(); // employeeId -> { name, team, totalHoursPerDay }
		const projectData = new Map(); // sheetName -> { employees: Map<employeeId, hoursPerDay> }

		for (const sheetName of projectSheetNames) {
			const sheet = ss.getSheetByName(sheetName);
			if (!sheet) {
				Logger.log(`Sheet not found: ${sheetName}`);
				continue;
			}

			const lastRow = sheet.getLastRow();
			if (lastRow < CONFIG.FIRST_DATA_ROW) {
				Logger.log(`No data in sheet: ${sheetName}`);
				continue;
			}

			const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;

			// Read employee info (columns A-D)
			const employeeData = sheet
				.getRange(CONFIG.FIRST_DATA_ROW, 1, numRows, 4)
				.getValues();

			// Read hours data (day columns only, excluding validated/override columns)
			const hoursData = {};

			for (const [dayStr, col] of Object.entries(dayColumns)) {
				const day = parseInt(dayStr);
				const colData = sheet
					.getRange(CONFIG.FIRST_DATA_ROW, col, numRows, 1)
					.getValues();
				hoursData[day] = colData.map((row) => row[0]);
			}

			// Store data for this sheet
			const sheetEmployeeHours = new Map();

			for (let i = 0; i < numRows; i++) {
				const employeeId = employeeData[i][0];
				const employeeName = employeeData[i][1];
				const team = employeeData[i][2];
				const project = employeeData[i][3];

				if (!employeeId) continue;

				// Initialize employee if not exists
				if (!employeeMap.has(employeeId)) {
					const totalHoursPerDay = {};
					for (let day = 1; day <= daysInMonth; day++) {
						totalHoursPerDay[day] = 0; // Initialize all days to 0
					}
					employeeMap.set(employeeId, {
						name: employeeName,
						projects: new Set([project]), // Track all projects from column C
						totalHoursPerDay: totalHoursPerDay,
					});
				} else {
					// Add project to existing employee
					employeeMap.get(employeeId).projects.add(project);
				}

				// Read hours for this employee-project assignment
				const hoursPerDay = {};
				for (let day = 1; day <= daysInMonth; day++) {
					const hours = hoursData[day] ? hoursData[day][i] : 0;
					const validHours = typeof hours === 'number' ? hours : 0;
					hoursPerDay[day] = validHours;

					// Add to employee's total hours for this day
					employeeMap.get(employeeId).totalHoursPerDay[day] += validHours;
				}

				Logger.log(
					`Added ${employeeName} (${employeeId}) - ${team} - ${project}: Day 1=${hoursPerDay[1]}h`,
				);

				// Store individual project data for reference
				sheetEmployeeHours.set(`${employeeId}_${project}`, hoursPerDay);
			}

			projectData.set(sheetName, { employees: sheetEmployeeHours });
			Logger.log(
				`Loaded ${sheetEmployeeHours.size} employee assignments from ${sheetName}`,
			);
		}

		if (employeeMap.size === 0) {
			ui.alert('No employees found in the selected sheets.');
			return;
		}

		// Sort employees by name
		const sortedEmployees = Array.from(employeeMap.entries()).sort((a, b) =>
			a[1].name.localeCompare(b[1].name),
		);

		// Build the Capacity View sheet
		const dayNames = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];

		// Headers
		const headerRow1 = ['ID', 'Name', 'Projects'];
		const headerRow2 = ['', '', ''];
		const headerBgColors = ['#356854', '#356854', '#356854'];

		for (let day = 1; day <= daysInMonth; day++) {
			const date = new Date(year, month, day);
			const dayOfWeek = date.getDay();
			headerRow1.push(dayNames[dayOfWeek]);
			headerRow2.push(day);
			headerBgColors.push(
				dayOfWeek >= 1 && dayOfWeek <= 5 ? '#356854' : '#efefef',
			);
		}

		// Add Total Days Off and Total Free columns
		headerRow1.push('', '');
		headerRow2.push('Total Days Off', 'Total Free');
		headerBgColors.push('#356854', '#356854');

		// Write headers
		capacitySheet.getRange(1, 1, 1, headerRow1.length).setValues([headerRow1]);
		capacitySheet.getRange(2, 1, 1, headerRow2.length).setValues([headerRow2]);

		const headerRange = capacitySheet.getRange(2, 1, 1, headerRow2.length);
		headerRange.setBackgrounds([headerBgColors]);
		headerRange.setFontColors([
			headerBgColors.map((bg) => (bg === '#356854' ? '#FFFFFF' : '#000000')),
		]);

		// Calculate and write capacity data
		const dataRows = [];
		const bgRows = [];

		for (const [employeeId, empInfo] of sortedEmployees) {
			// Join all projects with comma separator
			const projectsDisplay = Array.from(empInfo.projects).sort().join(', ');
			const row = [employeeId, empInfo.name, projectsDisplay];
			const bgRow = [null, null, null];

			for (let day = 1; day <= daysInMonth; day++) {
				const date = new Date(year, month, day);
				const dayOfWeek = date.getDay();
				const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
				const isHoliday = holidayDays.has(day);

				if (isWeekend) {
					row.push('');
					bgRow.push('#efefef');
				} else if (isHoliday) {
					row.push('');
					bgRow.push('#FFCCCB');
				} else {
					// Check if employee is on leave
					const empLeaveDays = leaveData.get(employeeId);
					const leaveInfo = empLeaveDays && empLeaveDays.get(day);

					if (leaveInfo) {
						// Employee on leave - capacity is 0 regardless of assignments
						Logger.log(
							`Employee ${empInfo.name} on leave day ${day}: half_day=${leaveInfo.is_half_day}`,
						);
						let capacity = 0;
						let bgColor = '';

						if (leaveInfo.is_half_day) {
							// Half-day leave - capacity = 4 - total assigned hours
							const totalHours = empInfo.totalHoursPerDay[day] || 0;
							capacity = Math.max(0, 4 - totalHours);
							bgColor = CONFIG.COLORS.HALF_DAY; // Orange for half-day leave
							Logger.log(
								`  Half-day leave: total assigned hours=${totalHours}, capacity=${capacity}`,
							);
						} else {
							// Full-day leave - no capacity
							capacity = 0;
							bgColor = CONFIG.COLORS.FULL_DAY; // Red for full-day leave
							Logger.log(`  Full-day leave: capacity=0`);
						}

						row.push(capacity);
						bgRow.push(bgColor);
					} else {
						// Calculate total capacity: 8 - sum of all project hours
						const totalHours = empInfo.totalHoursPerDay[day] || 0;
						const capacity = Math.max(0, 8 - totalHours);

						Logger.log(
							`Employee ${empInfo.name} day ${day}: total assigned hours=${totalHours}, capacity=${capacity}`,
						);

						row.push(capacity);

						// Color based on capacity
						if (capacity === 0) {
							bgRow.push('#D4EDDA'); // Green - fully assigned or on leave (no action needed)
						} else if (capacity >= 1 && capacity <= 8) {
							bgRow.push('#FFE6E6'); // Light red - actions needed (lost hours need allocation)
						} else {
							bgRow.push('#FFFFFF'); // White - default (shouldn't occur with current logic)
						}
					}
				}
			}

			// Add Total Days Off and Total Free formula placeholders (will be set as formulas)
			row.push(0, 0);
			bgRow.push(null, null);

			dataRows.push(row);
			bgRows.push(bgRow);
		}

		// Write data
		if (dataRows.length > 0) {
			const dataRange = capacitySheet.getRange(
				3,
				1,
				dataRows.length,
				dataRows[0].length,
			);
			dataRange.setValues(dataRows);
			dataRange.setBackgrounds(bgRows);

			// Set Total Free formulas
			const totalCol = 4 + daysInMonth; // Column after all days
			const formulas = [];
			for (let i = 0; i < dataRows.length; i++) {
				const rowNum = 3 + i;
				// Sum all capacity values (weekends and holidays are empty, so use SUMPRODUCT to ignore them)
				const firstDayCol = columnToLetter(4);
				const lastDayCol = columnToLetter(3 + daysInMonth);
				formulas.push([
					`=SUMPRODUCT((${firstDayCol}${rowNum}:${lastDayCol}${rowNum}<>""),(${firstDayCol}${rowNum}:${lastDayCol}${rowNum}))`,
				]);
			}
			capacitySheet
				.getRange(3, totalCol, formulas.length, 1)
				.setFormulas(formulas);
		}

		// Set column widths
		capacitySheet.setColumnWidth(1, 80); // ID
		capacitySheet.setColumnWidth(2, 150); // Name
		capacitySheet.setColumnWidth(3, 100); // Team
		for (let i = 4; i <= 3 + daysInMonth; i++) {
			capacitySheet.setColumnWidth(i, 35);
		}
		capacitySheet.setColumnWidth(4 + daysInMonth, 80); // Total

		// Freeze header rows and employee columns
		capacitySheet.setFrozenRows(2);
		capacitySheet.setFrozenColumns(3);

		SpreadsheetApp.flush();

		ui.alert(
			`Capacity View generated successfully!\n\n` +
				`• Month: ${month + 1}/${year}\n` +
				`• Employees: ${sortedEmployees.length}\n` +
				`• Project sheets included: ${projectSheetNames.length}\n` +
				`• Sheets: ${projectSheetNames.join(', ')}\n\n` +
				`Sheet: "${capacitySheetName}"`,
		);
	} catch (error) {
		Logger.log('Error creating Capacity View: ' + error.message);
		Logger.log('Stack: ' + error.stack);
		ui.alert('Error: ' + error.message);
	}
}

/**
 * Refresh the current Capacity View sheet
 * Re-reads data from all project sheets and updates capacity values
 */
function refreshCapacityView() {
	const ui = SpreadsheetApp.getUi();
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const activeSheet = ss.getActiveSheet();
	const sheetName = activeSheet.getName();

	// Check if this is a Capacity View sheet
	if (!sheetName.startsWith('CV ')) {
		ui.alert(
			'Please select a Capacity View sheet to refresh.\n\nCapacity View sheets are named "CV [Month] [Year]" (e.g., "CV January 2026")',
		);
		return;
	}

	// Parse month/year from sheet name
	const monthYearPart = sheetName.replace('CV ', '');
	const parsed = parseMonthYearFromSheetName(monthYearPart);

	if (!parsed) {
		ui.alert('Could not parse month/year from sheet name.');
		return;
	}

	// Get project sheets that exactly match the capacity view's month/year name
	const allSheets = ss.getSheets();
	const projectSheetNames = [];

	// Build the expected exact sheet name from the capacity view
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
	const expectedSheetName = `${monthNames[parsed.month]} ${parsed.year}`;

	for (const sheet of allSheets) {
		const name = sheet.getName();
		if (
			name === 'Attendance' ||
			name === 'Config' ||
			name === 'Template' ||
			name.startsWith('CV ')
		) {
			continue;
		}

		// Check if sheet name exactly matches the expected name
		if (name === expectedSheetName) {
			if (sheet.getLastColumn() >= CONFIG.FIRST_DAY_COL) {
				projectSheetNames.push(name);
			}
		}
	}

	if (projectSheetNames.length === 0) {
		ui.alert(
			`No project sheets found for ${parsed.month + 1}/${parsed.year}.\n\nMake sure you have sheets named "[Month] [Year]" that match this capacity view.`,
		);
		return;
	}

	// Re-create the Capacity View with current data
	try {
		createCapacityViewSheet(parsed.month, parsed.year, projectSheetNames);
	} catch (error) {
		ui.alert('Error refreshing: ' + error.message);
	}
}

/**
 * Quick update - updates capacity values without regenerating the whole sheet
 * Useful for real-time updates after editing project sheets
 */
function updateCapacityValues() {
	const ui = SpreadsheetApp.getUi();
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const activeSheet = ss.getActiveSheet();
	const sheetName = activeSheet.getName();

	// Check if this is a Capacity View sheet
	if (!sheetName.startsWith('CV ')) {
		ui.alert(
			'Please select a Capacity View sheet to refresh.\n\nCapacity View sheets are named "CV [Month] [Year]" (e.g., "CV January 2026")',
		);
		return;
	}

	const monthYearPart = sheetName.replace('CV ', '');
	const parsed = parseMonthYearFromSheetName(monthYearPart);

	if (!parsed) {
		ui.alert('Could not parse month/year from sheet name.');
		return;
	}

	const { month, year } = parsed;

	// Get project sheets that exactly match the capacity view's month/year name
	const allSheets = ss.getSheets();
	const projectSheetNames = [];

	// Build the expected exact sheet name from the capacity view
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
	const expectedSheetName = `${monthNames[month]} ${year}`;

	for (const sheet of allSheets) {
		const name = sheet.getName();
		if (
			name === 'Attendance' ||
			name === 'Config' ||
			name === 'Template' ||
			name.startsWith('CV ')
		) {
			continue;
		}

		// Check if sheet name exactly matches the expected name
		if (name === expectedSheetName) {
			if (sheet.getLastColumn() >= CONFIG.FIRST_DAY_COL) {
				projectSheetNames.push(name);
			}
		}
	}

	if (projectSheetNames.length === 0) {
		ui.alert(
			`No project sheets found for ${month + 1}/${year}.\n\nMake sure you have sheets named "[Month] [Year]" that match this capacity view.`,
		);
		return;
	}

	// Update with all found project sheets
	try {
		createCapacityViewSheet(month, year, projectSheetNames);
	} catch (error) {
		ui.alert('Error updating: ' + error.message);
	}
}
