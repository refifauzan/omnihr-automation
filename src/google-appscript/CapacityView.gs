/**
 * Capacity View - Shows free capacity for each employee per day
 *
 * Calculates: 8 hours - (sum of hours across all project sheets for the same month)
 *
 * Example:
 * - Aaron has 8 hours on Atlas Monday = 0 capacity
 * - Aaron has 1 hour on Atlas + 4 hours on GBG = 3 capacity (8 - 5 = 3)
 * - Holiday or weekend = 0 capacity (greyed out)
 */

/**
 * Generate or update Capacity View for a specific month
 * Prompts user for month/year and which sheets to include
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

	// Get all sheets that could be project sheets for this month
	const allSheets = ss.getSheets();
	const projectSheets = [];

	// Find sheets that have the month/year structure (day columns starting at K)
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

		// Check if sheet has the expected structure (day columns)
		if (sheet.getLastColumn() >= CONFIG.FIRST_DAY_COL) {
			projectSheets.push(sheetName);
		}
	}

	if (projectSheets.length === 0) {
		ui.alert(
			'No project sheets found.\n\nMake sure you have sheets with the day column structure.',
		);
		return;
	}

	// Let user select which sheets to include
	const sheetListHtml = projectSheets
		.map(
			(name, i) =>
				`<label><input type="checkbox" name="sheet" value="${name}" checked> ${name}</label><br>`,
		)
		.join('');

	const html = HtmlService.createHtmlOutput(
		`
		<style>
			body { font-family: Arial, sans-serif; padding: 15px; }
			h3 { margin-top: 0; }
			.sheets { max-height: 200px; overflow-y: auto; margin: 10px 0; }
			button { padding: 10px 20px; background: #4285f4; color: white; border: none; cursor: pointer; margin-right: 10px; }
			button:hover { background: #357abd; }
			.cancel { background: #666; }
		</style>
		<h3>Select Project Sheets</h3>
		<p>Choose which sheets to include in the capacity calculation for ${month + 1}/${year}:</p>
		<div class="sheets">
			${sheetListHtml}
		</div>
		<button onclick="submit()">Generate</button>
		<button class="cancel" onclick="google.script.host.close()">Cancel</button>
		<script>
			function submit() {
				const checkboxes = document.querySelectorAll('input[name="sheet"]:checked');
				const selected = Array.from(checkboxes).map(cb => cb.value);
				if (selected.length === 0) {
					alert('Please select at least one sheet');
					return;
				}
				google.script.run
					.withSuccessHandler(() => google.script.host.close())
					.withFailureHandler(err => alert('Error: ' + err))
					.createCapacityViewSheet(${month}, ${year}, selected);
			}
		</script>
	`,
	)
		.setWidth(400)
		.setHeight(350);

	ui.showModalDialog(html, 'Capacity View - Select Sheets');
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

		// Fetch holidays
		let holidayDays = new Set();
		try {
			const token = getAccessToken();
			if (token) {
				const holidays = fetchHolidaysForMonth(token, month, year);
				holidayDays = new Set(holidays.map((h) => h.date));
				Logger.log(`Found ${holidays.length} holidays`);
			}
		} catch (e) {
			Logger.log('Could not fetch holidays: ' + e.message);
		}

		// Collect all employees from all project sheets
		const employeeMap = new Map(); // employeeId -> { name, team, project }
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

				// Add to global employee map
				if (!employeeMap.has(employeeId)) {
					employeeMap.set(employeeId, {
						name: employeeName,
						team: team,
						project: project,
					});
				}

				// Store hours per day for this employee in this sheet
				const hoursPerDay = {};
				for (let day = 1; day <= daysInMonth; day++) {
					const hours = hoursData[day] ? hoursData[day][i] : 0;
					hoursPerDay[day] = typeof hours === 'number' ? hours : 0;
				}
				sheetEmployeeHours.set(employeeId, hoursPerDay);
			}

			projectData.set(sheetName, { employees: sheetEmployeeHours });
			Logger.log(
				`Loaded ${sheetEmployeeHours.size} employees from ${sheetName}`,
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
		const headerRow1 = ['ID', 'Name', 'Team'];
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

		// Add Total column
		headerRow1.push('');
		headerRow2.push('Total Free');
		headerBgColors.push('#356854');

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
			const row = [employeeId, empInfo.name, empInfo.team];
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
					// Calculate total hours across all project sheets
					let totalHours = 0;
					for (const [sheetName, sheetData] of projectData) {
						const empHours = sheetData.employees.get(employeeId);
						if (empHours && empHours[day]) {
							totalHours += empHours[day];
						}
					}

					// Capacity = 8 - total hours (minimum 0)
					const capacity = Math.max(0, 8 - totalHours);
					row.push(capacity);

					// Color based on capacity
					if (capacity === 0) {
						bgRow.push('#FFE6E6'); // Light red - fully occupied
					} else if (capacity <= 2) {
						bgRow.push('#FFF3CD'); // Light yellow - almost full
					} else if (capacity < 8) {
						bgRow.push('#D4EDDA'); // Light green - partially available
					} else {
						bgRow.push('#FFFFFF'); // White - fully available
					}
				}
			}

			// Add Total Free formula placeholder (will be set as formula)
			row.push(0);
			bgRow.push(null);

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
				// Sum only weekday columns (skip weekends)
				formulas.push([
					`=SUMPRODUCT((D${rowNum}:${columnToLetter(3 + daysInMonth)}${rowNum})<>"",(D${rowNum}:${columnToLetter(3 + daysInMonth)}${rowNum}))`,
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
				`• Project sheets included: ${projectSheetNames.length}\n\n` +
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

	// Get all project sheets and refresh
	const allSheets = ss.getSheets();
	const projectSheetNames = [];

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
		if (sheet.getLastColumn() >= CONFIG.FIRST_DAY_COL) {
			projectSheetNames.push(name);
		}
	}

	if (projectSheetNames.length === 0) {
		ui.alert('No project sheets found.');
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

	// Get all project sheets (same logic as generateCapacityView)
	const allSheets = ss.getSheets();
	const projectSheetNames = [];

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
		if (sheet.getLastColumn() >= CONFIG.FIRST_DAY_COL) {
			projectSheetNames.push(name);
		}
	}

	if (projectSheetNames.length === 0) {
		ui.alert('No project sheets found.');
		return;
	}

	// Update with all found project sheets
	try {
		createCapacityViewSheet(month, year, projectSheetNames);
	} catch (error) {
		ui.alert('Error updating: ' + error.message);
	}
}
