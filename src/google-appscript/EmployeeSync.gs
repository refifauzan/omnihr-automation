/**
 * Employee Sync Functions
 * Functions to sync employee list from OmniHR and grey out days based on hire/termination dates
 */

/**
 * Sync full employee list from OmniHR - overwrites existing data in columns A and B
 * Should be used at the beginning of each month
 */
function syncEmployeeList() {
	const ui = SpreadsheetApp.getUi();
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

	// Confirm with user since this will overwrite existing data
	const response = ui.alert(
		'Sync Employee List',
		'This will overwrite the existing employee list in columns A and B.\n\n' +
			'Are you sure you want to continue?',
		ui.ButtonSet.YES_NO
	);

	if (response !== ui.Button.YES) {
		return;
	}

	try {
		const token = getAccessToken();
		if (!token) {
			ui.alert('Failed to get API token. Please setup credentials first.');
			return;
		}

		Logger.log('Fetching all employees from OmniHR...');
		const employees = fetchAllEmployeesWithDetails(token);

		if (!employees || employees.length === 0) {
			ui.alert('No employees found in OmniHR');
			return;
		}

		// Clear existing employee data (columns A and B from row 3 onwards)
		const lastRow = sheet.getLastRow();
		if (lastRow >= CONFIG.FIRST_DATA_ROW) {
			sheet
				.getRange(
					CONFIG.FIRST_DATA_ROW,
					1,
					lastRow - CONFIG.FIRST_DATA_ROW + 1,
					2
				)
				.clearContent();
		}

		// Prepare employee data for sheet
		const employeeData = employees.map((emp) => [
			emp.employee_id || '',
			emp.full_name || '',
		]);

		// Write to sheet
		if (employeeData.length > 0) {
			sheet
				.getRange(CONFIG.FIRST_DATA_ROW, 1, employeeData.length, 2)
				.setValues(employeeData);
		}

		SpreadsheetApp.flush();

		ui.alert(
			`Employee list synced successfully!\n\n` +
				`• ${employeeData.length} employees loaded from OmniHR\n` +
				`• Columns A (ID) and B (Name) updated`
		);
	} catch (error) {
		Logger.log('Error syncing employee list: ' + error.message);
		ui.alert('Error: ' + error.message);
	}
}

/**
 * Add new employees from OmniHR - only adds employees not already in the sheet
 */
function addNewEmployees() {
	const ui = SpreadsheetApp.getUi();
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

	try {
		const token = getAccessToken();
		if (!token) {
			ui.alert('Failed to get API token. Please setup credentials first.');
			return;
		}

		Logger.log('Fetching all employees from OmniHR...');
		const employees = fetchAllEmployeesWithDetails(token);

		if (!employees || employees.length === 0) {
			ui.alert('No employees found in OmniHR');
			return;
		}

		// Get existing employee IDs from the sheet
		const lastRow = sheet.getLastRow();
		const existingIds = new Set();

		if (lastRow >= CONFIG.FIRST_DATA_ROW) {
			const idRange = sheet.getRange(
				CONFIG.FIRST_DATA_ROW,
				CONFIG.EMPLOYEE_ID_COL,
				lastRow - CONFIG.FIRST_DATA_ROW + 1,
				1
			);
			const idValues = idRange.getValues();
			for (const row of idValues) {
				const id = String(row[0] || '')
					.trim()
					.toUpperCase();
				if (id) existingIds.add(id);
			}
		}

		Logger.log(`Found ${existingIds.size} existing employees in sheet`);

		// Find new employees (not in sheet)
		const newEmployees = employees.filter((emp) => {
			const empId = String(emp.employee_id || '')
				.trim()
				.toUpperCase();
			return empId && !existingIds.has(empId);
		});

		if (newEmployees.length === 0) {
			ui.alert(
				'No new employees to add.\n\nAll employees from OmniHR are already in the sheet.'
			);
			return;
		}

		// Prepare new employee data
		const newEmployeeData = newEmployees.map((emp) => [
			emp.employee_id || '',
			emp.full_name || '',
		]);

		// Find the next empty row
		const nextRow =
			lastRow >= CONFIG.FIRST_DATA_ROW ? lastRow + 1 : CONFIG.FIRST_DATA_ROW;

		// Write new employees to sheet
		sheet
			.getRange(nextRow, 1, newEmployeeData.length, 2)
			.setValues(newEmployeeData);

		SpreadsheetApp.flush();

		// Log new employees
		Logger.log(`Added ${newEmployees.length} new employees:`);
		newEmployees.forEach((emp) => {
			Logger.log(`  - ${emp.employee_id}: ${emp.full_name}`);
		});

		ui.alert(
			`New employees added successfully!\n\n` +
				`• ${newEmployees.length} new employees added\n` +
				`• Starting from row ${nextRow}\n\n` +
				`New employees:\n` +
				newEmployees.map((e) => `• ${e.full_name}`).join('\n')
		);
	} catch (error) {
		Logger.log('Error adding new employees: ' + error.message);
		ui.alert('Error: ' + error.message);
	}
}

/**
 * Apply grey-out formatting for employees based on hire/termination dates
 * Grey out days before hire date and after termination date
 * Skips weekends and holidays (does not change their color)
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 */
function applyEmployeeDateGreyOut(month, year) {
	const ui = SpreadsheetApp.getUi();
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

	try {
		const token = getAccessToken();
		if (!token) {
			Logger.log('Failed to get API token for grey-out');
			return 0;
		}

		Logger.log(`Applying grey-out for ${month + 1}/${year}...`);
		const employees = fetchAllEmployeesWithDetails(token);

		if (!employees || employees.length === 0) {
			Logger.log('No employees found');
			return 0;
		}

		// Fetch holidays to skip them
		const holidays = fetchHolidaysForMonth(token, month, year);
		const holidayDays = new Set(holidays.map((h) => h.date));
		Logger.log(`Found ${holidays.length} holidays to skip`);

		const { dayColumns } = calculateDayColumns(month, year);
		const employeeLookup = buildEmployeeLookup(sheet);

		const greyColor = '#D3D3D3'; // Light grey
		let greyedCells = 0;

		for (const emp of employees) {
			const rows = findEmployeeRows(
				employeeLookup,
				emp.employee_id,
				emp.full_name
			);
			if (rows.length === 0) continue;

			// Parse hire date (format: DD/MM/YYYY from API)
			let hireDay = null;
			if (emp.hired_date) {
				const hireDate = parseDateDMY(emp.hired_date);
				if (
					hireDate &&
					hireDate.getMonth() === month &&
					hireDate.getFullYear() === year
				) {
					hireDay = hireDate.getDate();
					Logger.log(`${emp.full_name} hired on day ${hireDay}`);
				} else if (
					hireDate &&
					(hireDate.getFullYear() > year ||
						(hireDate.getFullYear() === year && hireDate.getMonth() > month))
				) {
					// Hired after this month - grey out entire month
					hireDay = 32; // Will grey out all days
					Logger.log(
						`${emp.full_name} hired after this month - greying out all days`
					);
				}
			}

			// Parse termination date
			let terminationDay = null;
			if (emp.termination_date) {
				const termDate = parseDateDMY(emp.termination_date);
				if (
					termDate &&
					termDate.getMonth() === month &&
					termDate.getFullYear() === year
				) {
					terminationDay = termDate.getDate();
					Logger.log(`${emp.full_name} terminated on day ${terminationDay}`);
				} else if (
					termDate &&
					(termDate.getFullYear() < year ||
						(termDate.getFullYear() === year && termDate.getMonth() < month))
				) {
					// Terminated before this month - grey out entire month
					terminationDay = 0; // Will grey out all days
					Logger.log(
						`${emp.full_name} terminated before this month - greying out all days`
					);
				}
			}

			// Skip if no date restrictions
			if (hireDay === null && terminationDay === null) continue;

			// Apply grey-out to each row for this employee
			for (const row of rows) {
				for (const [dayStr, col] of Object.entries(dayColumns)) {
					const dayNum = parseInt(dayStr);
					const date = new Date(year, month, dayNum);
					const dayOfWeek = date.getDay();

					// Skip weekends - don't change their color
					if (dayOfWeek === 0 || dayOfWeek === 6) continue;

					// Skip holidays - don't change their color
					if (holidayDays.has(dayNum)) continue;

					let shouldGrey = false;

					// Grey out days before hire date
					if (hireDay !== null && dayNum < hireDay) {
						shouldGrey = true;
					}

					// Grey out days after termination date
					if (terminationDay !== null && dayNum > terminationDay) {
						shouldGrey = true;
					}

					if (shouldGrey) {
						const cell = sheet.getRange(row, col);
						cell.setBackground(greyColor);
						cell.setValue(''); // Clear any value
						greyedCells++;
					}
				}
			}
		}

		Logger.log(
			`Greyed out ${greyedCells} cells based on hire/termination dates`
		);
		SpreadsheetApp.flush();

		return greyedCells;
	} catch (error) {
		Logger.log('Error applying grey-out: ' + error.message);
		return 0;
	}
}

/**
 * Menu function to apply grey-out with month/year prompt
 */
function applyEmployeeDateGreyOutMenu() {
	const ui = SpreadsheetApp.getUi();

	const monthResponse = ui.prompt(
		'Apply Grey-Out',
		'Enter month (1-12):',
		ui.ButtonSet.OK_CANCEL
	);
	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	const yearResponse = ui.prompt(
		'Apply Grey-Out',
		'Enter year (e.g., 2026):',
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
		const greyedCells = applyEmployeeDateGreyOut(month, year);
		ui.alert(
			`Grey-out applied successfully!\n\n` +
				`• ${greyedCells} cells greyed out based on hire/termination dates`
		);
	} catch (error) {
		ui.alert('Error: ' + error.message);
	}
}
