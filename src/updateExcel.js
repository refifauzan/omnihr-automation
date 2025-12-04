require('dotenv').config();
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { month, year } = require('./config');

const COLORS = {
	FULL_DAY: 'FFFF0000', // Red for full day
	HALF_DAY: 'FFFFA500', // Orange for half day
	WEEKEND: 'FFD3D3D3', // Light grey for weekend
	FONT_BLACK: 'FF000000',
	FONT_WHITE: 'FFFFFFFF',
	FONT_GREY: 'FF808080',
};

const HALF_DAY_HOURS = 4;
const HEADER_ROW = 3;

/**
 * Check if a date is a weekend (Saturday or Sunday).
 *
 * @param {number} day - Day of month (1-31).
 * @param {number} month - Month (0-11).
 * @param {number} year - Full year.
 * @return {boolean} True if weekend.
 */
function isWeekend(day, month, year) {
	const date = new Date(year, month, day);
	const dayOfWeek = date.getDay();
	return dayOfWeek === 0 || dayOfWeek === 6;
}

/**
 * Load leave data from a JSON file, filtering out weekends.
 *
 * @param {string} leaveDataPath - Path to the JSON file containing leave data.
 * @param {number} month - Month (0-11) for weekend check.
 * @param {number} year - Year for weekend check.
 * @return {Object} An object mapping employee names to their leave data.
 */
function loadLeaveData(leaveDataPath, month = 10, year = 2025) {
	const leaveData = JSON.parse(fs.readFileSync(leaveDataPath, 'utf8'));
	const employeeLeaves = {};

	for (const emp of leaveData) {
		if (emp.leave_requests && emp.leave_requests.length > 0) {
			// Filter out weekend leaves
			const filteredRequests = emp.leave_requests.filter(
				(leave) => !isWeekend(leave.date, month, year)
			);

			if (filteredRequests.length > 0) {
				const name = emp.employee_name.trim().toLowerCase();
				employeeLeaves[name] = {
					leave_requests: filteredRequests,
					employee_id: emp.employee_id,
				};
			}
		}
	}

	return employeeLeaves;
}

/**
 * Parse all day columns from the header row.
 *
 * @param {Object} sheet - The Excel sheet object.
 * @return {Object} An object mapping day numbers to column indices.
 */
function parseAllDayColumns(sheet) {
	const headerRow = sheet.getRow(HEADER_ROW);
	const dayColumns = {};

	headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
		const value = cell.value;
		let dayNum = null;

		if (typeof value === 'number' && value >= 1 && value <= 31) {
			dayNum = value;
		} else if (typeof value === 'string') {
			const num = parseInt(value, 10);
			if (!isNaN(num) && num >= 1 && num <= 31) {
				dayNum = num;
			}
		}

		if (dayNum) {
			dayColumns[dayNum] = colNumber;
		}
	});

	return dayColumns;
}

/**
 * Get the number of days in a month.
 *
 * @param {number} month - Month (0-11).
 * @param {number} year - Year.
 * @return {number} Number of days in the month.
 */
function getDaysInMonth(month, year) {
	return new Date(year, month + 1, 0).getDate();
}

/**
 * Get weekdays for a given month.
 *
 * @param {number} month - Month (0-11).
 * @param {number} year - Year.
 * @return {Array} Array of weekday numbers.
 */
function getWeekdaysInMonth(month, year) {
	const daysInMonth = getDaysInMonth(month, year);
	const weekdays = [];

	for (let day = 1; day <= daysInMonth; day++) {
		if (!isWeekend(day, month, year)) {
			weekdays.push(day);
		}
	}

	return weekdays;
}

/**
 * Get the day name abbreviation for a date.
 *
 * @param {number} day - Day of month (1-31).
 * @param {number} month - Month (0-11).
 * @param {number} year - Year.
 * @return {string} Day name abbreviation (S, M, T, W, T, F, S).
 */
function getDayName(day, month, year) {
	const dayNames = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];
	const date = new Date(year, month, day);
	return dayNames[date.getDay()];
}

/**
 * Restructure the sheet columns for the specific month.
 * Updates headers to show correct day numbers, day names, and styles weekend columns.
 *
 * @param {Object} sheet - The Excel sheet object.
 * @param {Object} allDayColumns - All day columns mapping from template.
 * @param {number} month - Month (0-11).
 * @param {number} year - Year.
 * @return {Object} Weekday columns mapping (for leave processing).
 */
function restructureColumnsForMonth(
	sheet,
	allDayColumns,
	month,
	year,
	employeeRowsById
) {
	const daysInMonth = getDaysInMonth(month, year);
	const DEFAULT_HOURS = 8;

	// Sort template columns by their original day number (from template header)
	const sortedDays = Object.keys(allDayColumns)
		.map(Number)
		.sort((a, b) => a - b);

	const dayNameRow = sheet.getRow(2);
	const headerRow = sheet.getRow(HEADER_ROW);
	const weekdayColumns = {};
	const weekendDays = [];

	// Get all employee row numbers for setting default values
	const allEmployeeRowNums = new Set();
	Object.values(employeeRowsById).forEach((rows) => {
		rows.forEach((r) => allEmployeeRowNums.add(r.rowNum));
	});

	// Helper function to find the next non-empty value in subsequent columns for a row
	const findNextValidValue = (rowNum, startColIndex) => {
		const colIndices = sortedDays.map((d) => allDayColumns[d]);
		const startIdx = colIndices.indexOf(startColIndex);

		// Search forward
		for (let i = startIdx + 1; i < colIndices.length; i++) {
			const cell = sheet.getCell(rowNum, colIndices[i]);
			const value = cell.value;
			if (value !== null && value !== undefined && value !== '' && value > 0) {
				return value;
			}
		}

		// Search backward if nothing found forward
		for (let i = startIdx - 1; i >= 0; i--) {
			const cell = sheet.getCell(rowNum, colIndices[i]);
			const value = cell.value;
			if (value !== null && value !== undefined && value !== '' && value > 0) {
				return value;
			}
		}

		// Fallback to default
		return DEFAULT_HOURS;
	};

	// Get the last column index from template
	const lastTemplateColIndex = Math.max(...Object.values(allDayColumns));
	const templateColumnCount = sortedDays.length;

	// Process all days in the month
	for (let day = 1; day <= daysInMonth; day++) {
		let colIndex;

		if (day <= templateColumnCount) {
			// Use existing template column
			colIndex = allDayColumns[sortedDays[day - 1]];
		} else {
			// Add new column after the last template column
			colIndex = lastTemplateColIndex + (day - templateColumnCount);
		}

		dayNameRow.getCell(colIndex).value = getDayName(day, month, year);

		// Update row 3 with day number
		headerRow.getCell(colIndex).value = day.toString();

		if (isWeekend(day, month, year)) {
			weekendDays.push(day);
			// Style weekend columns with grey background and 0 value
			allEmployeeRowNums.forEach((rowNum) => {
				const cell = sheet.getCell(rowNum, colIndex);
				cell.value = 0;
				cell.style = {
					fill: {
						type: 'pattern',
						pattern: 'solid',
						fgColor: { argb: COLORS.WEEKEND },
					},
					font: { color: { argb: COLORS.FONT_GREY } },
				};
			});
		} else {
			// Weekday - ensure cells have value if empty (copy from adjacent column)
			allEmployeeRowNums.forEach((rowNum) => {
				const cell = sheet.getCell(rowNum, colIndex);
				const value = cell.value;
				// If cell is empty/null/0, find value from next valid column
				if (
					value === null ||
					value === undefined ||
					value === '' ||
					value === 0
				) {
					cell.value = findNextValidValue(rowNum, colIndex);
				}
			});
			weekdayColumns[day] = colIndex;
		}
	}

	// Clear extra columns if month has fewer days than template
	for (let i = daysInMonth; i < templateColumnCount; i++) {
		const colIndex = allDayColumns[sortedDays[i]];
		dayNameRow.getCell(colIndex).value = '';
		headerRow.getCell(colIndex).value = '';

		allEmployeeRowNums.forEach((rowNum) => {
			const cell = sheet.getCell(rowNum, colIndex);
			cell.value = null;
			cell.style = {};
		});
	}

	console.log(`Restructured for ${month + 1}/${year}: ${daysInMonth} days`);
	console.log(`Weekend days (grey): ${weekendDays.join(', ')}`);
	console.log(
		`Weekday columns: ${Object.keys(weekdayColumns)
			.sort((a, b) => a - b)
			.join(', ')}`
	);

	return weekdayColumns;
}

/**
 * Parse employee rows from the sheet.
 *
 * @param {Object} sheet - The Excel sheet object.
 * @param {Object} dayColumns - Object mapping day numbers to column indices.
 * @return {Object} Objects mapping employee IDs and names to their row data.
 */
function parseEmployeeRows(sheet, dayColumns) {
	const employeeRowsById = {};
	const employeeRowsByName = {};

	sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
		if (rowNumber <= HEADER_ROW) return;

		const sheetId = row.getCell(1).value;
		const employeeName = row.getCell(2).value;

		const dayHours = {};
		for (const [day, colIdx] of Object.entries(dayColumns)) {
			const cellValue = row.getCell(colIdx).value;
			let hours = 0;

			if (typeof cellValue === 'number') {
				hours = cellValue;
			} else if (
				cellValue &&
				typeof cellValue === 'object' &&
				cellValue.result !== undefined
			) {
				// Handle formula cells - use the result value
				hours =
					typeof cellValue.result === 'number'
						? cellValue.result
						: parseFloat(cellValue.result) || 0;
			} else if (typeof cellValue === 'string') {
				hours = parseFloat(cellValue) || 0;
			}

			dayHours[day] = hours;
		}

		const rowData = { rowNum: rowNumber, dayHours };

		if (sheetId && typeof sheetId === 'string') {
			const idUpper = sheetId.trim().toUpperCase();
			if (!employeeRowsById[idUpper]) employeeRowsById[idUpper] = [];
			employeeRowsById[idUpper].push(rowData);
		}

		if (employeeName && typeof employeeName === 'string') {
			const nameLower = employeeName.trim().toLowerCase();
			if (!employeeRowsByName[nameLower]) employeeRowsByName[nameLower] = [];
			employeeRowsByName[nameLower].push(rowData);
		}
	});

	return { employeeRowsById, employeeRowsByName };
}

/**
 * Find employee rows based on name and ID.
 *
 * @param {string} empName - The employee name to search for.
 * @param {string} employeeId - The employee ID to search for.
 * @param {Object} employeeRowsById - Object mapping employee IDs to row data.
 * @param {Object} employeeRowsByName - Object mapping employee names to row data.
 * @return {Array} Array of matching row objects.
 */
function findEmployeeRows(
	empName,
	employeeId,
	employeeRowsById,
	employeeRowsByName
) {
	// Priority 1: Match by employee_id
	if (employeeId) {
		const rows = employeeRowsById[employeeId.toUpperCase()];
		if (rows && rows.length > 0) return rows;
	}

	// Priority 2: Match by exact name
	const exactMatch = employeeRowsByName[empName];
	if (exactMatch && exactMatch.length > 0) return exactMatch;

	// Priority 3: Partial name match
	const partialMatch = Object.keys(employeeRowsByName).find(
		(sheetName) => sheetName.includes(empName) || empName.includes(sheetName)
	);
	if (partialMatch) return employeeRowsByName[partialMatch];

	return null;
}

/**
 * Find employee rows based on name and ID.
 *
 * @param {string} empName - The employee name to search for.
 * @param {string} employeeId - The employee ID to search for.
 * @param {Object} employeeRowsById - Object mapping employee IDs to row data.
 * @param {Object} employeeRowsByName - Object mapping employee names to row data.
 * @return {Array} Array of matching row objects.
 */
function findEmployeeRows(
	empName,
	employeeId,
	employeeRowsById,
	employeeRowsByName
) {
	// Priority 1: Match by employee_id
	if (employeeId) {
		const rows = employeeRowsById[employeeId.toUpperCase()];
		if (rows && rows.length > 0) return rows;
	}

	// Priority 2: Match by exact name
	const exactMatch = employeeRowsByName[empName];
	if (exactMatch && exactMatch.length > 0) return exactMatch;

	// Priority 3: Partial name match
	const partialMatch = Object.keys(employeeRowsByName).find(
		(sheetName) => sheetName.includes(empName) || empName.includes(sheetName)
	);
	if (partialMatch) return employeeRowsByName[partialMatch];

	return null;
}

/**
 * Find employee rows based on name and ID.
 *
 * @param {string} empName - The employee name to search for.
 * @param {string} employeeId - The employee ID to search for.
 * @param {Object} employeeRowsById - Object mapping employee IDs to row data.
 * @param {Object} employeeRowsByName - Object mapping employee names to row data.
 * @return {Array} Array of matching row objects.
 */
function findEmployeeRows(
	empName,
	employeeId,
	employeeRowsById,
	employeeRowsByName
) {
	// Priority 1: Match by employee_id
	if (employeeId) {
		const rows = employeeRowsById[employeeId.toUpperCase()];
		if (rows && rows.length > 0) return rows;
	}

	// Priority 2: Match by exact name
	const exactMatch = employeeRowsByName[empName];
	if (exactMatch && exactMatch.length > 0) return exactMatch;

	// Priority 3: Partial name match
	const partialMatch = Object.keys(employeeRowsByName).find(
		(sheetName) => sheetName.includes(empName) || empName.includes(sheetName)
	);
	if (partialMatch) return employeeRowsByName[partialMatch];

	return null;
}

/**
 * Apply full day leave style to a cell.
 *
 * @param {Object} cell - The Excel cell object.
 */
function applyFullDayStyle(cell) {
	cell.value = 0;
	cell.style = {
		fill: {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: COLORS.FULL_DAY },
		},
		font: { color: { argb: COLORS.FONT_WHITE }, bold: true },
	};
}

/**
 * Apply half day leave style to a cell.
 *
 * @param {Object} cell - The Excel cell object.
 * @param {number} newHours - The new hours value to display.
 */
function applyHalfDayStyle(cell, newHours) {
	cell.value = newHours;
	cell.style = {
		fill: {
			type: 'pattern',
			pattern: 'solid',
			fgColor: { argb: COLORS.HALF_DAY },
		},
		font: { color: { argb: COLORS.FONT_BLACK }, bold: true },
	};
}

/**
 * Process full day leave entries.
 *
 * @param {Object} sheet - The Excel sheet object.
 * @param {Array} targetRows - Array of row objects to process.
 * @param {number} colIndex - The column index to process.
 * @return {number} The count of processed rows.
 */
function processFullDayLeave(sheet, targetRows, colIndex) {
	let count = 0;
	for (const { rowNum } of targetRows) {
		const cell = sheet.getCell(rowNum, colIndex);
		applyFullDayStyle(cell);
		count++;
	}
	return count;
}

/**
 * Process half day leave entries.
 *
 * @param {Object} sheet - The Excel sheet object.
 * @param {Array} targetRows - Array of row objects to process.
 * @param {number} colIndex - The column index to process.
 * @param {number} dayNum - The day number (0-6) to process.
 * @return {number} The count of processed rows.
 */
function processHalfDayLeave(sheet, targetRows, colIndex, dayNum) {
	let count = 0;

	const totalHours = targetRows.reduce(
		(sum, { dayHours }) => sum + (dayHours[dayNum] || 0),
		0
	);

	for (const { rowNum, dayHours } of targetRows) {
		const cell = sheet.getCell(rowNum, colIndex);
		const originalHours = dayHours[dayNum] || 0;

		let hoursToDeduct = 0;
		if (totalHours > 0) {
			hoursToDeduct = (originalHours / totalHours) * HALF_DAY_HOURS;
		}
		const newHours = Math.max(0, originalHours - hoursToDeduct);

		applyHalfDayStyle(cell, newHours);
		count++;
	}

	return count;
}

/**
 * Process leave requests and update the Excel sheet
 *
 * @param {Object} sheet - The Excel sheet object
 * @param {Object} employeeLeaves - The leave data for employees
 * @param {Object} dayColumns - Mapping of dates to column indices
 * @param {Object} employeeRowsById - Mapping of employee IDs to row objects
 * @param {Object} employeeRowsByName - Mapping of employee names to row objects
 * @return {Object} Statistics about the processing
 */
function processLeaveRequests(
	sheet,
	employeeLeaves,
	dayColumns,
	employeeRowsById,
	employeeRowsByName
) {
	let updatedCells = 0;
	let matchedEmployees = 0;
	const notFoundEmployees = [];

	for (const [empName, empData] of Object.entries(employeeLeaves)) {
		const { leave_requests: leaves, employee_id: employeeId } = empData;

		const targetRows = findEmployeeRows(
			empName,
			employeeId,
			employeeRowsById,
			employeeRowsByName
		);

		if (!targetRows || targetRows.length === 0) {
			notFoundEmployees.push(`${empName} (${employeeId || 'no ID'})`);
			continue;
		}

		matchedEmployees++;

		for (const leave of leaves) {
			const colIndex = dayColumns[leave.date];
			if (!colIndex) continue;

			if (leave.is_half_day) {
				updatedCells += processHalfDayLeave(
					sheet,
					targetRows,
					colIndex,
					leave.date
				);
			} else {
				updatedCells += processFullDayLeave(sheet, targetRows, colIndex);
			}
		}
	}

	return { updatedCells, matchedEmployees, notFoundEmployees };
}

/**
 * Convert formulas to static values only in day columns.
 * This preserves formulas in Total Hours, Total Days, etc.
 *
 * @param {Object} sheet - The Excel sheet object.
 * @param {Object} dayColumns - Day columns to convert.
 */
function convertDayColumnFormulas(sheet, dayColumns) {
	const colIndices = new Set(Object.values(dayColumns));

	sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
		if (rowNumber <= HEADER_ROW) return;

		colIndices.forEach((colIdx) => {
			try {
				const cell = row.getCell(colIdx);
				const cellValue = cell.value;
				if (cellValue && typeof cellValue === 'object') {
					if (cellValue.formula || cellValue.sharedFormula) {
						// Preserve the style but convert formula to value
						const style = cell.style;
						cell.value =
							cellValue.result !== undefined ? cellValue.result : null;
						if (style) cell.style = style;
					}
				}
			} catch (e) {
				// Ignore errors for individual cells
			}
		});
	});
}

/**
 * Save the workbook to the specified output path.
 *
 * @param {Object} workbook - The Excel workbook object.
 * @param {string} outputPath - The path to save the workbook.
 */
async function saveWorkbook(workbook, outputPath) {
	// Force Excel to recalculate all formulas when the file is opened
	workbook.calcProperties = { fullCalcOnLoad: true };
	await workbook.xlsx.writeFile(outputPath);
	console.log(`Saved updated file to: ${outputPath}`);
}

/**
 * Log the results of the update operation.
 *
 * @param {Object} results - The results object from processLeaveRequests.
 */
function logResults({ updatedCells, matchedEmployees, notFoundEmployees }) {
	console.log(`Matched ${matchedEmployees} employees with leave requests`);
	console.log(`Updated ${updatedCells} cells`);

	if (notFoundEmployees.length > 0) {
		console.log(`Employees not found in sheet (${notFoundEmployees.length}):`);
		notFoundEmployees
			.slice(0, 10)
			.forEach((name) => console.log(`  - ${name}`));
		if (notFoundEmployees.length > 10) {
			console.log(`  ... and ${notFoundEmployees.length - 10} more`);
		}
	}
}

/**
 * Main function to update Excel with leave data.
 *
 * @param {string} excelPath - Path to the Excel file.
 * @param {string} leaveDataPath - Path to the leave data JSON file.
 * @param {number} month - Month (0-11).
 * @param {number} year - Year.
 */
async function updateExcelWithLeaves(excelPath, leaveDataPath, month, year) {
	console.log(`Processing leaves for ${month + 1}/${year}`);
	console.log('Loading Excel file...');
	const workbook = new ExcelJS.Workbook();
	await workbook.xlsx.readFile(excelPath);

	const sheet = workbook.getWorksheet(1);
	if (!sheet) throw new Error('No worksheet found in Excel file');

	console.log('Loading leave data...');
	const employeeLeaves = loadLeaveData(leaveDataPath, month, year);
	console.log(
		`Found ${Object.keys(employeeLeaves).length} employees with leave requests`
	);

	// Parse all day columns first
	const allDayColumns = parseAllDayColumns(sheet);
	console.log(
		`Found ${Object.keys(allDayColumns).length} day columns in template`
	);

	// Convert formulas to values FIRST (before any restructuring)
	convertDayColumnFormulas(sheet, allDayColumns);

	// Parse employee rows AFTER converting formulas (to get actual values)
	const { employeeRowsById, employeeRowsByName } = parseEmployeeRows(
		sheet,
		allDayColumns
	);

	// Restructure columns for the specific month (weekdays only)
	const dayColumns = restructureColumnsForMonth(
		sheet,
		allDayColumns,
		month,
		year,
		employeeRowsById
	);

	// Reparse employee rows AFTER restructuring to get correct dayHours mapping
	const {
		employeeRowsById: updatedRowsById,
		employeeRowsByName: updatedRowsByName,
	} = parseEmployeeRows(sheet, dayColumns);

	console.log(
		`Found ${Object.keys(updatedRowsById).length} unique employee IDs in sheet`
	);
	console.log(
		`Found ${
			Object.keys(updatedRowsByName).length
		} unique employee names in sheet`
	);

	const results = processLeaveRequests(
		sheet,
		employeeLeaves,
		dayColumns,
		updatedRowsById,
		updatedRowsByName
	);

	logResults(results);

	const outputPath = excelPath.replace('.xlsx', `_${month + 1}_${year}.xlsx`);
	await saveWorkbook(workbook, outputPath);

	return results;
}

/**
 * Entry point of the application.
 *
 * @returns {Promise<void>}
 */
async function main() {
	const excelPath = path.join(__dirname, 'data', 'template.xlsx');
	const leaveDataPath = path.join(__dirname, 'data', 'leave_data.json');

	try {
		await updateExcelWithLeaves(excelPath, leaveDataPath, month, year);
	} catch (error) {
		console.error('Error:', error.message);
		process.exit(1);
	}
}

main();
