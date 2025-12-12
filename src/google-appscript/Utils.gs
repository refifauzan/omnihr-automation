/**
 * Utility functions for OmniHR Leave Integration
 */

/**
 * Format date as DD/MM/YYYY
 * @param {Date} d - Date object
 * @returns {string} Formatted date
 */
function formatDateDMY(d) {
	const day = String(d.getDate()).padStart(2, '0');
	const month = String(d.getMonth() + 1).padStart(2, '0');
	const year = d.getFullYear();
	return `${day}/${month}/${year}`;
}

/**
 * Parse date in DD/MM/YYYY format
 * @param {string} dateStr - Date string
 * @returns {Date|null}
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
 * Convert column number to letter (1=A, 2=B, 27=AA, etc.)
 * @param {number} column - Column number
 * @returns {string} Column letter
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
 * Calculate day columns, validated columns, and week override columns for a given month/year
 * Layout: [days...] [Validated] [Week Override] [days...] [Validated] [Week Override] ...
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @returns {Object} { dayColumns, validatedColumns, weekOverrideColumns, weekRanges }
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
 * @param {Sheet} sheet - The sheet
 * @param {number} daysInMonth - Days in month
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @returns {Object} Day to column mapping
 */
function getDayColumns(sheet, daysInMonth, month, year) {
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
 * Build lookup of employees from sheet
 * @param {Sheet} sheet - The sheet
 * @returns {Object} Lookup with byId and byName
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
 * Find all employee rows by ID or name
 * @param {Object} lookup - Employee lookup
 * @param {string} employeeId - Employee ID
 * @param {string} employeeName - Employee name
 * @returns {Array} Array of row numbers
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
 * Scan sheet for existing leave cells by background color
 * @param {Sheet} sheet - The sheet
 * @param {Object} dayColumns - Day to column mapping
 * @returns {Array} Array of cell A1 notations
 */
function scanForLeaveCells(sheet, dayColumns) {
	const leaveCells = [];
	const lastRow = sheet.getLastRow();
	const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;

	if (numRows <= 0) return leaveCells;

	const cols = Object.values(dayColumns);
	if (cols.length === 0) return leaveCells;

	const minCol = Math.min(...cols);
	const maxCol = Math.max(...cols);
	const numCols = maxCol - minCol + 1;

	const range = sheet.getRange(CONFIG.FIRST_DATA_ROW, minCol, numRows, numCols);
	const backgrounds = range.getBackgrounds();

	const fullDayColor = CONFIG.COLORS.FULL_DAY.toUpperCase();
	const halfDayColor = CONFIG.COLORS.HALF_DAY.toUpperCase();

	for (let rowIdx = 0; rowIdx < numRows; rowIdx++) {
		for (let colIdx = 0; colIdx < numCols; colIdx++) {
			const bg = backgrounds[rowIdx][colIdx].toUpperCase();

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
