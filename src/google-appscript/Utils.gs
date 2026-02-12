/**
 * Utility functions for OmniHR Leave Integration
 *
 * WHY: Shared utility functions prevent code duplication and ensure consistent
 * date formatting, column handling, and common operations across all modules.
 * This makes the system more maintainable and reduces bugs from inconsistent implementations.
 */

/**
 * WHY: Standardizes date formatting for API communication and display
 * DD/MM/YYYY format is used because it matches the regional settings and
 * prevents ambiguity between US and European date formats when exchanging data.
 *
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
 * WHY: Safely parses dates from API responses and user input
 * Robust date parsing prevents crashes from malformed data and ensures
 * consistent date handling across the entire system.
 *
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
		parseInt(parts[0]),
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
			} day columns for ${month + 1}/${year}`,
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
		} day columns from header`,
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
 * @param {boolean} exactMatchOnly - If true, do not use partial name match (avoids wrong row when names overlap, e.g. 'Selva Kumaran' matching 'Pavithira Selvakumaran')
 * @param {boolean} idOnly - If true, match only by employee ID; do not use name. Use for hire/termination grey-out so formatting is based on ID only.
 * @returns {Array} Array of row numbers
 */
function findEmployeeRows(
	lookup,
	employeeId,
	employeeName,
	exactMatchOnly,
	idOnly,
) {
	// Try by ID first
	if (employeeId) {
		const idUpper = String(employeeId).trim().toUpperCase();
		const rows = lookup.byId[idUpper];
		if (rows && rows.length > 0) return rows;
	}

	if (idOnly) return [];

	// Try by exact name
	if (employeeName) {
		const nameLower = employeeName.trim().toLowerCase();
		const rows = lookup.byName[nameLower];
		if (rows && rows.length > 0) return rows;

		// Partial match only when allowed (can wrongly match e.g. 'Selva Kumaran' to 'Pavithira Selvakumaran' and apply wrong hire/termination grey-out)
		if (!exactMatchOnly) {
			for (const [sheetName, rows] of Object.entries(lookup.byName)) {
				if (sheetName.includes(nameLower) || nameLower.includes(sheetName)) {
					return rows;
				}
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

/**
 * Clear all leave markings (colors and values) for cells where Time off Override is NOT checked
 * This ensures the sheet is recalculated from OmniHR data only
 * @param {Sheet} sheet - The sheet
 * @param {Object} dayColumns - Day to column mapping
 * @param {Object} dayToOverrideCol - Day to override column mapping
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {Set} holidayDays - Set of holiday day numbers
 */
function clearLeaveCellsRespectingOverride(
	sheet,
	dayColumns,
	dayToOverrideCol,
	month,
	year,
	holidayDays,
) {
	const lastRow = sheet.getLastRow();
	const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;

	if (numRows <= 0) return;

	const cols = Object.values(dayColumns);
	if (cols.length === 0) return;

	const minCol = Math.min(...cols);
	const maxCol = Math.max(...cols);
	const numCols = maxCol - minCol + 1;

	const range = sheet.getRange(CONFIG.FIRST_DATA_ROW, minCol, numRows, numCols);
	const backgrounds = range.getBackgrounds();
	const values = range.getValues();
	const fontColors = range.getFontColors();
	const fontWeights = range.getFontWeights();

	// Get team data from column C to check for Operations team
	const teamData = sheet
		.getRange(CONFIG.FIRST_DATA_ROW, CONFIG.PROJECT_COL, numRows, 1)
		.getValues()
		.map((row) => row[0]);

	// Pre-fetch all override column values (batch read)
	const overrideColValues = {};
	const uniqueOverrideCols = [...new Set(Object.values(dayToOverrideCol))];
	for (const overrideCol of uniqueOverrideCols) {
		overrideColValues[overrideCol] = sheet
			.getRange(CONFIG.FIRST_DATA_ROW, overrideCol, numRows, 1)
			.getValues()
			.map((row) => row[0]);
	}

	// Build reverse lookup: col -> dayStr for faster access
	const colToDayStr = {};
	for (const [dayStr, col] of Object.entries(dayColumns)) {
		colToDayStr[col] = dayStr;
	}

	const fullDayColor = CONFIG.COLORS.FULL_DAY.toUpperCase();
	const halfDayColor = CONFIG.COLORS.HALF_DAY.toUpperCase();

	let clearedCount = 0;

	for (let rowIdx = 0; rowIdx < numRows; rowIdx++) {
		const isOperations =
			teamData[rowIdx] &&
			teamData[rowIdx].toString().toLowerCase() === 'operations';

		for (let colIdx = 0; colIdx < numCols; colIdx++) {
			const bg = backgrounds[rowIdx][colIdx].toUpperCase();

			// Only process cells with leave colors
			if (bg !== fullDayColor && bg !== halfDayColor) continue;

			const col = minCol + colIdx;
			const dayStr = colToDayStr[col];
			if (!dayStr) continue;

			const dayNum = parseInt(dayStr);

			// Check if Time off Override is checked for this day's week (using pre-fetched values)
			const overrideCol = dayToOverrideCol[dayStr];
			if (overrideCol && overrideColValues[overrideCol]) {
				if (overrideColValues[overrideCol][rowIdx] === true) {
					continue; // Skip - Time off Override is checked
				}
			}

			// Check if it's a weekend
			const date = new Date(year, month, dayNum);
			const dayOfWeek = date.getDay();
			if (dayOfWeek === 0 || dayOfWeek === 6) continue; // Skip weekends

			// Check if it's a holiday
			if (holidayDays.has(dayNum)) continue; // Skip holidays

			// Clear the leave marking (batch update arrays)
			backgrounds[rowIdx][colIdx] = null;
			fontColors[rowIdx][colIdx] = '#000000';
			fontWeights[rowIdx][colIdx] = 'normal';
			values[rowIdx][colIdx] = isOperations ? 8 : 0;
			clearedCount++;
		}
	}

	// Batch write all changes at once
	if (clearedCount > 0) {
		range.setValues(values);
		range.setBackgrounds(backgrounds);
		range.setFontColors(fontColors);
		range.setFontWeights(fontWeights);
	}

	Logger.log(
		`Cleared ${clearedCount} leave cells (respecting Time off Override)`,
	);
}

/**
 * Fix manually added rows by applying proper formatting, default values, and checkboxes
 * This ensures rows added mid-month get the same structure as original rows
 * @param {Sheet} sheet - The sheet
 * @param {Object} dayColumns - Day to column mapping
 * @param {Array} validatedColumns - Validated checkbox columns
 * @param {Array} weekOverrideColumns - Week override checkbox columns
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {Set} holidayDays - Set of holiday day numbers
 */
function fixManuallyAddedRows(
	sheet,
	dayColumns,
	validatedColumns,
	weekOverrideColumns,
	month,
	year,
	holidayDays,
) {
	const lastRow = sheet.getLastRow();
	const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;

	if (numRows <= 0) return;

	// Get team data from column C
	const teamData = sheet
		.getRange(CONFIG.FIRST_DATA_ROW, CONFIG.PROJECT_COL, numRows, 1)
		.getValues()
		.map((row) => row[0]);

	const cols = Object.values(dayColumns);
	if (cols.length === 0) return;

	const minCol = Math.min(...cols);
	const maxCol = Math.max(...cols);

	// Include validated and override columns in the range
	const allCols = [...cols, ...validatedColumns, ...weekOverrideColumns];
	const totalLastCol = Math.max(...allCols);
	const numCols = totalLastCol - minCol + 1;

	const range = sheet.getRange(CONFIG.FIRST_DATA_ROW, minCol, numRows, numCols);
	const values = range.getValues();
	const backgrounds = range.getBackgrounds();

	// Build reverse lookup: col -> dayStr for faster access
	const colToDayStr = {};
	for (const [dayStr, col] of Object.entries(dayColumns)) {
		colToDayStr[col] = dayStr;
	}

	// Convert arrays to Sets for O(1) lookup
	const validatedColSet = new Set(validatedColumns);
	const weekOverrideColSet = new Set(weekOverrideColumns);

	const weekendColor = '#efefef';
	const holidayColor = '#FFCCCB';

	let fixedCount = 0;

	for (let rowIdx = 0; rowIdx < numRows; rowIdx++) {
		const isOperations =
			teamData[rowIdx] &&
			teamData[rowIdx].toString().toLowerCase() === 'operations';

		for (let col = minCol; col <= totalLastCol; col++) {
			const colIdx = col - minCol;
			const dayStr = colToDayStr[col];

			if (dayStr) {
				const dayNum = parseInt(dayStr);
				const date = new Date(year, month, dayNum);
				const dayOfWeek = date.getDay();
				const isHoliday = holidayDays.has(dayNum);
				const currentValue = values[rowIdx][colIdx];
				const currentBg = backgrounds[rowIdx][colIdx];

				if (dayOfWeek === 0 || dayOfWeek === 6) {
					// Weekend - ensure empty value and grey background
					if (
						currentValue === '' ||
						currentValue === null ||
						currentBg === '' ||
						currentBg === null ||
						currentBg === '#ffffff'
					) {
						values[rowIdx][colIdx] = '';
						backgrounds[rowIdx][colIdx] = weekendColor;
						fixedCount++;
					}
				} else if (isHoliday) {
					// Holiday - ensure empty value and pastel red background
					if (
						currentValue === '' ||
						currentValue === null ||
						currentBg === '' ||
						currentBg === null ||
						currentBg === '#ffffff'
					) {
						values[rowIdx][colIdx] = '';
						backgrounds[rowIdx][colIdx] = holidayColor;
						fixedCount++;
					}
				} else {
					// Weekday - set default hours if empty
					if (currentValue === '' || currentValue === null) {
						values[rowIdx][colIdx] = isOperations ? 8 : 0;
						fixedCount++;
					}
				}
			} else if (validatedColSet.has(col)) {
				// Validated checkbox column - ensure FALSE if empty
				const currentValue = values[rowIdx][colIdx];
				if (currentValue === '' || currentValue === null) {
					values[rowIdx][colIdx] = false;
					fixedCount++;
				}
			} else if (weekOverrideColSet.has(col)) {
				// Week override checkbox column - ensure FALSE if empty
				const currentValue = values[rowIdx][colIdx];
				if (currentValue === '' || currentValue === null) {
					values[rowIdx][colIdx] = false;
					fixedCount++;
				}
			}
		}
	}

	// Write back updated values and backgrounds
	if (fixedCount > 0) {
		range.setValues(values);
		range.setBackgrounds(backgrounds);

		// Ensure checkboxes are set up for validated and override columns
		for (const col of validatedColumns) {
			const checkboxRange = sheet.getRange(
				CONFIG.FIRST_DATA_ROW,
				col,
				numRows,
				1,
			);
			checkboxRange.insertCheckboxes();
		}
		for (const col of weekOverrideColumns) {
			const checkboxRange = sheet.getRange(
				CONFIG.FIRST_DATA_ROW,
				col,
				numRows,
				1,
			);
			checkboxRange.insertCheckboxes();
		}

		Logger.log(`Fixed ${fixedCount} cells in manually added rows`);
	}
}

/**
 * Set default hours for all employees on working days
 * Operations team: 8 hours, Others: 0 hours
 * Only sets value if current cell is empty (doesn't overwrite existing non-zero values)
 * Respects Time off Override checkbox
 * @param {Sheet} sheet - The sheet
 * @param {Object} dayColumns - Day to column mapping
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {Set} holidayDays - Set of holiday day numbers
 */
function setOperationsDefaultHours(
	sheet,
	dayColumns,
	month,
	year,
	holidayDays,
) {
	const lastRow = sheet.getLastRow();
	const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;

	if (numRows <= 0) return;

	// Get team data from column C
	const teamData = sheet
		.getRange(CONFIG.FIRST_DATA_ROW, CONFIG.PROJECT_COL, numRows, 1)
		.getValues()
		.map((row) => row[0]);

	// Build day -> override column mapping
	const { weekRanges } = calculateDayColumns(month, year);
	const dayToOverrideCol = {};
	for (const weekRange of weekRanges) {
		for (const [dayStr, col] of Object.entries(dayColumns)) {
			if (col >= weekRange.startCol && col <= weekRange.endCol) {
				dayToOverrideCol[dayStr] = weekRange.overrideCol;
			}
		}
	}

	// Pre-fetch all override column values (batch read)
	const overrideColValues = {};
	const uniqueOverrideCols = [...new Set(Object.values(dayToOverrideCol))];
	for (const overrideCol of uniqueOverrideCols) {
		overrideColValues[overrideCol] = sheet
			.getRange(CONFIG.FIRST_DATA_ROW, overrideCol, numRows, 1)
			.getValues()
			.map((row) => row[0]);
	}

	// Build reverse lookup: col -> dayStr for faster access
	const colToDayStr = {};
	for (const [dayStr, col] of Object.entries(dayColumns)) {
		colToDayStr[col] = dayStr;
	}

	const cols = Object.values(dayColumns);
	if (cols.length === 0) return;

	const minCol = Math.min(...cols);
	const maxCol = Math.max(...cols);
	const numCols = maxCol - minCol + 1;

	const range = sheet.getRange(CONFIG.FIRST_DATA_ROW, minCol, numRows, numCols);
	const values = range.getValues();
	const backgrounds = range.getBackgrounds();

	const fullDayColor = CONFIG.COLORS.FULL_DAY.toUpperCase();
	const halfDayColor = CONFIG.COLORS.HALF_DAY.toUpperCase();
	const weekendColor = '#EFEFEF';
	const holidayColor = '#FFCCCB';

	let updatedCount = 0;

	for (let rowIdx = 0; rowIdx < numRows; rowIdx++) {
		const team = teamData[rowIdx]
			? teamData[rowIdx].toString().toLowerCase()
			: '';
		const hasDefaultHours = CONFIG.DEFAULT_HOUR_TEAMS.some((t) =>
			team.includes(t),
		);

		for (let colIdx = 0; colIdx < numCols; colIdx++) {
			const col = minCol + colIdx;
			const dayStr = colToDayStr[col];
			if (!dayStr) continue;

			const dayNum = parseInt(dayStr);

			// Check if Time off Override is checked (using pre-fetched values)
			const overrideCol = dayToOverrideCol[dayStr];
			if (overrideCol && overrideColValues[overrideCol]) {
				if (overrideColValues[overrideCol][rowIdx] === true) {
					continue; // Skip - Time off Override is checked
				}
			}

			// Check if it's a weekend
			const date = new Date(year, month, dayNum);
			const dayOfWeek = date.getDay();
			if (dayOfWeek === 0 || dayOfWeek === 6) {
				// Weekend - ensure empty value and grey background
				if (values[rowIdx][colIdx] === '' || values[rowIdx][colIdx] === null) {
					backgrounds[rowIdx][colIdx] = weekendColor;
				}
				continue;
			}

			// Check if it's a holiday
			if (holidayDays.has(dayNum)) {
				// Holiday - ensure empty value and pastel red background
				// Clear any existing value on holidays
				values[rowIdx][colIdx] = '';
				backgrounds[rowIdx][colIdx] = holidayColor;
				continue;
			}

			// Check if cell has leave color (don't overwrite leave cells)
			const bg = backgrounds[rowIdx][colIdx].toUpperCase();
			if (bg === fullDayColor || bg === halfDayColor) continue;

			// Update if current value is empty, null, or 0 (but not if it has leave color)
			const currentValue = values[rowIdx][colIdx];
			if (currentValue === '' || currentValue === null || currentValue === 0) {
				// Set default: CONFIG.DEFAULT_HOURS for default teams, 0 for others
				values[rowIdx][colIdx] = hasDefaultHours ? CONFIG.DEFAULT_HOURS : 0;
				updatedCount++;
			}
		}
	}

	// Write back updated values and backgrounds
	if (updatedCount > 0) {
		range.setValues(values);
		range.setBackgrounds(backgrounds);
		Logger.log(
			`Set default hours for ${updatedCount} cells (Default teams: ${CONFIG.DEFAULT_HOURS}h, Others: 0h)`,
		);
	}
}

/**
 * Determine if a leave day is half-day based on duration codes
 * @param {boolean} isSingleDay - Is single day leave
 * @param {boolean} isFirstDay - Is first day of leave range
 * @param {boolean} isLastDay - Is last day of leave range
 * @param {number} effectiveDuration - Duration code for first day (1=full, 2=half AM, 3=half PM)
 * @param {number} endDuration - Duration code for last day
 * @returns {boolean} True if half-day
 */
function determineHalfDay(
	isSingleDay,
	isFirstDay,
	isLastDay,
	effectiveDuration,
	endDuration,
) {
	const isHalfDuration = (d) => d === 2 || d === 3;

	if (isSingleDay) return isHalfDuration(effectiveDuration);
	if (isFirstDay) return isHalfDuration(effectiveDuration);
	if (isLastDay) return isHalfDuration(endDuration);
	return false;
}

/**
 * Build row hours mapping from sheet cells
 * @param {Sheet} sheet - The sheet
 * @param {Array} rows - Row numbers
 * @param {Object} dayColumns - Day to column mapping
 * @returns {Object} Row to hours mapping
 */
function buildRowHoursFromSheet(sheet, rows, dayColumns) {
	const rowHoursMap = {};

	for (const row of rows) {
		let foundHours = null;

		for (const col of Object.values(dayColumns)) {
			const cellValue = sheet.getRange(row, col).getValue();
			if (typeof cellValue === 'number' && cellValue > 0) {
				foundHours = cellValue;
				break;
			}
		}

		rowHoursMap[row] = foundHours || CONFIG.DEFAULT_HOURS;
		Logger.log(`Row ${row}: ${rowHoursMap[row]} hours from sheet`);
	}

	return rowHoursMap;
}

/**
 * Build row hours mapping from attendance data
 * @param {Sheet} sheet - The sheet
 * @param {Array} rows - Row numbers
 * @param {Array} empAttendance - Employee attendance records
 * @returns {Object} Row to hours mapping
 */
function buildRowHoursFromAttendance(sheet, rows, empAttendance) {
	const rowHoursMap = {};

	for (const row of rows) {
		const projectName = String(sheet.getRange(row, 3).getValue() || '')
			.trim()
			.toUpperCase();

		const matchingAtt = empAttendance.find(
			(att) =>
				String(att.project || '')
					.trim()
					.toUpperCase() === projectName,
		);

		const hours =
			matchingAtt && matchingAtt.hours > 0
				? matchingAtt.hours
				: CONFIG.DEFAULT_HOURS;

		rowHoursMap[row] = hours;
		Logger.log(`Row ${row} (${projectName}): ${hours} hours`);
	}

	return rowHoursMap;
}

/**
 * Get active rows (not overridden) for a leave day
 * @param {Sheet} sheet - The sheet
 * @param {Array} rows - All employee rows
 * @param {number|null} overrideCol - Override column number
 * @param {number} leaveDate - Leave date (day of month)
 * @returns {Array} Active row numbers
 */
function getActiveRows(sheet, rows, overrideCol, leaveDate) {
	if (!overrideCol) {
		Logger.log(`No override column for day ${leaveDate}, all rows active`);
		return rows;
	}

	return rows.filter((row) => {
		const overrideValue = sheet.getRange(row, overrideCol).getValue();
		if (overrideValue === true) {
			Logger.log(
				`Skipping row ${row} for day ${leaveDate} - Week Override checked`,
			);
			return false;
		}
		return true;
	});
}

/**
 * Assign leave cells to full-day or half-day collections
 * @param {Object} leave - Leave object with is_half_day flag
 * @param {Array} activeRows - Active row numbers
 * @param {number} col - Column number
 * @param {Array} fullDayCells - Full-day cells array (mutated)
 * @param {Object} halfDayCellsMap - Half-day cells map (mutated)
 */
function assignLeaveCells(
	leave,
	activeRows,
	col,
	fullDayCells,
	halfDayCellsMap,
) {
	if (!leave.is_half_day) {
		for (const row of activeRows) {
			fullDayCells.push(columnToLetter(col) + row);
		}
		return;
	}

	const hoursPerProject = CONFIG.HALF_DAY_HOURS / activeRows.length;
	Logger.log(
		`Half-day leave: ${activeRows.length} projects, ${hoursPerProject}hr each`,
	);

	for (const row of activeRows) {
		const cellA1 = columnToLetter(col) + row;
		halfDayCellsMap[cellA1] = hoursPerProject;
		Logger.log(`Row ${row}: ${hoursPerProject}hr (ORANGE)`);
	}
}

/**
 * Assign leave cells with object format for half-day cells
 * @param {Object} leave - Leave object
 * @param {Array} activeRows - Active row numbers
 * @param {number} col - Column number
 * @param {Array} fullDayCells - Full-day cells array (mutated)
 * @param {Array} halfDayCells - Half-day cells array with {cell, value} (mutated)
 */
function assignLeaveCellsWithObjects(
	leave,
	activeRows,
	col,
	fullDayCells,
	halfDayCells,
) {
	if (!leave.is_half_day) {
		for (const row of activeRows) {
			fullDayCells.push(columnToLetter(col) + row);
		}
		return;
	}

	const hoursPerProject = CONFIG.HALF_DAY_HOURS / activeRows.length;
	Logger.log(
		`Half-day leave: ${activeRows.length} projects, ${hoursPerProject}hr each`,
	);

	for (const row of activeRows) {
		halfDayCells.push({
			cell: columnToLetter(col) + row,
			value: hoursPerProject,
		});
		Logger.log(`Row ${row}: ${hoursPerProject}hr (ORANGE)`);
	}
}

/**
 * Build conditional formatting rule for a week range
 * @param {Sheet} sheet - The sheet
 * @param {Object} weekRange - Week range object
 * @param {Object} dayColumns - Day to column mapping
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {number} lastRow - Last row number
 * @param {number} numRows - Number of data rows
 * @param {Set} leaveCellSet - Set of leave cell A1 notations
 * @returns {ConditionalFormatRule|null} Rule or null
 */
function buildWeekConditionalRule(
	sheet,
	weekRange,
	dayColumns,
	month,
	year,
	lastRow,
	numRows,
	leaveCellSet,
) {
	const { startCol, endCol, validatedCol } = weekRange;
	const checkboxColLetter = columnToLetter(validatedCol);
	const weekdayRanges = [];

	for (const [dayStr, col] of Object.entries(dayColumns)) {
		if (col < startCol || col > endCol) continue;

		const day = parseInt(dayStr);
		const date = new Date(year, month, day);
		const dayOfWeek = date.getDay();

		if (dayOfWeek < 1 || dayOfWeek > 5) continue;

		const colLetter = columnToLetter(col);
		const ranges = buildNonLeaveRanges(
			sheet,
			col,
			colLetter,
			lastRow,
			leaveCellSet,
		);
		weekdayRanges.push(...ranges);
	}

	weekdayRanges.push(
		sheet.getRange(CONFIG.FIRST_DATA_ROW, validatedCol, numRows, 1),
	);

	if (weekdayRanges.length === 0) return null;

	return SpreadsheetApp.newConditionalFormatRule()
		.whenFormulaSatisfied(`=$${checkboxColLetter}${CONFIG.FIRST_DATA_ROW}=TRUE`)
		.setBackground('#B8E1CD')
		.setRanges(weekdayRanges)
		.build();
}

/**
 * Build ranges excluding leave cells for conditional formatting
 * @param {Sheet} sheet - The sheet
 * @param {number} col - Column number
 * @param {string} colLetter - Column letter
 * @param {number} lastRow - Last row number
 * @param {Set} leaveCellSet - Set of leave cell A1 notations
 * @returns {Array} Array of Range objects
 */
function buildNonLeaveRanges(sheet, col, colLetter, lastRow, leaveCellSet) {
	const ranges = [];
	let rangeStart = null;

	for (let row = CONFIG.FIRST_DATA_ROW; row <= lastRow + 1; row++) {
		const cellA1 = `${colLetter}${row}`;
		const isLeave = leaveCellSet.has(cellA1.toUpperCase());
		const isLastRow = row > lastRow;

		const shouldEndRange = isLeave || isLastRow;
		const shouldStartRange = !isLeave && !isLastRow && rangeStart === null;

		if (shouldStartRange) {
			rangeStart = row;
			continue;
		}

		if (shouldEndRange && rangeStart !== null) {
			const rangeEnd = row - 1;
			ranges.push(
				sheet.getRange(rangeStart, col, rangeEnd - rangeStart + 1, 1),
			);
			rangeStart = null;
		}
	}

	return ranges;
}
