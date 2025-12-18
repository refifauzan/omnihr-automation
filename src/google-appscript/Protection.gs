/**
 * Protection functions for validated week areas
 *
 * When a "Validated" checkbox is TRUE, the corresponding week's day columns
 * are protected. Editing those cells will show a warning popup.
 */

/**
 * Simple trigger that runs on every edit
 * Checks if the edited cell is in a validated (protected) week area
 * @param {Object} e - Edit event object
 */
function onEdit(e) {
	if (!e || !e.range) return;

	const sheet = e.range.getSheet();
	const sheetName = sheet.getName();

	// Skip non-month sheets (e.g., "Attendance", "Config", etc.)
	if (
		sheetName === 'Attendance' ||
		sheetName === 'Config' ||
		sheetName === 'Template'
	) {
		return;
	}

	const row = e.range.getRow();
	const col = e.range.getColumn();

	// Only check data rows (skip header rows)
	if (row < CONFIG.FIRST_DATA_ROW) return;

	// Only check columns in the day area (starting from FIRST_DAY_COL)
	if (col < CONFIG.FIRST_DAY_COL) return;

	// Parse month/year from sheet name (e.g., "January 2025")
	const monthYear = parseMonthYearFromSheetName(sheetName);
	if (!monthYear) return;

	const { month, year } = monthYear;

	// Get week ranges to find which week this column belongs to
	const { weekRanges, validatedColumns } = calculateDayColumns(month, year);

	// Check if the edited column is a Validated or Override column (allow editing those)
	for (const weekRange of weekRanges) {
		if (col === weekRange.validatedCol || col === weekRange.overrideCol) {
			return; // Allow editing checkbox columns
		}
	}

	// Find which week this column belongs to
	const weekRange = findWeekRangeForColumn(col, weekRanges);
	if (!weekRange) return;

	// Check if this week is validated (checkbox is TRUE)
	const validatedValue = sheet.getRange(row, weekRange.validatedCol).getValue();

	if (validatedValue === true) {
		// Week is validated - show warning and revert the change
		const oldValue = e.oldValue !== undefined ? e.oldValue : '';

		// Revert the change
		e.range.setValue(oldValue);

		// Show warning popup
		SpreadsheetApp.getUi().alert(
			'⚠️ Protected Area',
			'This week has been validated. To edit this cell, please uncheck the "Validated" checkbox first.',
			SpreadsheetApp.getUi().ButtonSet.OK
		);
	}
}

/**
 * Parse month and year from sheet name
 * Expected format: "January 2025", "February 2025", etc.
 * @param {string} sheetName - Sheet name
 * @returns {Object|null} { month, year } or null if not a month sheet
 */
function parseMonthYearFromSheetName(sheetName) {
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

	const parts = sheetName.trim().split(' ');
	if (parts.length !== 2) return null;

	const monthIndex = monthNames.indexOf(parts[0]);
	if (monthIndex === -1) return null;

	const year = parseInt(parts[1]);
	if (isNaN(year)) return null;

	return { month: monthIndex, year: year };
}

/**
 * Find which week range a column belongs to
 * @param {number} col - Column number
 * @param {Array} weekRanges - Array of week range objects
 * @returns {Object|null} Week range object or null
 */
function findWeekRangeForColumn(col, weekRanges) {
	for (const weekRange of weekRanges) {
		if (col >= weekRange.startCol && col <= weekRange.endCol) {
			return weekRange;
		}
	}
	return null;
}

/**
 * Installable trigger version for onEdit (if simple trigger doesn't work)
 * This provides more permissions but needs to be installed manually
 * @param {Object} e - Edit event object
 */
function onEditInstallable(e) {
	onEdit(e);
}

/**
 * Install the onEdit trigger programmatically
 * Run this once to set up the installable trigger
 */
function installOnEditTrigger() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	// Remove existing onEdit triggers to avoid duplicates
	const triggers = ScriptApp.getUserTriggers(ss);
	for (const trigger of triggers) {
		if (trigger.getHandlerFunction() === 'onEditInstallable') {
			ScriptApp.deleteTrigger(trigger);
		}
	}

	// Create new trigger
	ScriptApp.newTrigger('onEditInstallable')
		.forSpreadsheet(ss)
		.onEdit()
		.create();

	SpreadsheetApp.getUi().alert(
		'Trigger Installed',
		'The edit protection trigger has been installed successfully.',
		SpreadsheetApp.getUi().ButtonSet.OK
	);
}

/**
 * Remove the onEdit trigger
 */
function removeOnEditTrigger() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const triggers = ScriptApp.getUserTriggers(ss);

	let removed = 0;
	for (const trigger of triggers) {
		if (trigger.getHandlerFunction() === 'onEditInstallable') {
			ScriptApp.deleteTrigger(trigger);
			removed++;
		}
	}

	SpreadsheetApp.getUi().alert(
		'Trigger Removed',
		`Removed ${removed} edit protection trigger(s).`,
		SpreadsheetApp.getUi().ButtonSet.OK
	);
}
