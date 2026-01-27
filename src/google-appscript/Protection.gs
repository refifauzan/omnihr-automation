/**
 * Protection functions for validated week areas
 *
 * WHY: Prevents accidental changes to validated timesheets while maintaining
 * data integrity. This protects against unauthorized modifications that could
 * affect payroll, billing, and project tracking. The system allows managers
 * to lock down weeks after review while still providing clear feedback when
 * someone attempts to edit protected data.
 *
 * When a "Validated" checkbox is TRUE, the corresponding week's day columns
 * are protected. Editing those cells will show a warning popup.
 */

/**
 * WHY: Disabled simple trigger to prevent duplicate alerts and race conditions
 * This function is disabled because having both simple and installable triggers
 * running simultaneously caused multiple popups and inconsistent behavior.
 * The installable trigger provides better user experience with proper UI alerts.
 *
 * @param {Object} e - Edit event object
 */
function onEdit(e) {
	// Simple trigger disabled - onEditInstallable handles protection with UI alerts
	// Having both triggers causes duplicate executions and multiple alerts
	return;

	/*
	try {
		if (!e || !e.range) return;

		const sheet = e.range.getSheet();
		const row = e.range.getRow();
		const col = e.range.getColumn();

		// Check if this edit is from a revert operation (prevent re-entry)
		const revertKey = `REVERTING_${sheet.getSheetId()}_${row}_${col}`;
		const cache = CacheService.getScriptCache();
		if (cache.get(revertKey)) {
			return; // Skip - this is a revert operation
		}

		const sheetName = sheet.getName();

		// Skip specific non-data sheets
		if (
			sheetName === 'Attendance' ||
			sheetName === 'Config' ||
			sheetName === 'Template'
		) {
			return;
		}

		// Only check data rows (skip header rows)
		if (row < CONFIG.FIRST_DATA_ROW) return;

		// Only check columns in the day area (starting from FIRST_DAY_COL)
		if (col < CONFIG.FIRST_DAY_COL) return;

		// Find validated column by scanning header row for "Validated" text
		const weekInfo = findWeekInfoFromHeader(sheet, col);
		if (!weekInfo) return;

		// Handle checkbox edits (Validated)
		if (weekInfo.isCheckboxColumn) {
			const headerValue = sheet.getRange(CONFIG.HEADER_ROW, col).getValue();
			if (
				String(headerValue).trim() === 'Validated' &&
				e.range.getValue() === true
			) {
				recordValidationTime(sheet, row, col);
			}
			return;
		}

		// Check if this week is validated (checkbox is TRUE)
		const validatedValue = sheet
			.getRange(row, weekInfo.validatedCol)
			.getValue();

		if (validatedValue === true) {
			// Check for grace period (race condition fix)
			if (wasRecentlyValidated(sheet, row, weekInfo.validatedCol)) {
				return;
			}

			// Week is validated - revert the change
			let oldValue = '';
			if (e.oldValue !== undefined) {
				oldValue = e.oldValue;
			}

			// Set lock before reverting to prevent re-entry
			cache.put(revertKey, 'true', 10); // 10 second lock

			// Revert the change immediately
			e.range.setValue(oldValue);

			// Add a note to the cell as warning (simple trigger can do this)
			e.range.setNote(
				'⚠️ PROTECTED: This week is validated.\nUncheck "Validated" to edit.',
			);
		}
	} catch (error) {
		Logger.log('onEdit error: ' + error.message);
	}
	*/
}

/**
 * Installable trigger version - has full permissions including UI alerts
 * @param {Object} e - Edit event object
 */
function onEditInstallable(e) {
	try {
		if (!e || !e.range) return;

		const sheet = e.range.getSheet();
		const row = e.range.getRow();
		const col = e.range.getColumn();

		const sheetName = sheet.getName();

		// Skip specific non-data sheets
		if (
			sheetName === 'Attendance' ||
			sheetName === 'Config' ||
			sheetName === 'Template'
		) {
			return;
		}

		// Only check data rows
		if (row < CONFIG.FIRST_DATA_ROW) return;

		// Only check columns in the day area
		if (col < CONFIG.FIRST_DAY_COL) return;

		// Find validated column by scanning header row for "Validated" text
		const weekInfo = findWeekInfoFromHeader(sheet, col);
		if (!weekInfo) return;

		// Handle checkbox edits (Validated)
		if (weekInfo.isCheckboxColumn) {
			const headerValue = sheet.getRange(CONFIG.HEADER_ROW, col).getValue();
			if (
				String(headerValue).trim() === 'Validated' &&
				e.range.getValue() === true
			) {
				recordValidationTime(sheet, row, col);
			}
			return;
		}

		// Check if this week is validated
		const validatedValue = sheet
			.getRange(row, weekInfo.validatedCol)
			.getValue();

		if (validatedValue === true) {
			// Check for grace period (race condition fix)
			if (wasRecentlyValidated(sheet, row, weekInfo.validatedCol)) {
				Logger.log('Edit allowed during grace period');
				return;
			}

			// Use LockService to prevent multiple alerts from concurrent triggers
			const lock = LockService.getUserLock();
			const hasLock = lock.tryLock(100); // Try to acquire lock, wait max 100ms

			if (!hasLock) {
				// Another trigger is already handling this - just exit
				return;
			}

			try {
				// Double-check with cache to prevent alert on revert
				const alertKey = `ALERT_SHOWN_${sheet.getSheetId()}_${row}_${col}`;
				const cache = CacheService.getScriptCache();
				if (cache.get(alertKey)) {
					return; // Alert already shown for this cell
				}

				// Mark that we're showing alert for this cell
				cache.put(alertKey, 'true', 5); // 5 second window

				// Revert the change - use oldValue if available
				let oldValue = '';
				if (e.oldValue !== undefined) {
					oldValue = e.oldValue;
				} else {
					const currentValue = e.range.getValue();
					if (typeof currentValue === 'number') {
						oldValue = 0;
					} else {
						oldValue = '';
					}
				}

				e.range.setValue(oldValue);

				// Clear any note that might have been set
				e.range.clearNote();

				// Show warning popup ONCE
				SpreadsheetApp.getUi().alert(
					'⚠️ Protected Area',
					'This week has been validated.\n\nTo edit this cell, please uncheck the "Validated" checkbox first.',
					SpreadsheetApp.getUi().ButtonSet.OK,
				);
			} finally {
				lock.releaseLock();
			}
		}
	} catch (error) {
		Logger.log('onEditInstallable error: ' + error.message);
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

	const parts = sheetName.trim().split(' ');
	if (parts.length !== 2) return null;

	// Try full month names first
	let monthIndex = monthNames.indexOf(parts[0]);

	// If not found, try abbreviated names
	if (monthIndex === -1) {
		monthIndex = monthAbbrevs.indexOf(parts[0]);
	}

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
 * Find week info by scanning the header row for "Validated" columns
 * Works with any sheet name - doesn't require month/year parsing
 * @param {Sheet} sheet - The sheet
 * @param {number} editedCol - The column that was edited
 * @returns {Object|null} { validatedCol, isCheckboxColumn } or null
 */
function findWeekInfoFromHeader(sheet, editedCol) {
	const lastCol = sheet.getLastColumn();
	if (lastCol < CONFIG.FIRST_DAY_COL) return null;

	// Read header row (row 2)
	const headerRange = sheet.getRange(
		CONFIG.HEADER_ROW,
		CONFIG.FIRST_DAY_COL,
		1,
		lastCol - CONFIG.FIRST_DAY_COL + 1,
	);
	const headerValues = headerRange.getValues()[0];

	// Find all "Validated" and "Override" column positions
	const validatedCols = [];
	const overrideCols = [];

	for (let i = 0; i < headerValues.length; i++) {
		const colNum = CONFIG.FIRST_DAY_COL + i;
		const value = String(headerValues[i]).trim();

		if (value === 'Validated') {
			validatedCols.push(colNum);
		} else if (value === 'Override' || value === 'Time off Override') {
			overrideCols.push(colNum);
		}
	}

	// Check if edited column is a checkbox column
	if (validatedCols.includes(editedCol) || overrideCols.includes(editedCol)) {
		return { validatedCol: null, isCheckboxColumn: true };
	}

	// Find the next "Validated" column after the edited column
	// This is the validated checkbox for the week containing the edited cell
	let validatedCol = null;
	for (const vc of validatedCols) {
		if (vc > editedCol) {
			validatedCol = vc;
			break;
		}
	}

	// If no validated column found after, the edited column might be outside week areas
	if (!validatedCol) return null;

	return { validatedCol: validatedCol, isCheckboxColumn: false };
}

/**
 * Install the onEdit trigger programmatically (silent version - no UI alert)
 * Used by sync functions to auto-enable protection
 */
function installOnEditTriggerSilent() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	// Check if trigger already exists
	const triggers = ScriptApp.getUserTriggers(ss);
	for (const trigger of triggers) {
		if (trigger.getHandlerFunction() === 'onEditInstallable') {
			Logger.log('Edit protection trigger already installed');
			return;
		}
	}

	// Create new trigger
	ScriptApp.newTrigger('onEditInstallable')
		.forSpreadsheet(ss)
		.onEdit()
		.create();

	Logger.log('Edit protection trigger installed silently');
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
		SpreadsheetApp.getUi().ButtonSet.OK,
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
		SpreadsheetApp.getUi().ButtonSet.OK,
	);
}

/**
 * Test function to debug protection logic
 * Run this manually to check if everything is set up correctly
 */
function testProtectionSetup() {
	const ui = SpreadsheetApp.getUi();
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getActiveSheet();
	const sheetName = sheet.getName();

	let message = `Sheet name: "${sheetName}"\n\n`;

	// Scan header row for Validated columns
	const lastCol = sheet.getLastColumn();
	if (lastCol >= CONFIG.FIRST_DAY_COL) {
		const headerRange = sheet.getRange(
			CONFIG.HEADER_ROW,
			CONFIG.FIRST_DAY_COL,
			1,
			lastCol - CONFIG.FIRST_DAY_COL + 1,
		);
		const headerValues = headerRange.getValues()[0];

		const validatedCols = [];
		for (let i = 0; i < headerValues.length; i++) {
			const value = String(headerValues[i]).trim();
			if (value === 'Validated') {
				validatedCols.push(CONFIG.FIRST_DAY_COL + i);
			}
		}

		if (validatedCols.length > 0) {
			message += `✅ Found ${validatedCols.length} "Validated" column(s):\n`;
			for (let i = 0; i < validatedCols.length; i++) {
				const col = validatedCols[i];
				const colLetter = columnToLetter(col);
				const checkboxValue = sheet
					.getRange(CONFIG.FIRST_DATA_ROW, col)
					.getValue();
				message += `  Week ${
					i + 1
				}: Column ${colLetter} (${col}) = ${checkboxValue}\n`;
			}
		} else {
			message += `❌ No "Validated" columns found in header row ${CONFIG.HEADER_ROW}\n`;
			message += `Make sure row 2 contains "Validated" text in the checkbox columns.`;
		}
	} else {
		message += `❌ Sheet has no columns from ${CONFIG.FIRST_DAY_COL} onwards.`;
	}

	// Check triggers
	const triggers = ScriptApp.getUserTriggers(ss);
	const editTriggers = triggers.filter(
		(t) => t.getHandlerFunction() === 'onEditInstallable',
	);
	message += `\n\nInstalled edit triggers: ${editTriggers.length}`;

	if (editTriggers.length === 0) {
		message += `\n⚠️ No trigger installed! Go to OmniHR > Protection > Enable Edit Protection`;
	}

	ui.alert('Protection Setup Test', message, ui.ButtonSet.OK);
}

/**
 * Check if a validation checkbox was recently checked (within 5 seconds)
 * This handles the race condition where a user edits a cell and immediately checks "Validated"
 */
function wasRecentlyValidated(sheet, row, validatedCol) {
	try {
		const key = `VALIDATED_${sheet.getSheetId()}_${row}_${validatedCol}`;
		const cache = CacheService.getScriptCache();
		const timestamp = cache.get(key);

		if (timestamp) {
			const timeDiff = new Date().getTime() - parseInt(timestamp);
			// 5 second grace period
			if (timeDiff < 5000) {
				return true;
			}
		}
	} catch (e) {
		// Ignore cache errors
	}
	return false;
}

/**
 * Record when a validation checkbox is checked
 */
function recordValidationTime(sheet, row, validatedCol) {
	try {
		const key = `VALIDATED_${sheet.getSheetId()}_${row}_${validatedCol}`;
		const cache = CacheService.getScriptCache();
		// Store for 1 minute (sufficient for the 5s grace period)
		cache.put(key, new Date().getTime().toString(), 60);
	} catch (e) {
		// Ignore errors
	}
}
