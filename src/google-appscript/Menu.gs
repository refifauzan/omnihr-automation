/**
 * Menu and Trigger functions for OmniHR Leave Integration
 */

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
	try {
		const ui = SpreadsheetApp.getUi();
		ui.createMenu('OmniHR')
			.addItem('Create Empty Table', 'createEmptyTableStructure')
			.addSeparator()
			.addSubMenu(
				ui
					.createMenu('Employees')
					.addItem('Sync Employee List (Full)', 'syncEmployeeList')
					.addItem('Add New Employees', 'addNewEmployees')
					.addSeparator()
					.addItem(
						'Apply Grey-Out (Hire/Termination)',
						'applyEmployeeDateGreyOutMenu',
					),
			)
			.addSeparator()
			.addItem('Sync Leave Data (Custom Month)', 'syncLeaveData')
			.addItem('Sync Leave Only (Current Month)', 'syncLeaveOnly')
			.addItem('Sync Holidays', 'syncHolidays')
			.addSeparator()
			.addSubMenu(
				ui
					.createMenu('Schedule')
					.addItem('View Current Schedule', 'viewTriggers')
					.addItem('Disable Automation', 'removeTriggers'),
			)
			.addSeparator()
			.addSubMenu(
				ui
					.createMenu('Protection')
					.addItem('Enable Edit Protection', 'installOnEditTrigger')
					.addItem('Disable Edit Protection', 'removeOnEditTrigger')
					.addSeparator()
					.addItem('Test Protection Setup', 'testProtectionSetup'),
			)
			.addSeparator()
			.addSubMenu(
				ui
					.createMenu('Capacity View')
					.addItem('Generate Capacity View', 'generateCapacityView')
					.addItem('Update Current View', 'updateCapacityValues'),
			)
			.addSeparator()
			.addItem('Setup API Credentials', 'showCredentialsDialog')
			.addToUi();
	} catch (e) {
		// UI not available (e.g., running from trigger or API context)
		Logger.log('onOpen: UI not available - ' + e.message);
	}
}

/**
 * Setup daily trigger for leave-only sync for a specific month/year
 * Supports multiple sheets - each sheet can have its own month/year configuration
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {string} sheetName - Sheet name to sync to
 */
function setupDailyLeaveOnlyTrigger(month, year, sheetName) {
	const props = PropertiesService.getScriptProperties();

	// Get existing configurations (stored as JSON array)
	let configs = [];
	const existingConfigs = props.getProperty('DAILY_SYNC_CONFIGS');
	if (existingConfigs) {
		try {
			configs = JSON.parse(existingConfigs);
		} catch (e) {
			Logger.log(
				'Error parsing existing configs, starting fresh: ' + e.message,
			);
			configs = [];
		}
	}

	// Check if this sheet already has a config - update it instead of adding duplicate
	const existingIndex = configs.findIndex((c) => c.sheetName === sheetName);
	const newConfig = { month: month, year: year, sheetName: sheetName };

	if (existingIndex >= 0) {
		configs[existingIndex] = newConfig;
		Logger.log(`Updated existing config for sheet "${sheetName}"`);
	} else {
		configs.push(newConfig);
		Logger.log(`Added new config for sheet "${sheetName}"`);
	}

	// Save updated configurations
	props.setProperty('DAILY_SYNC_CONFIGS', JSON.stringify(configs));

	// Ensure we have exactly one daily trigger (don't create duplicates)
	const triggers = ScriptApp.getProjectTriggers();
	const hasTrigger = triggers.some(
		(trigger) => trigger.getHandlerFunction() === 'scheduledLeaveOnlySync',
	);

	if (!hasTrigger) {
		ScriptApp.newTrigger('scheduledLeaveOnlySync')
			.timeBased()
			.everyDays(1)
			.atHour(6)
			.inTimezone('Asia/Kuala_Lumpur')
			.create();
		Logger.log('Created daily sync trigger at 6 AM Malaysia time');
	}

	Logger.log(
		`Daily leave-only sync configured for ${
			month + 1
		}/${year} on sheet "${sheetName}". Total configured sheets: ${
			configs.length
		}`,
	);
}

/**
 * Scheduled leave-only sync function (called by trigger)
 * Syncs leave data to ALL configured sheets
 * Each sheet has its own month/year configuration
 * Automatically removes configs when month/year has passed
 */
function scheduledLeaveOnlySync() {
	const props = PropertiesService.getScriptProperties();
	const configsJson = props.getProperty('DAILY_SYNC_CONFIGS');

	if (!configsJson) {
		Logger.log('Scheduled sync: No configurations found');
		return;
	}

	let configs = [];
	try {
		configs = JSON.parse(configsJson);
	} catch (e) {
		Logger.log('Error parsing configs: ' + e.message);
		return;
	}

	if (configs.length === 0) {
		Logger.log('Scheduled sync: No sheet configurations');
		removeDailySyncTrigger();
		return;
	}

	const now = new Date();
	const currentYear = now.getFullYear();
	const currentMonth = now.getMonth();
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	// Get token once for all sheets
	let token;
	try {
		token = getAccessToken();
		if (!token) {
			Logger.log('Failed to get API token');
			return;
		}
	} catch (error) {
		Logger.log('Error getting access token: ' + error.message);
		return;
	}

	// Fetch employees once for all sheets
	let employees;
	try {
		employees = fetchAllEmployees(token);
		if (!employees || employees.length === 0) {
			Logger.log('No employees found');
			return;
		}
	} catch (error) {
		Logger.log('Error fetching employees: ' + error.message);
		return;
	}

	// Track which configs to keep (not expired)
	const activeConfigs = [];
	let totalSynced = 0;

	Logger.log(`Starting scheduled sync for ${configs.length} sheets`);

	for (const config of configs) {
		const { month, year, sheetName } = config;

		// Check if this config's month/year has passed
		if (year < currentYear || (year === currentYear && month < currentMonth)) {
			Logger.log(
				`Config expired for sheet "${sheetName}": ${
					month + 1
				}/${year} has passed (current: ${currentMonth + 1}/${currentYear})`,
			);
			continue; // Don't add to activeConfigs - effectively removes it
		}

		// Config is still valid
		activeConfigs.push(config);

		const sheet = ss.getSheetByName(sheetName);
		if (!sheet) {
			Logger.log(`Sheet "${sheetName}" not found, skipping`);
			continue;
		}

		Logger.log(
			`Syncing leave data for ${month + 1}/${year} to sheet: ${sheetName}`,
		);

		try {
			const leaveData = fetchLeaveDataForMonth(token, employees, month, year);
			if (!leaveData || Object.keys(leaveData).length === 0) {
				Logger.log(`No leave data found for ${month + 1}/${year}`);
				continue;
			}

			// Fetch holidays to exclude from Total Days Off
			const holidays = fetchHolidaysForMonth(token, month, year);
			const holidayDays = new Set(holidays.map((h) => h.date));

			updateSheetWithLeaveData(
				sheet,
				leaveData,
				month,
				year,
				true,
				holidayDays,
			);

			// Apply grey-out for employees who joined/left mid-month
			Logger.log(`Applying grey-out for ${month + 1}/${year}...`);
			applyEmployeeDateGreyOut(month, year);

			totalSynced++;
			Logger.log(
				`Leave sync complete for "${sheetName}": ${
					Object.keys(leaveData).length
				} employees with leave`,
			);
		} catch (error) {
			Logger.log(`Error syncing sheet "${sheetName}": ${error.message}`);
		}
	}

	// Update stored configs (removes expired ones)
	if (activeConfigs.length !== configs.length) {
		props.setProperty('DAILY_SYNC_CONFIGS', JSON.stringify(activeConfigs));
		Logger.log(
			`Removed ${configs.length - activeConfigs.length} expired config(s)`,
		);
	}

	// If no more active configs, remove the trigger
	if (activeConfigs.length === 0) {
		removeDailySyncTrigger();
		Logger.log('All configs expired, trigger removed');
	}

	SpreadsheetApp.flush();
	Logger.log(
		`Scheduled sync complete: ${totalSynced}/${activeConfigs.length} sheets synced successfully`,
	);
}

/**
 * Remove the daily sync trigger and optionally clear configs
 * @param {string} sheetName - Optional. If provided, only removes config for this sheet
 */
function removeDailySyncTrigger(sheetName) {
	const props = PropertiesService.getScriptProperties();

	if (sheetName) {
		// Remove only the specified sheet's config
		const configsJson = props.getProperty('DAILY_SYNC_CONFIGS');
		if (configsJson) {
			try {
				let configs = JSON.parse(configsJson);
				configs = configs.filter((c) => c.sheetName !== sheetName);
				props.setProperty('DAILY_SYNC_CONFIGS', JSON.stringify(configs));
				Logger.log(
					`Removed config for sheet "${sheetName}". Remaining: ${configs.length}`,
				);

				// If no configs left, remove the trigger
				if (configs.length === 0) {
					removeDailySyncTriggerOnly();
				}
			} catch (e) {
				Logger.log('Error updating configs: ' + e.message);
			}
		}
	} else {
		// Remove all configs and the trigger
		props.deleteProperty('DAILY_SYNC_CONFIGS');
		removeDailySyncTriggerOnly();
		Logger.log('All daily sync configs and trigger removed');
	}
}

/**
 * Remove only the trigger (helper function)
 */
function removeDailySyncTriggerOnly() {
	const triggers = ScriptApp.getProjectTriggers();
	triggers.forEach((trigger) => {
		if (trigger.getHandlerFunction() === 'scheduledLeaveOnlySync') {
			ScriptApp.deleteTrigger(trigger);
			Logger.log('Daily sync trigger removed');
		}
	});
}

/**
 * Remove all triggers (disable automation)
 */
function removeTriggers() {
	const triggers = ScriptApp.getProjectTriggers();
	triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));
	Logger.log('All triggers removed');
	SpreadsheetApp.getUi().alert(
		'Automation disabled.\n\nAll scheduled syncs have been removed.',
	);
}

/**
 * View current triggers and configured sheets
 */
function viewTriggers() {
	const ui = SpreadsheetApp.getUi();
	const props = PropertiesService.getScriptProperties();
	const triggers = ScriptApp.getProjectTriggers();

	// Get daily sync configurations
	let configs = [];
	const configsJson = props.getProperty('DAILY_SYNC_CONFIGS');
	if (configsJson) {
		try {
			configs = JSON.parse(configsJson);
		} catch (e) {
			// Ignore parse errors
		}
	}

	// Check for daily sync trigger
	const hasDailySyncTrigger = triggers.some(
		(t) => t.getHandlerFunction() === 'scheduledLeaveOnlySync',
	);

	if (configs.length === 0 && triggers.length === 0) {
		ui.alert(
			'No scheduled syncs.\n\nUse "Sync Leave Data (Custom Month)" on each sheet to set up daily automation.',
		);
		return;
	}

	let info = '📅 Daily Sync Schedule (6 AM Malaysia Time)\n';
	info += '═══════════════════════════════════\n\n';

	if (configs.length > 0) {
		info += `Configured Sheets: ${configs.length}\n\n`;
		configs.forEach((config, i) => {
			info += `${i + 1}. "${config.sheetName}"\n`;
			info += `   Month: ${config.month + 1}/${config.year}\n\n`;
		});

		if (hasDailySyncTrigger) {
			info += '✅ Daily trigger is ACTIVE\n';
		} else {
			info +=
				'⚠️ Daily trigger is MISSING (run sync on any sheet to recreate)\n';
		}
	}

	// Show other triggers (like protection)
	const otherTriggers = triggers.filter(
		(t) => t.getHandlerFunction() !== 'scheduledLeaveOnlySync',
	);

	if (otherTriggers.length > 0) {
		info += '\n\nOther Triggers:\n';
		otherTriggers.forEach((trigger, i) => {
			info += `• ${trigger.getHandlerFunction()}\n`;
		});
	}

	ui.alert(info);
}

/**
 * Show dialog to set API credentials
 */
function showCredentialsDialog() {
	const html = HtmlService.createHtmlOutput(
		`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 10px; font-weight: bold; }
      input { width: 100%; padding: 8px; margin-top: 5px; box-sizing: border-box; }
      button { margin-top: 20px; padding: 10px 20px; background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #357abd; }
      .hint { font-size: 12px; color: #666; margin-top: 2px; }
    </style>
    <form onsubmit="saveCredentials(event)">
      <label>Base URL:</label>
      <input type="text" id="baseUrl" placeholder="https://api.omnihr.co/api/v1" required>
      <div class="hint">e.g., https://api.omnihr.co/api/v1</div>

      <label>Subdomain:</label>
      <input type="text" id="subdomain" required>

      <label>Username:</label>
      <input type="text" id="username" required>

      <label>Password:</label>
      <input type="password" id="password" required>

      <button type="submit">Save Credentials</button>
    </form>
    <script>
      function saveCredentials(e) {
        e.preventDefault();
        google.script.run
          .withSuccessHandler(() => {
            alert('Credentials saved!');
            google.script.host.close();
          })
          .withFailureHandler((err) => alert('Error: ' + err))
          .saveCredentials(
            document.getElementById('baseUrl').value,
            document.getElementById('subdomain').value,
            document.getElementById('username').value,
            document.getElementById('password').value
          );
      }
    </script>
  `,
	)
		.setWidth(400)
		.setHeight(380);

	SpreadsheetApp.getUi().showModalDialog(html, 'OmniHR API Credentials');
}

/**
 * Save credentials to script properties
 */
function saveCredentials(baseUrl, subdomain, username, password) {
	const props = PropertiesService.getScriptProperties();
	props.setProperty('OMNIHR_BASE_URL', baseUrl);
	props.setProperty('OMNIHR_SUBDOMAIN', subdomain);
	props.setProperty('OMNIHR_USERNAME', username);
	props.setProperty('OMNIHR_PASSWORD', password);
}
