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
			.addItem('Sync Leave Data', 'syncLeaveData')
			.addItem('Sync Current Month', 'syncCurrentMonth')
			.addItem('Sync Leave Only (Keep Hours)', 'syncLeaveOnly')
			.addSeparator()
			.addSubMenu(
				ui
					.createMenu('Schedule')
					.addItem(
						'Enable Daily Sync (Leave Only)',
						'setupDailyLeaveOnlyTrigger'
					)
					.addSeparator()
					.addItem('View Current Schedule', 'viewTriggers')
					.addItem('Disable Automation', 'removeTriggers')
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
 * Setup daily trigger for leave-only sync - runs every day at 6 AM
 */
function setupDailyLeaveOnlyTrigger() {
	const triggers = ScriptApp.getProjectTriggers();
	triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));

	ScriptApp.newTrigger('scheduledLeaveOnlySync')
		.timeBased()
		.everyDays(1)
		.atHour(6)
		.create();

	Logger.log('Daily leave-only sync trigger created for 6 AM');
	SpreadsheetApp.getUi().alert(
		'Daily leave-only sync enabled!\n\n' +
			'The script will automatically sync leave data (keeping hours) every day at 6 AM.'
	);
}

/**
 * Scheduled leave-only sync function (called by trigger)
 * Syncs leave data while keeping existing hours
 */
function scheduledLeaveOnlySync() {
	const now = new Date();
	const month = now.getMonth();
	const year = now.getFullYear();

	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getActiveSheet();

	Logger.log(`Scheduled leave-only sync to active sheet: ${sheet.getName()}`);

	try {
		const token = getAccessToken();
		if (!token) {
			Logger.log('Failed to get API token');
			return;
		}

		const employees = fetchAllEmployees(token);
		if (!employees || employees.length === 0) {
			Logger.log('No employees found');
			return;
		}

		const leaveData = fetchLeaveDataForMonth(token, employees, month, year);
		if (!leaveData || Object.keys(leaveData).length === 0) {
			Logger.log('No leave data found for this month');
			return;
		}

		updateSheetWithLeaveData(sheet, leaveData, month, year, true);
		SpreadsheetApp.flush();
		Logger.log(
			`Leave-only sync complete: ${
				Object.keys(leaveData).length
			} employees with leave`
		);
	} catch (error) {
		Logger.log('Error in scheduled leave-only sync: ' + error.message);
	}
}

/**
 * Remove all triggers (disable automation)
 */
function removeTriggers() {
	const triggers = ScriptApp.getProjectTriggers();
	triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));
	Logger.log('All triggers removed');
	SpreadsheetApp.getUi().alert(
		'Automation disabled.\n\nAll scheduled syncs have been removed.'
	);
}

/**
 * View current triggers
 */
function viewTriggers() {
	const triggers = ScriptApp.getProjectTriggers();
	if (triggers.length === 0) {
		SpreadsheetApp.getUi().alert(
			'No scheduled syncs.\n\nUse OmniHR > Schedule to set up automation.'
		);
		return;
	}

	let info = 'Current scheduled syncs:\n\n';
	triggers.forEach((trigger, i) => {
		info += `${
			i + 1
		}. ${trigger.getHandlerFunction()} - ${trigger.getEventType()}\n`;
	});
	SpreadsheetApp.getUi().alert(info);
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
  `
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
