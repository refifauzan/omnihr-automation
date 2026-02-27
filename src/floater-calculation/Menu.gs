/**
 * Menu functions for Floater Calculation
 */

/**
 * Create the Floater menu when the spreadsheet opens
 */
function onOpen() {
	const ui = SpreadsheetApp.getUi();
	ui.createMenu('Floater')
		.addItem('Generate Floater View', 'generateFloaterViewMenu')
		.addItem('Update Current View', 'updateFloaterView')
		.addSeparator()
		.addItem('Setup API Credentials', 'setupCredentials')
		.addToUi();
}

/**
 * Menu handler - Generate Floater View with month/year prompts
 */
function generateFloaterViewMenu() {
	const ui = SpreadsheetApp.getUi();

	const monthResponse = ui.prompt(
		'Generate Floater View',
		'Enter month (1-12):',
		ui.ButtonSet.OK_CANCEL
	);

	if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

	const yearResponse = ui.prompt(
		'Generate Floater View',
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

	generateFloaterView(month, year);
}

/**
 * Setup API credentials in Script Properties
 */
function setupCredentials() {
	const ui = SpreadsheetApp.getUi();
	const props = PropertiesService.getScriptProperties();

	const baseUrlResponse = ui.prompt(
		'API Base URL',
		'Enter OmniHR API base URL (default: https://api.omnihr.co/api/v1):',
		ui.ButtonSet.OK_CANCEL
	);

	if (baseUrlResponse.getSelectedButton() !== ui.Button.OK) return;

	const subdomainResponse = ui.prompt(
		'Subdomain',
		'Enter OmniHR subdomain (e.g., snappymob):',
		ui.ButtonSet.OK_CANCEL
	);

	if (subdomainResponse.getSelectedButton() !== ui.Button.OK) return;

	const usernameResponse = ui.prompt(
		'Username',
		'Enter OmniHR username (email):',
		ui.ButtonSet.OK_CANCEL
	);

	if (usernameResponse.getSelectedButton() !== ui.Button.OK) return;

	const passwordResponse = ui.prompt(
		'Password',
		'Enter OmniHR password:',
		ui.ButtonSet.OK_CANCEL
	);

	if (passwordResponse.getSelectedButton() !== ui.Button.OK) return;

	const baseUrl = baseUrlResponse.getResponseText().trim() || 'https://api.omnihr.co/api/v1';

	props.setProperty('OMNIHR_BASE_URL', baseUrl);
	props.setProperty('OMNIHR_SUBDOMAIN', subdomainResponse.getResponseText().trim());
	props.setProperty('OMNIHR_USERNAME', usernameResponse.getResponseText().trim());
	props.setProperty('OMNIHR_PASSWORD', passwordResponse.getResponseText().trim());

	ui.alert('API credentials saved successfully!');
}
