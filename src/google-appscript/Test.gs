/**
 * Test function to verify auto-stop logic
 * Run this manually from the Apps Script editor
 */
function testAutoStopLogic() {
	// Test 1: Past month
	const props = PropertiesService.getScriptProperties();
	const now = new Date();
	const currentYear = now.getFullYear();
	const currentMonth = now.getMonth();

	// Simulate a past month
	const testMonth = currentMonth - 1;
	const testYear = currentMonth === 0 ? currentYear - 1 : currentYear;

	props.setProperty('DAILY_SYNC_MONTH', String(testMonth));
	props.setProperty('DAILY_SYNC_YEAR', String(testYear));
	props.setProperty('DAILY_SYNC_SHEET', 'Test Sheet');

	Logger.log(`Testing with past month: ${testMonth + 1}/${testYear}`);
	Logger.log(`Current month: ${currentMonth + 1}/${currentYear}`);

	// Run the sync function
	scheduledLeaveOnlySync();

	// Check if trigger was removed
	const triggers = ScriptApp.getProjectTriggers();
	const syncTrigger = triggers.find(
		(t) => t.getHandlerFunction() === 'scheduledLeaveOnlySync'
	);

	if (syncTrigger) {
		Logger.log('FAIL: Trigger still exists');
	} else {
		Logger.log('PASS: Trigger was removed');
	}

	// Test 2: Future month
	const futureMonth = currentMonth + 1;
	const futureYear = currentMonth === 11 ? currentYear + 1 : currentYear;

	props.setProperty('DAILY_SYNC_MONTH', String(futureMonth));
	props.setProperty('DAILY_SYNC_YEAR', String(futureYear));

	Logger.log(`\nTesting with future month: ${futureMonth + 1}/${futureYear}`);

	// Create a temporary trigger
	ScriptApp.newTrigger('scheduledLeaveOnlySync')
		.timeBased()
		.everyDays(1)
		.atHour(6)
		.create();

	scheduledLeaveOnlySync();

	// Check if trigger still exists
	const triggers2 = ScriptApp.getProjectTriggers();
	const syncTrigger2 = triggers2.find(
		(t) => t.getHandlerFunction() === 'scheduledLeaveOnlySync'
	);

	if (syncTrigger2) {
		Logger.log('PASS: Trigger still exists for future month');
	} else {
		Logger.log('FAIL: Trigger was removed for future month');
	}

	// Clean up
	removeTriggers();
	props.deleteProperty('DAILY_SYNC_MONTH');
	props.deleteProperty('DAILY_SYNC_YEAR');
	props.deleteProperty('DAILY_SYNC_SHEET');
}
