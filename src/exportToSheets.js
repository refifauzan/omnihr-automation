#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const LeaveService = require('./leaveService');
const config = require('./config');
const {
	exportToCSV,
	exportBalancesToCSV,
	GoogleSheetsExporter,
} = require('./googleSheetsExport');

async function main() {
	const args = process.argv.slice(2);

	// Parse arguments
	const monthArg = args.find((a) => a.startsWith('--month='));
	const yearArg = args.find((a) => a.startsWith('--year='));
	const outputArg = args.find((a) => a.startsWith('--output='));
	const credentialsArg = args.find((a) => a.startsWith('--credentials='));
	const useCache = args.includes('--cache');
	const pushToSheets = args.includes('--push');
	const csvOnly = args.includes('--csv-only');

	// Default to config month/year
	const month = monthArg ? parseInt(monthArg.split('=')[1]) - 1 : config.month;
	const year = yearArg ? parseInt(yearArg.split('=')[1]) : config.year;
	const outputDir = outputArg
		? outputArg.split('=')[1]
		: path.join(__dirname, 'data');

	console.log(`\n📅 Exporting leave data for ${month + 1}/${year}\n`);

	// Ensure output directory exists
	if (!fs.existsSync(outputDir)) {
		fs.mkdirSync(outputDir, { recursive: true });
	}

	let leaveData;
	const cacheFile = path.join(__dirname, 'data', 'leave_data.json');

	if (useCache && fs.existsSync(cacheFile)) {
		console.log('📂 Using cached data from leave_data.json\n');
		leaveData = JSON.parse(fs.readFileSync(cacheFile, 'utf8'));
	} else {
		// Fetch fresh data from OmniHR
		console.log('🔄 Fetching data from OmniHR API...\n');
		const leaveService = new LeaveService();

		leaveData = await leaveService.getAllLeaveData({
			month,
			year,
			onProgress: (completed, total, name) => {
				process.stdout.write(`\r  Progress: ${completed}/${total} employees`);
			},
		});

		console.log('\n');

		// Save to cache
		fs.writeFileSync(cacheFile, JSON.stringify(leaveData, null, 2));
		console.log(`💾 Data cached to ${cacheFile}\n`);
	}

	// Generate month string for filenames
	const monthStr = `${year}-${String(month + 1).padStart(2, '0')}`;

	// Export to CSV
	const leaveCSV = path.join(outputDir, `leave_requests_${monthStr}.csv`);
	const balancesCSV = path.join(outputDir, `leave_balances_${monthStr}.csv`);

	// Export to CSV
	if (!pushToSheets || csvOnly) {
		console.log('📊 Exporting to CSV...\n');
		exportToCSV(leaveData, month, year, leaveCSV);
		exportBalancesToCSV(leaveData, balancesCSV);
		console.log('\nFiles created:');
		console.log(`  - ${leaveCSV}`);
		console.log(`  - ${balancesCSV}`);
	}

	// Push to Google Sheets if --push flag is set
	if (pushToSheets) {
		const credentialsPath = credentialsArg
			? credentialsArg.split('=')[1]
			: path.join(__dirname, '..', 'google-credentials.json');

		const { spreadsheetId, leaveRequestsSheet, leaveBalancesSheet } =
			config.googleSheets;

		console.log('\n☁️  Pushing to Google Sheets...\n');
		console.log(`  Spreadsheet: ${spreadsheetId}`);

		const exporter = new GoogleSheetsExporter(credentialsPath);
		await exporter.initialize();

		await exporter.uploadLeaveData(
			spreadsheetId,
			leaveRequestsSheet,
			leaveData,
			month,
			year
		);
		await exporter.uploadLeaveBalances(
			spreadsheetId,
			leaveBalancesSheet,
			leaveData
		);

		console.log(
			`\n✅ Data pushed to: https://docs.google.com/spreadsheets/d/${spreadsheetId}`
		);
	}

	console.log('\n✅ Export complete!\n');
}

main().catch((err) => {
	console.error('Error:', err.message);
	process.exit(1);
});
