require('dotenv').config();
const fs = require('fs');
const path = require('path');
const LeaveService = require('./leaveService');
const { month, year } = require('./config');

async function main() {
	try {
		const leaveService = new LeaveService();

		console.log(
			`Fetching leave data with requests for ${month + 1}/${year}...\n`
		);

		const leaveData = await leaveService.getAllLeaveData({
			month,
			year,
			onProgress: (current, total, name) => {
				process.stdout.write(`Progress: ${current}/${total} - ${name}\r`);
			},
		});

		console.log(`Total employees processed: ${leaveData.length}`);

		// Save to JSON file
		const outputPath = path.join(__dirname, 'data', 'leave_data.json');
		fs.writeFileSync(outputPath, JSON.stringify(leaveData, null, 2));
		console.log(`Data saved to: ${outputPath}`);

		// Show sample with leave requests
		const withRequests = leaveData.filter(
			(e) => e.leave_requests && e.leave_requests.length > 0
		);
		console.log(
			`Employees with leave requests in ${month + 1}/${year}: ${
				withRequests.length
			}`
		);

		if (withRequests.length > 0) {
			console.log('Sample employee with leave requests:');
			console.log(JSON.stringify(withRequests[0], null, 2));
		}
	} catch (error) {
		console.error('Error:', error.message);
		process.exit(1);
	}
}

main();
