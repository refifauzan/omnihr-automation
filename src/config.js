/**
 * Global configuration for the application.
 * Change month and year here to affect all scripts.
 */
module.exports = {
	// Month is 0-indexed: 0=Jan, 1=Feb, ..., 11=Dec
	month: 11,
	year: 2025,

	// Google Sheets configuration
	googleSheets: {
		spreadsheetId: '1bNoKQwTNT-w4MFHT_3Slw-ArLEWJlFfMjvHaIerUkTg',
		leaveRequestsSheet: 'Leave Requests',
		leaveBalancesSheet: 'Leave Balances',
	},
};
