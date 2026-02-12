/**
 * Configuration constants for OmniHR Leave Integration
 *
 * Centralized configuration makes the system maintainable and allows
 * easy customization for different spreadsheet layouts or business rules.
 * Changes here automatically apply across all functions.
 */
const CONFIG = {
	// Sheet configuration
	// These positions match the Excel template structure to ensure consistency
	// when migrating from Excel to Google Sheets and to maintain data integrity.
	DAY_NAME_ROW: 1, // Row 1 containing day names (S, M, T, W, T, F, S)
	HEADER_ROW: 2, // Row 2 containing day numbers
	FIRST_DATA_ROW: 3, // Row 3 = First row with employee data
	EMPLOYEE_ID_COL: 1, // Column A = Employee ID
	EMPLOYEE_NAME_COL: 2, // Column B = Employee Name
	PROJECT_COL: 3, // Column C = Project (for matching with attendance)
	FIRST_DAY_COL: 11, // Column K = Day 1 (same as Excel template)

	// Attendance sheet configuration
	// Separate attendance tracking allows HR to manage working hours
	// independently from project assignments while maintaining data linkage.
	ATTENDANCE_SHEET_NAME: 'Attendance',
	ATTENDANCE_ID_COL: 1, // Column A = Employee ID
	ATTENDANCE_NAME_COL: 2, // Column B = Employee Name
	ATTENDANCE_PROJECT_COL: 3, // Column C = Project
	ATTENDANCE_TYPE_COL: 4, // Column D = Type (Fulltime, Parttime, Custom)
	ATTENDANCE_HOURS_COL: 5, // Column E = Daily Hours
	ATTENDANCE_FIRST_DATA_ROW: 2, // First row with data (after header)

	// Colors (in hex without #)
	// Consistent color coding helps users quickly identify different types of leave
	// and maintain visual consistency across all sheets and reports.
	COLORS: {
		FULL_DAY: '#FF0000', // Red for full day leave
		HALF_DAY: '#FFA500', // Orange for half day leave
		WEEKEND: '#D3D3D3', // Light grey for weekend
		HOLIDAY: '#FF0000', // Red for public holidays
	},

	// Hours
	// Standardized hour definitions ensure accurate capacity calculations
	// and consistent leave processing across the entire system.
	DEFAULT_HOURS: 8,
	HALF_DAY_HOURS: 4,

	// Teams that get 8 hours by default on working days
	// These teams are automatically assigned 8 hours when creating empty tables
	// or during leave sync for empty cells. Team names are matched as substrings.
	DEFAULT_HOUR_TEAMS: ['operations', 'astro', 'mediacorp'],
};
