/**
 * Configuration constants for OmniHR Leave Integration
 */
const CONFIG = {
	// Sheet configuration
	DAY_NAME_ROW: 1, // Row 1 containing day names (S, M, T, W, T, F, S)
	HEADER_ROW: 2, // Row 2 containing day numbers
	FIRST_DATA_ROW: 3, // Row 3 = First row with employee data
	EMPLOYEE_ID_COL: 1, // Column A = Employee ID
	EMPLOYEE_NAME_COL: 2, // Column B = Employee Name
	PROJECT_COL: 3, // Column C = Project (for matching with attendance)
	FIRST_DAY_COL: 11, // Column K = Day 1 (same as Excel template)

	// Attendance sheet configuration
	ATTENDANCE_SHEET_NAME: 'Attendance',
	ATTENDANCE_ID_COL: 1, // Column A = Employee ID
	ATTENDANCE_NAME_COL: 2, // Column B = Employee Name
	ATTENDANCE_PROJECT_COL: 3, // Column C = Project
	ATTENDANCE_TYPE_COL: 4, // Column D = Type (Fulltime, Parttime, Custom)
	ATTENDANCE_HOURS_COL: 5, // Column E = Daily Hours
	ATTENDANCE_FIRST_DATA_ROW: 2, // First row with data (after header)

	// Colors (in hex without #)
	COLORS: {
		FULL_DAY: '#FF0000', // Red for full day leave
		HALF_DAY: '#FFA500', // Orange for half day leave
		WEEKEND: '#D3D3D3', // Light grey for weekend
	},

	// Hours
	DEFAULT_HOURS: 8,
	HALF_DAY_HOURS: 4,
};
