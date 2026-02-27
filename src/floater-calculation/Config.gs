/**
 * Configuration constants for Floater Calculation
 */
const CONFIG = {
	// Sheet layout
	TITLE_ROW: 1,
	MONTH_ROW: 2,
	HEADER_ROW: 3,
	FIRST_DATA_ROW: 4,

	// Data columns
	NAME_COL: 1, // Column A = Name
	DEPARTMENT_COL: 2, // Column B = Department
	FLOATER_PCT_COL: 3, // Column C = Floater %
	CURRENT_PROJECT_COL: 4, // Column D = Current Project

	// Conditional Scales legend position
	LEGEND_START_COL: 6, // Column F
	LEGEND_LABEL_COL: 6, // Column F = Label
	LEGEND_COLOR_COL: 7, // Column G = Color swatch

	// Conditional Scale colors (row background based on floater cost)
	SCALES: {
		ABOVE_10K: { label: 'Above RM10k', color: '#E06666', fontColor: '#000000' },
		FROM_7K_TO_10K: { label: '7k to 10k', color: '#EA9999', fontColor: '#000000' },
		FROM_4K_TO_7K: { label: '4k to 7k', color: '#F4CCCC', fontColor: '#000000' },
		BELOW_4K: { label: 'Below 4k', color: '#FCE5CD', fontColor: '#000000' },
		LEAVERS: { label: 'Leavers', color: '#CFE2F3', fontColor: '#000000' },
	},

	// Header styling
	HEADER_BG: '#356854',
	HEADER_FONT_COLOR: '#FFFFFF',

	// Average salary for floater cost calculation (RM)
	AVERAGE_SALARY: 8000,

	// Employees to exclude
	EXCLUDED_EMPLOYEES: ['Omni Support', 'People Culture'],
};
