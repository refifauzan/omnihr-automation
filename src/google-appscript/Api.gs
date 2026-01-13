/**
 * OmniHR API functions
 */

/**
 * Get access token from OmniHR using username/password
 * @returns {string} Access token
 */
function getAccessToken() {
	const props = PropertiesService.getScriptProperties();
	const baseUrl = props.getProperty('OMNIHR_BASE_URL');
	const subdomain = props.getProperty('OMNIHR_SUBDOMAIN');
	const username = props.getProperty('OMNIHR_USERNAME');
	const password = props.getProperty('OMNIHR_PASSWORD');

	if (!baseUrl || !subdomain || !username || !password) {
		throw new Error(
			'API credentials not configured. Use OmniHR > Setup API Credentials'
		);
	}

	const response = UrlFetchApp.fetch(`${baseUrl}/auth/token/`, {
		method: 'post',
		contentType: 'application/x-www-form-urlencoded',
		payload: `username=${encodeURIComponent(
			username
		)}&password=${encodeURIComponent(password)}`,
		headers: {
			'x-subdomain': subdomain,
		},
		muteHttpExceptions: true,
	});

	const responseText = response.getContentText();
	const data = JSON.parse(responseText);

	const token = data.access || data.token || data.access_token;
	if (token) {
		return token;
	}

	throw new Error('Failed to get access token: ' + responseText);
}

/**
 * Make authenticated API request
 * @param {string} token - Access token
 * @param {string} endpoint - API endpoint
 * @param {Object} params - Query parameters
 * @returns {Object} Parsed JSON response
 */
function apiRequest(token, endpoint, params = {}) {
	const props = PropertiesService.getScriptProperties();
	const baseUrl = props.getProperty('OMNIHR_BASE_URL');
	const subdomain = props.getProperty('OMNIHR_SUBDOMAIN');

	let url = `${baseUrl}${endpoint}`;

	if (Object.keys(params).length > 0) {
		const queryString = Object.entries(params)
			.map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
			.join('&');
		url += '?' + queryString;
	}

	const response = UrlFetchApp.fetch(url, {
		method: 'get',
		headers: {
			Authorization: `Bearer ${token}`,
			'x-subdomain': subdomain,
			'Content-Type': 'application/json',
		},
		muteHttpExceptions: true,
	});

	return JSON.parse(response.getContentText());
}

/**
 * Fetch all employees with pagination
 * @param {string} token - Access token
 * @returns {Array} All employees
 */
function fetchAllEmployees(token) {
	let allEmployees = [];
	let page = 1;
	let hasMore = true;

	while (hasMore) {
		const response = apiRequest(token, '/employee/list/', {
			page,
			page_size: 100,
		});
		const results = response.results || response;
		allEmployees = allEmployees.concat(results);
		hasMore = response.next !== null && response.next !== undefined;
		page++;
	}

	return allEmployees;
}

/**
 * Build batch request objects for UrlFetchApp.fetchAll()
 * @param {string} token - Access token
 * @param {Array} employees - Employee list
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @returns {Object} { requests, requestMeta }
 */
function buildBatchRequests(token, employees, startDate, endDate) {
	const props = PropertiesService.getScriptProperties();
	const baseUrl = props.getProperty('OMNIHR_BASE_URL');
	const subdomain = props.getProperty('OMNIHR_SUBDOMAIN');

	const headers = {
		Authorization: `Bearer ${token}`,
		'x-subdomain': subdomain,
		'Content-Type': 'application/json',
	};

	const requests = [];
	const requestMeta = [];

	for (const emp of employees) {
		const userId = emp.id || emp.user_id;
		const empName = emp.full_name || emp.name || `User ${userId}`;

		// Base data request
		requests.push({
			url: `${baseUrl}/employee/2.0/users/${userId}/base-data/`,
			method: 'get',
			headers: headers,
			muteHttpExceptions: true,
		});
		requestMeta.push({ userId, empName, type: 'base' });

		// Time-off calendar request
		const calendarUrl = `${baseUrl}/employee/1.1/${userId}/time-off-calendar/?start_date=${formatDateDMY(
			startDate
		)}&end_date=${formatDateDMY(endDate)}`;
		requests.push({
			url: calendarUrl,
			method: 'get',
			headers: headers,
			muteHttpExceptions: true,
		});
		requestMeta.push({ userId, empName, type: 'calendar' });
	}

	return { requests, requestMeta };
}

/**
 * Build batch requests for fetching employee base data only
 * @param {string} token - Access token
 * @param {Array} employees - Employee list
 * @returns {Object} { requests }
 */
function buildBaseDataRequests(token, employees) {
	const props = PropertiesService.getScriptProperties();
	const baseUrl = props.getProperty('OMNIHR_BASE_URL');
	const subdomain = props.getProperty('OMNIHR_SUBDOMAIN');

	const headers = {
		Authorization: `Bearer ${token}`,
		'x-subdomain': subdomain,
		'Content-Type': 'application/json',
	};

	const requests = [];

	for (const emp of employees) {
		const userId = emp.id || emp.user_id;
		requests.push({
			url: `${baseUrl}/employee/2.0/users/${userId}/base-data/`,
			method: 'get',
			headers: headers,
			muteHttpExceptions: true,
		});
	}

	return { requests };
}

/**
 * Fetch public holidays for a specific month/year
 * Uses the time-off-calendar endpoint which includes public holidays
 * @param {string} token - Access token
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @returns {Array} Array of { date: dayNumber, name: holidayName }
 */
function fetchHolidaysForMonth(token, month, year) {
	const startDate = new Date(year, month, 1);
	const endDate = new Date(year, month + 1, 0);

	const startDateStr = formatDateDMY(startDate);
	const endDateStr = formatDateDMY(endDate);

	// First we need to get an employee ID to query the calendar
	// (public holidays are the same for all employees)
	try {
		const employees = apiRequest(token, '/employee/list/', {
			page: 1,
			page_size: 1,
		});
		const firstEmployee = (employees.results || employees)[0];

		if (!firstEmployee) {
			Logger.log('No employees found to fetch holidays');
			return [];
		}

		const userId = firstEmployee.id || firstEmployee.user_id;

		// Fetch time-off calendar which includes holiday data
		// Use the month's date range
		const calendar = apiRequest(
			token,
			`/employee/1.1/${userId}/time-off-calendar/`,
			{ start_date: startDateStr, end_date: endDateStr }
		);

		// Extract holidays from the 'holiday' field (not 'public_holiday')
		const holidayGroups = calendar.holiday || [];
		const holidayDays = [];

		for (const holidayGroup of holidayGroups) {
			const results = holidayGroup.results || [];

			for (const holiday of results) {
				const dateStr = holiday.date_from;
				if (!dateStr) continue;

				// Parse DD/MM/YYYY format
				const holidayDate = parseDateDMY(dateStr);

				if (
					holidayDate &&
					holidayDate.getMonth() === month &&
					holidayDate.getFullYear() === year
				) {
					holidayDays.push({
						date: holidayDate.getDate(),
						name: holiday.name || 'Public Holiday',
					});

					Logger.log(
						`Holiday found: ${holidayDate.getDate()}/${month + 1}/${year} - ${
							holiday.name
						}`
					);
				}
			}
		}

		Logger.log(
			`Found ${holidayDays.length} public holidays for ${month + 1}/${year}`
		);
		return holidayDays;
	} catch (e) {
		Logger.log('Error fetching holidays: ' + e.message);
		return [];
	}
}

/**
 * Employees to exclude from sync
 */
const EXCLUDED_EMPLOYEES = ['Omni Support', 'People Culture'];

/**
 * Fetch all employees with their base data (ID, name, hire date, termination date, team)
 * Uses employee/list endpoint for hired_date, base-data for employee_id,
 * job endpoint for team, and onboarding/workflow-dashboard for termination_date
 * @param {string} token - Access token
 * @returns {Array} Array of employee objects with full details
 */
function fetchAllEmployeesWithDetails(token) {
	const allEmployees = fetchAllEmployees(token);

	// Filter out excluded employees
	const employees = allEmployees.filter((emp) => {
		const fullName = emp.full_name || emp.name || '';
		const isExcluded = EXCLUDED_EMPLOYEES.some(
			(excluded) => fullName.toLowerCase() === excluded.toLowerCase()
		);
		if (isExcluded) {
			Logger.log(`Excluding employee: ${fullName}`);
		}
		return !isExcluded;
	});

	Logger.log(
		`Filtered ${allEmployees.length - employees.length} excluded employees`
	);

	const employeeDetails = [];
	const BATCH_SIZE = 50;

	// First, fetch termination dates from onboarding/workflow-dashboard
	const terminationDates = fetchTerminationDates(token);

	// Fetch team and project contribution data for all employees
	const { teamData, projectContribution } = fetchEmployeeJobData(
		token,
		employees
	);

	for (let i = 0; i < employees.length; i += BATCH_SIZE) {
		const batch = employees.slice(i, i + BATCH_SIZE);
		Logger.log(
			`Fetching employee details batch ${
				Math.floor(i / BATCH_SIZE) + 1
			}/${Math.ceil(employees.length / BATCH_SIZE)}`
		);

		const { requests } = buildBaseDataRequests(token, batch);
		const responses = UrlFetchApp.fetchAll(requests);

		for (let j = 0; j < responses.length; j++) {
			try {
				const response = responses[j];
				const responseCode = response.getResponseCode();
				const emp = batch[j];
				const userId = emp.id || emp.user_id;

				// Get employee_id from base-data endpoint
				let employeeId = '';
				if (responseCode === 200) {
					const data = JSON.parse(response.getContentText());
					const baseData = data.data || data;
					employeeId = baseData.employee_id || '';
				}

				// Get termination_date from workflow-dashboard data
				const terminationDate = terminationDates[userId] || null;

				employeeDetails.push({
					user_id: userId,
					employee_id: employeeId,
					full_name: emp.full_name || emp.name || `User ${userId}`,
					hired_date: emp.hired_date || null,
					termination_date: terminationDate,
					team: teamData[userId] || '',
					project_contribution: projectContribution[userId] || '',
					employment_status: emp.employment_status || null,
					employment_status_display: emp.employment_status_display || null,
				});
			} catch (e) {
				const emp = batch[j];
				const userId = emp.id || emp.user_id;
				Logger.log(`Error parsing employee ${userId}: ${e.message}`);
				employeeDetails.push({
					user_id: userId,
					employee_id: '',
					full_name: emp.full_name || emp.name || `User ${userId}`,
					hired_date: emp.hired_date || null,
					termination_date: terminationDates[userId] || null,
					team: teamData[userId] || '',
					project_contribution: projectContribution[userId] || '',
					employment_status: emp.employment_status || null,
					employment_status_display: emp.employment_status_display || null,
				});
			}
		}
	}

	Logger.log(`Fetched details for ${employeeDetails.length} employees`);
	return employeeDetails;
}

/**
 * Custom attribute ID for Project Contribution (Full Time / Part Time)
 */
const PROJECT_CONTRIBUTION_ATTR_ID = 8337;

/**
 * Fetch team and project contribution data for all employees from job endpoint
 * @param {string} token - Access token
 * @param {Array} employees - Employee list
 * @returns {Object} { teamData: Map of user_id -> team_display, projectContribution: Map of user_id -> contribution }
 */
function fetchEmployeeJobData(token, employees) {
	const props = PropertiesService.getScriptProperties();
	const baseUrl = props.getProperty('OMNIHR_BASE_URL');
	const subdomain = props.getProperty('OMNIHR_SUBDOMAIN');

	const headers = {
		Authorization: `Bearer ${token}`,
		'x-subdomain': subdomain,
		'Content-Type': 'application/json',
	};

	const teamData = {};
	const projectContribution = {};
	const BATCH_SIZE = 50;

	Logger.log(
		'Fetching job data (team & project contribution) for employees...'
	);

	for (let i = 0; i < employees.length; i += BATCH_SIZE) {
		const batch = employees.slice(i, i + BATCH_SIZE);
		Logger.log(
			`Fetching job data batch ${Math.floor(i / BATCH_SIZE) + 1}/${Math.ceil(
				employees.length / BATCH_SIZE
			)}`
		);

		const requests = batch.map((emp) => {
			const userId = emp.id || emp.user_id;
			return {
				url: `${baseUrl}/employee/${userId}/job/`,
				method: 'get',
				headers: headers,
				muteHttpExceptions: true,
			};
		});

		const responses = UrlFetchApp.fetchAll(requests);

		for (let j = 0; j < responses.length; j++) {
			try {
				const response = responses[j];
				const responseCode = response.getResponseCode();
				const emp = batch[j];
				const userId = emp.id || emp.user_id;

				if (responseCode === 200) {
					const jobs = JSON.parse(response.getContentText());
					// Get the first (most recent) job record
					if (jobs && jobs.length > 0) {
						const currentJob = jobs[0];

						// Get team
						if (currentJob.team_display) {
							teamData[userId] = currentJob.team_display;
						}

						// Get project contribution from custom_data_attributes_values
						const customAttrs = currentJob.custom_data_attributes_values || [];
						const contributionAttr = customAttrs.find(
							(attr) => attr.attr === PROJECT_CONTRIBUTION_ATTR_ID
						);
						if (
							contributionAttr &&
							contributionAttr.value &&
							contributionAttr.value.value
						) {
							projectContribution[userId] = contributionAttr.value.value;
						}
					}
				}
			} catch (e) {
				// Silently continue if job data fetch fails
			}
		}
	}

	Logger.log(`Fetched team data for ${Object.keys(teamData).length} employees`);
	Logger.log(
		`Fetched project contribution for ${
			Object.keys(projectContribution).length
		} employees`
	);
	return { teamData, projectContribution };
}

/**
 * Fetch termination dates from onboarding/workflow-dashboard endpoint
 * @param {string} token - Access token
 * @returns {Object} Map of user_id -> termination_date
 */
function fetchTerminationDates(token) {
	const terminationDates = {};
	let page = 1;
	let hasMore = true;

	Logger.log('Fetching termination dates from workflow-dashboard...');

	while (hasMore) {
		try {
			const response = apiRequest(token, '/onboarding/workflow-dashboard/', {
				page: page,
				page_size: 100,
			});

			const results = response.results || [];
			for (const emp of results) {
				if (emp.termination_date) {
					terminationDates[emp.id] = emp.termination_date;
				}
			}

			hasMore = response.next !== null && response.next !== undefined;
			page++;
		} catch (e) {
			Logger.log('Error fetching termination dates: ' + e.message);
			hasMore = false;
		}
	}

	Logger.log(
		`Found ${
			Object.keys(terminationDates).length
		} employees with termination dates`
	);
	return terminationDates;
}
