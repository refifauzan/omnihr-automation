/**
 * OmniHR API functions
 */

/**
 * Get access token from the remote Token Service.
 * The Token Service handles Google SSO login automatically.
 * @returns {string} Access token
 */
function getAccessToken() {
	const props = PropertiesService.getScriptProperties();
	const tokenServiceUrl = props.getProperty('TOKEN_SERVICE_URL');
	const tokenServiceKey = props.getProperty('TOKEN_SERVICE_API_KEY');

	if (!tokenServiceUrl || !tokenServiceKey) {
		throw new Error(
			'Token Service not configured. Use OmniHR > Setup Token Service',
		);
	}

	const url = tokenServiceUrl.replace(/\/+$/, '') + '/api/token';

	const response = UrlFetchApp.fetch(url, {
		method: 'get',
		headers: {
			'X-API-Key': tokenServiceKey,
		},
		muteHttpExceptions: true,
	});

	const code = response.getResponseCode();
	const text = response.getContentText();

	if (code < 200 || code >= 300) {
		throw new Error('Token Service returned ' + code + ': ' + text);
	}

	const data = JSON.parse(text);
	const token = data.access_token;
	if (!token) {
		throw new Error('Token Service response missing access_token: ' + text);
	}

	Logger.log('Access token obtained from Token Service');
	return token;
}

/**
 * Ask the Token Service to force-refresh the token.
 * Called when an OmniHR API request returns 401.
 * @private
 * @returns {string} A guaranteed-fresh access token
 */
function forceRefreshTokenFromService_() {
	const props = PropertiesService.getScriptProperties();
	const tokenServiceUrl = props.getProperty('TOKEN_SERVICE_URL');
	const tokenServiceKey = props.getProperty('TOKEN_SERVICE_API_KEY');

	if (!tokenServiceUrl || !tokenServiceKey) return null;

	const url = tokenServiceUrl.replace(/\/+$/, '') + '/api/force-fresh-token';

	const response = UrlFetchApp.fetch(url, {
		method: 'post',
		headers: {
			'X-API-Key': tokenServiceKey,
		},
		muteHttpExceptions: true,
	});

	const code = response.getResponseCode();
	const text = response.getContentText();

	if (code < 200 || code >= 300) {
		throw new Error(
			'Token Service force-refresh failed (' + code + '): ' + text,
		);
	}

	const data = JSON.parse(text);
	const token = data.access_token;
	if (!token) {
		throw new Error('Token Service force-refresh returned no token: ' + text);
	}

	Logger.log('Obtained force-refreshed token from Token Service');
	return token;
}

/**
 * Make authenticated API request with automatic token retry on 401.
 * If the request fails with 401 (token expired), it asks the Token Service
 * for a force-refreshed token and retries the request once.
 * @param {string} token - Access token
 * @param {string} endpoint - API endpoint
 * @param {Object} params - Query parameters
 * @returns {Object} Parsed JSON response
 */
function apiRequest(token, endpoint, params = {}) {
	const props = PropertiesService.getScriptProperties();
	const baseUrl = props.getProperty('OMNIHR_BASE_URL');
	const subdomain = props.getProperty('OMNIHR_SUBDOMAIN');

	let url = baseUrl + endpoint;

	if (Object.keys(params).length > 0) {
		const queryString = Object.entries(params)
			.map(function (entry) {
				return (
					encodeURIComponent(entry[0]) + '=' + encodeURIComponent(entry[1])
				);
			})
			.join('&');
		url += '?' + queryString;
	}

	const response = UrlFetchApp.fetch(url, {
		method: 'get',
		headers: {
			Authorization: 'Bearer ' + token,
			'x-subdomain': subdomain,
			'Content-Type': 'application/json',
		},
		muteHttpExceptions: true,
	});

	const code = response.getResponseCode();

	// ── Auto-retry on 401 (token expired) ────────────────────────
	if (code === 401) {
		Logger.log(
			'apiRequest 401 on ' + endpoint + ' — attempting token refresh...',
		);
		try {
			const freshToken = forceRefreshTokenFromService_();
			if (freshToken) {
				Logger.log('Retrying ' + endpoint + ' with fresh token...');
				const retryResponse = UrlFetchApp.fetch(url, {
					method: 'get',
					headers: {
						Authorization: 'Bearer ' + freshToken,
						'x-subdomain': subdomain,
						'Content-Type': 'application/json',
					},
					muteHttpExceptions: true,
				});
				const retryCode = retryResponse.getResponseCode();
				if (retryCode === 200) {
					Logger.log('Retry succeeded for ' + endpoint);
					return JSON.parse(retryResponse.getContentText());
				}
				Logger.log('Retry also failed (' + retryCode + ') for ' + endpoint);
			}
		} catch (e) {
			Logger.log('Token force-refresh failed: ' + e.message);
		}
	}

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
			startDate,
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
			{ start_date: startDateStr, end_date: endDateStr },
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
						}`,
					);
				}
			}
		}

		Logger.log(
			`Found ${holidayDays.length} public holidays for ${month + 1}/${year}`,
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
 * @param {Object} [options] - Optional; set includeMergedTerminatedNotInList: false for Sync Employee List (active only, no workflow-only terminated rows)
 * @returns {Array} Array of employee objects with full details
 */
function fetchAllEmployeesWithDetails(token, options) {
	options = options || {};
	const includeMergedTerminated =
		options.includeMergedTerminatedNotInList !== false;

	const allEmployees = fetchAllEmployees(token);

	// Filter out excluded employees
	const employees = allEmployees.filter((emp) => {
		const fullName = emp.full_name || emp.name || '';
		const isExcluded = EXCLUDED_EMPLOYEES.some(
			(excluded) => fullName.toLowerCase() === excluded.toLowerCase(),
		);
		if (isExcluded) {
			Logger.log(`Excluding employee: ${fullName}`);
		}
		return !isExcluded;
	});

	Logger.log(
		`Filtered ${allEmployees.length - employees.length} excluded employees`,
	);

	const employeeDetails = [];
	const BATCH_SIZE = 50;

	// Fetch termination data (dates + full employee objects) from workflow-dashboard
	const termData = fetchTerminationData(token);
	const terminationDates = termData.dates;
	const terminatedEmployees = termData.employees;

	// Build a set of active user IDs so we can identify terminated-only employees later
	const activeUserIds = new Set(
		employees.map(function (e) {
			return e.id || e.user_id;
		}),
	);

	// Fetch team and project contribution data for all employees
	const { teamData, projectContribution } = fetchEmployeeJobData(
		token,
		employees,
	);

	for (let i = 0; i < employees.length; i += BATCH_SIZE) {
		const batch = employees.slice(i, i + BATCH_SIZE);
		Logger.log(
			`Fetching employee details batch ${
				Math.floor(i / BATCH_SIZE) + 1
			}/${Math.ceil(employees.length / BATCH_SIZE)}`,
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

	// Merge terminated employees from workflow-dashboard that are no longer in /employee/list/
	// (Used for hire/termination grey-out — not for Sync Employee List; omit when includeMergedTerminatedNotInList is false.)
	let mergedCount = 0;
	if (includeMergedTerminated) {
		for (const termEmp of terminatedEmployees) {
			const termUserId = termEmp.id || termEmp.user_id;
			if (!activeUserIds.has(termUserId)) {
				employeeDetails.push({
					user_id: termUserId,
					employee_id: termEmp.employee_id || '',
					full_name: termEmp.full_name || termEmp.name || `User ${termUserId}`,
					hired_date: termEmp.hired_date || null,
					termination_date: terminationDates[termUserId] || null,
					team: teamData[termUserId] || '',
					project_contribution: projectContribution[termUserId] || '',
					employment_status: termEmp.employment_status || null,
					employment_status_display: termEmp.employment_status_display || null,
				});
				mergedCount++;
			}
		}
		if (mergedCount > 0) {
			Logger.log(
				`Merged ${mergedCount} terminated employees not in /employee/list/`,
			);
		}
	}

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
		'Fetching job data (team & project contribution) for employees...',
	);

	for (let i = 0; i < employees.length; i += BATCH_SIZE) {
		const batch = employees.slice(i, i + BATCH_SIZE);
		Logger.log(
			`Fetching job data batch ${Math.floor(i / BATCH_SIZE) + 1}/${Math.ceil(
				employees.length / BATCH_SIZE,
			)}`,
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
							(attr) => attr.attr === PROJECT_CONTRIBUTION_ATTR_ID,
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
		} employees`,
	);
	return { teamData, projectContribution };
}

/**
 * Fetch termination data from onboarding/workflow-dashboard endpoint
 * Returns both a date map and full employee objects for terminated employees
 * @param {string} token - Access token
 * @returns {Object} { dates: Map of user_id -> termination_date, employees: Array of workflow-dashboard entries with termination_date }
 */
function fetchTerminationData(token) {
	const dates = {};
	const employees = [];
	let page = 1;
	let hasMore = true;

	Logger.log('Fetching termination data from workflow-dashboard...');

	while (hasMore) {
		try {
			const response = apiRequest(token, '/onboarding/workflow-dashboard/', {
				page: page,
				page_size: 100,
			});

			const results = response.results || [];
			for (const emp of results) {
				if (emp.termination_date) {
					dates[emp.id] = emp.termination_date;
					employees.push(emp);
				}
			}

			hasMore = response.next !== null && response.next !== undefined;
			page++;
		} catch (e) {
			Logger.log('Error fetching termination data: ' + e.message);
			hasMore = false;
		}
	}

	Logger.log(`Found ${employees.length} employees with termination dates`);
	return { dates: dates, employees: employees };
}

/**
 * Fetch termination dates from onboarding/workflow-dashboard endpoint
 * Backward-compatible wrapper around fetchTerminationData
 * @param {string} token - Access token
 * @returns {Object} Map of user_id -> termination_date
 */
function fetchTerminationDates(token) {
	return fetchTerminationData(token).dates;
}
