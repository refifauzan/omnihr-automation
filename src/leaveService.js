const OmniHRAPIClient = require('./apiClient');

class LeaveService {
	constructor() {
		this.apiClient = new OmniHRAPIClient();
	}

	/**
	 * Get all employees from the API with pagination
	 * @returns {Promise<Array>} Array of all employee objects
	 */
	async getAllEmployees() {
		let allEmployees = [];
		let page = 1;
		let hasMore = true;

		while (hasMore) {
			const response = await this.apiClient.get('/employee/list/', {
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
	 * Get employee base data including employee_id (e.g., SM0068)
	 * @param {number} userId - User ID
	 * @returns {Promise<Object>} Base data object with employee_id
	 */
	async getEmployeeBaseData(userId) {
		const response = await this.apiClient.get(
			`/employee/2.0/users/${userId}/base-data/`
		);
		return response.data || response;
	}

	/**
	 * Get leave balances for a specific employee
	 * @param {string} userId - The user ID to fetch leave balances for
	 * @returns {Promise<Array>} Array of leave balance objects
	 */
	async getEmployeeLeaveBalances(userId) {
		const timeOffTypes = await this.apiClient.get(
			`/employee/1.1/users/${userId}/time-off-types/`
		);

		return timeOffTypes.map((t) => ({
			leave_type: t.time_off?.name,
			entitlement: t.time_off_balance?.entitlement_earned,
			taken: t.time_off_balance?.display_taken,
			remaining: t.time_off_balance?.days,
		}));
	}

	/**
	 * Get time-off calendar events for a user in a date range
	 * @param {number} userId - User ID
	 * @param {Date} startDate - Start date
	 * @param {Date} endDate - End date
	 * @returns {Promise<Array>} Array of time-off calendar events
	 */
	async getUserTimeOffCalendar(userId, startDate, endDate) {
		// Format: DD/MM/YYYY as required by the API
		const formatDate = (d) => {
			const day = String(d.getDate()).padStart(2, '0');
			const month = String(d.getMonth() + 1).padStart(2, '0');
			const year = d.getFullYear();
			return `${day}/${month}/${year}`;
		};

		const start = formatDate(startDate);
		const end = formatDate(endDate);

		const response = await this.apiClient.get(
			`/employee/1.1/${userId}/time-off-calendar/`,
			{ start_date: start, end_date: end }
		);

		return response;
	}

	/**
	 * Get all leave events for all employees in a date range
	 * @param {Date} startDate - Start date
	 * @param {Date} endDate - End date
	 * @returns {Promise<Array>} Array of leave events with employee info
	 */
	async getAllLeaveEvents(startDate, endDate) {
		const employees = await this.getAllEmployees();
		const allLeaveEvents = [];

		for (const emp of employees) {
			const userId = emp.id || emp.user_id;
			const empName = emp.full_name || emp.name || `User ${userId}`;

			try {
				const events = await this.getUserTimeOffCalendar(
					userId,
					startDate,
					endDate
				);

				if (events && events.length > 0) {
					events.forEach((event) => {
						allLeaveEvents.push({
							...event,
							user_id: userId,
							employee_name: empName,
						});
					});
				}
			} catch (err) {
				// Skip errors for individual users
			}
		}

		return allLeaveEvents;
	}

	/**
	 * Get all employee leave data with optional monthly leave requests
	 * Uses parallel requests for faster fetching
	 * @param {Object} options - Configuration options
	 * @param {Function} options.onProgress - Callback function for progress updates
	 * @param {number} options.month - Month (0-11) to fetch leave requests for
	 * @param {number} options.year - Year to fetch leave requests for
	 * @param {number} options.concurrency - Number of parallel requests (default: 5)
	 * @returns {Promise<Array>} Array of employee leave data objects
	 */
	async getAllLeaveData(options = {}) {
		const { onProgress, month, year, concurrency = 5 } = options;

		const employees = await this.getAllEmployees();

		// Determine date range for monthly leave requests
		let startDate = null;
		let endDate = null;
		if (month !== undefined && year !== undefined) {
			startDate = new Date(year, month, 1);
			endDate = new Date(year, month + 1, 0); // Last day of month
		}

		// Helper to parse DD/MM/YYYY to Date
		const parseDateDMY = (dateStr) => {
			if (!dateStr) return null;
			const parts = dateStr.split('/');
			if (parts.length !== 3) return null;
			return new Date(
				parseInt(parts[2]),
				parseInt(parts[1]) - 1,
				parseInt(parts[0])
			);
		};

		// Process a single employee
		const processEmployee = async (employee) => {
			const userId = employee.id || employee.user_id;
			const employeeName =
				employee.full_name || employee.name || `User ${userId}`;

			try {
				// Fetch base data, leave balances, and calendar in parallel
				const promises = [
					this.getEmployeeBaseData(userId),
					this.getEmployeeLeaveBalances(userId),
				];
				if (startDate && endDate) {
					promises.push(
						this.getUserTimeOffCalendar(userId, startDate, endDate)
					);
				}

				const results = await Promise.all(promises);
				const baseData = results[0];
				const leaveBalances = results[1];
				const calendarResponse = results[2];

				const employeeData = {
					user_id: userId,
					employee_id: baseData?.employee_id,
					employee_name: employeeName,
					leave_balances: leaveBalances,
				};

				// Process leave requests if calendar was fetched
				if (calendarResponse) {
					const leaveDays = [];
					const approvedRequests = (
						calendarResponse.time_off_request || []
					).filter((r) => r.status === 3);

					for (const r of approvedRequests) {
						const leaveStart = parseDateDMY(r.effective_date);
						const leaveEnd = r.end_date ? parseDateDMY(r.end_date) : leaveStart;

						if (!leaveStart) continue;

						// Iterate through each day in the leave range
						const currentDate = new Date(leaveStart);
						while (currentDate <= leaveEnd) {
							const dayOfWeek = currentDate.getDay();
							const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

							if (
								!isWeekend &&
								currentDate.getMonth() === month &&
								currentDate.getFullYear() === year
							) {
								leaveDays.push({
									date: currentDate.getDate(),
									leave_type: r.time_off?.name,
									is_half_day:
										r.effective_date_duration === 2 ||
										r.effective_date_duration === 3,
								});
							}

							// Move to next day
							currentDate.setDate(currentDate.getDate() + 1);
						}
					}
					employeeData.leave_requests = leaveDays;
				}

				return employeeData;
			} catch (err) {
				return {
					user_id: userId,
					employee_name: employeeName,
					leave_balances: [],
					leave_requests: [],
					error: err.message,
				};
			}
		};

		// Process employees in batches for parallel execution
		const allLeaveData = [];
		let completed = 0;

		for (let i = 0; i < employees.length; i += concurrency) {
			const batch = employees.slice(i, i + concurrency);
			const results = await Promise.all(batch.map(processEmployee));
			allLeaveData.push(...results);

			completed += batch.length;
			if (onProgress) {
				onProgress(
					completed,
					employees.length,
					batch[batch.length - 1]?.full_name || ''
				);
			}
		}

		return allLeaveData;
	}
}

module.exports = LeaveService;
