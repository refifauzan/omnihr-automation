const OmniHRAPIClient = require('./apiClient');

// OmniHR API status codes
const LEAVE_STATUS_APPROVED = 3;

// OmniHR duration types: 1=full day, 2=half AM, 3=half PM
const DURATION_HALF_AM = 2;
const DURATION_HALF_PM = 3;

/**
 * Service for fetching and processing leave data from OmniHR API
 */
class LeaveService {
	constructor() {
		this.apiClient = new OmniHRAPIClient();
	}

	/**
	 * @param {Date} date
	 * @returns {string} Date in DD/MM/YYYY format
	 */
	formatDateDMY(date) {
		const day = String(date.getDate()).padStart(2, '0');
		const month = String(date.getMonth() + 1).padStart(2, '0');
		const year = date.getFullYear();
		return `${day}/${month}/${year}`;
	}

	/**
	 * @param {string} dateStr - Date string in DD/MM/YYYY format
	 * @returns {Date|null}
	 */
	parseDateDMY(dateStr) {
		if (!dateStr) return null;
		const parts = dateStr.split('/');
		if (parts.length !== 3) return null;
		return new Date(
			parseInt(parts[2]),
			parseInt(parts[1]) - 1,
			parseInt(parts[0])
		);
	}

	/**
	 * @param {Date} date
	 * @returns {boolean}
	 */
	isWeekend(date) {
		const day = date.getDay();
		return day === 0 || day === 6;
	}

	/**
	 * @param {number} duration - OmniHR duration type
	 * @returns {boolean}
	 */
	isHalfDayDuration(duration) {
		return duration === DURATION_HALF_AM || duration === DURATION_HALF_PM;
	}

	/**
	 * @param {Object} employee
	 * @returns {string}
	 */
	getEmployeeName(employee) {
		return (
			employee.full_name ||
			employee.name ||
			`User ${employee.id || employee.user_id}`
		);
	}

	/**
	 * @param {Object} employee
	 * @returns {number}
	 */
	getUserId(employee) {
		return employee.id || employee.user_id;
	}

	/**
	 * Fetches all employees with pagination
	 * @returns {Promise<Array>}
	 */
	async getAllEmployees() {
		const allEmployees = [];
		let page = 1;
		let hasMore = true;

		while (hasMore) {
			const response = await this.apiClient.get('/employee/list/', {
				page,
				page_size: 100,
			});
			const results = response.results || response;
			allEmployees.push(...results);
			hasMore = response.next != null;
			page++;
		}

		return allEmployees;
	}

	/**
	 * @param {number} userId
	 * @returns {Promise<Object>} Base data including employee_id (e.g., SM0068)
	 */
	async getEmployeeBaseData(userId) {
		const response = await this.apiClient.get(
			`/employee/2.0/users/${userId}/base-data/`
		);
		return response.data || response;
	}

	/**
	 * @param {number} userId
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
	 * @param {number} userId
	 * @param {Date} startDate
	 * @param {Date} endDate
	 * @returns {Promise<Object>} Calendar response with time_off_request array
	 */
	async getUserTimeOffCalendar(userId, startDate, endDate) {
		return this.apiClient.get(`/employee/1.1/${userId}/time-off-calendar/`, {
			start_date: this.formatDateDMY(startDate),
			end_date: this.formatDateDMY(endDate),
		});
	}

	/**
	 * @param {Date} startDate
	 * @param {Date} endDate
	 * @returns {Promise<Array>} All leave events with employee info
	 */
	async getAllLeaveEvents(startDate, endDate) {
		const employees = await this.getAllEmployees();
		const allLeaveEvents = [];

		for (const emp of employees) {
			const userId = this.getUserId(emp);
			const empName = this.getEmployeeName(emp);

			try {
				const events = await this.getUserTimeOffCalendar(
					userId,
					startDate,
					endDate
				);
				if (events?.length > 0) {
					const eventsWithUser = events.map((event) => ({
						...event,
						user_id: userId,
						employee_name: empName,
					}));
					allLeaveEvents.push(...eventsWithUser);
				}
			} catch {
				// Skip errors for individual users
			}
		}

		return allLeaveEvents;
	}

	/**
	 * First/last day uses effective_date_duration/end_date_duration, middle days are always full
	 * @param {Object} request - Leave request object
	 * @param {Date} currentDate
	 * @param {Date} leaveStart
	 * @param {Date} leaveEnd
	 * @returns {boolean}
	 */
	determineHalfDay(request, currentDate, leaveStart, leaveEnd) {
		const isFirstDay = currentDate.getTime() === leaveStart.getTime();
		const isLastDay = currentDate.getTime() === leaveEnd.getTime();

		if (isFirstDay) {
			return this.isHalfDayDuration(request.effective_date_duration);
		}

		if (isLastDay) {
			return this.isHalfDayDuration(request.end_date_duration);
		}

		return false;
	}

	/**
	 * @param {Date} date
	 * @param {number} month - 0-indexed month
	 * @param {number} year
	 * @returns {boolean}
	 */
	isInTargetMonth(date, month, year) {
		return date.getMonth() === month && date.getFullYear() === year;
	}

	/**
	 * @param {Object} request - Leave request from API
	 * @param {number} month - 0-indexed month
	 * @param {number} year
	 * @returns {Array} Array of leave day objects
	 */
	processLeaveRequest(request, month, year) {
		const leaveStart = this.parseDateDMY(request.effective_date);
		if (!leaveStart) return [];

		const leaveEnd = request.end_date
			? this.parseDateDMY(request.end_date)
			: leaveStart;

		const leaveDays = [];
		const currentDate = new Date(leaveStart);

		while (currentDate <= leaveEnd) {
			const shouldInclude =
				!this.isWeekend(currentDate) &&
				this.isInTargetMonth(currentDate, month, year);

			if (shouldInclude) {
				leaveDays.push({
					date: currentDate.getDate(),
					leave_type: request.time_off?.name,
					is_half_day: this.determineHalfDay(
						request,
						currentDate,
						leaveStart,
						leaveEnd
					),
				});
			}

			currentDate.setDate(currentDate.getDate() + 1);
		}

		return leaveDays;
	}

	/**
	 * Only process approved leave requests (status=3)
	 * @param {Object} calendarResponse
	 * @param {number} month - 0-indexed month
	 * @param {number} year
	 * @returns {Array}
	 */
	processCalendarResponse(calendarResponse, month, year) {
		const timeOffRequests = calendarResponse?.time_off_request || [];
		const approvedRequests = timeOffRequests.filter(
			(r) => r.status === LEAVE_STATUS_APPROVED
		);

		return approvedRequests.flatMap((request) =>
			this.processLeaveRequest(request, month, year)
		);
	}

	/**
	 * Fetch base data, leave balances, and calendar in parallel for performance
	 * @param {Object} employee
	 * @param {Date|null} startDate
	 * @param {Date|null} endDate
	 * @param {number} month
	 * @param {number} year
	 * @returns {Promise<Object>} Employee leave data
	 */
	async processEmployee(employee, startDate, endDate, month, year) {
		const userId = this.getUserId(employee);
		const employeeName = this.getEmployeeName(employee);

		try {
			const promises = [
				this.getEmployeeBaseData(userId),
				this.getEmployeeLeaveBalances(userId),
			];

			if (startDate && endDate) {
				promises.push(this.getUserTimeOffCalendar(userId, startDate, endDate));
			}

			const [baseData, leaveBalances, calendarResponse] = await Promise.all(
				promises
			);

			const employeeData = {
				user_id: userId,
				employee_id: baseData?.employee_id,
				employee_name: employeeName,
				leave_balances: leaveBalances,
			};

			if (calendarResponse) {
				employeeData.leave_requests = this.processCalendarResponse(
					calendarResponse,
					month,
					year
				);
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
	}

	/**
	 * Main entry point: fetches all employees and processes in batches
	 * @param {Object} options
	 * @param {Function} [options.onProgress] - Progress callback (completed, total, lastEmployeeName)
	 * @param {number} [options.month] - 0-indexed month
	 * @param {number} [options.year]
	 * @param {number} [options.concurrency=5] - Number of parallel requests
	 * @returns {Promise<Array>} Array of employee leave data
	 */
	async getAllLeaveData(options = {}) {
		const { onProgress, month, year, concurrency = 5 } = options;

		const employees = await this.getAllEmployees();

		const hasDateRange = month !== undefined && year !== undefined;
		const startDate = hasDateRange ? new Date(year, month, 1) : null;
		const endDate = hasDateRange ? new Date(year, month + 1, 0) : null;

		const allLeaveData = [];
		let completed = 0;

		for (let i = 0; i < employees.length; i += concurrency) {
			const batch = employees.slice(i, i + concurrency);

			const results = await Promise.all(
				batch.map((emp) =>
					this.processEmployee(emp, startDate, endDate, month, year)
				)
			);

			allLeaveData.push(...results);
			completed += batch.length;

			if (onProgress) {
				const lastEmployee = batch[batch.length - 1];
				onProgress(completed, employees.length, lastEmployee?.full_name || '');
			}
		}

		return allLeaveData;
	}
}

module.exports = LeaveService;
