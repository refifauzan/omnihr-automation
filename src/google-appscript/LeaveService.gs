/**
 * Leave data fetching and processing functions
 */

/**
 * Fetch leave data for all employees using batch requests
 * @param {string} token - Access token
 * @param {Array} employees - Employee list
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @returns {Object} Leave data keyed by employee ID
 */
function fetchLeaveDataForMonth(token, employees, month, year) {
  const startDate = new Date(year, month, 1);
  const endDate = new Date(year, month + 1, 0);

  const leaveData = {};
  const BATCH_SIZE = 50;

  for (let i = 0; i < employees.length; i += BATCH_SIZE) {
    const batch = employees.slice(i, i + BATCH_SIZE);
    Logger.log(
      `Processing batch ${Math.floor(i / BATCH_SIZE) + 1}/${Math.ceil(
        employees.length / BATCH_SIZE,
      )} (${batch.length} employees)`,
    );

    const { requests, requestMeta } = buildBatchRequests(
      token,
      batch,
      startDate,
      endDate,
    );

    const responses = UrlFetchApp.fetchAll(requests);

    // Group responses by employee
    const employeeData = {};
    for (let j = 0; j < responses.length; j++) {
      const meta = requestMeta[j];
      const response = responses[j];

      try {
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        if (responseCode !== 200) {
          Logger.log(
            `API error for ${meta.empName} (${
              meta.type
            }): HTTP ${responseCode} - ${responseText.substring(0, 200)}`,
          );
          continue;
        }

        const data = JSON.parse(responseText);

        if (data.error || data.detail || data.message) {
          Logger.log(
            `API error for ${meta.empName}: ${
              data.error || data.detail || data.message
            }`,
          );
          continue;
        }

        if (!employeeData[meta.userId]) {
          employeeData[meta.userId] = { empName: meta.empName };
        }

        if (meta.type === "base") {
          employeeData[meta.userId].baseData = data.data || data;
        } else {
          employeeData[meta.userId].calendar = data;
        }
      } catch (e) {
        Logger.log(`Error parsing response for ${meta.empName}: ${e.message}`);
      }
    }

    // Process the collected data
    for (const [userId, data] of Object.entries(employeeData)) {
      const employeeId = data.baseData?.employee_id;
      const empName = data.empName;
      const calendar = data.calendar || {};

      const approvedRequests = (calendar.time_off_request || []).filter(
        (r) => r.status === 3,
      );

      if (approvedRequests.length === 0) continue;

      const leaveDays = [];

      for (const request of approvedRequests) {
        const leaveStart = parseDateDMY(request.effective_date);
        const leaveEnd = request.end_date
          ? parseDateDMY(request.end_date)
          : leaveStart;

        if (!leaveStart) continue;

        const currentDate = new Date(leaveStart);
        while (currentDate <= leaveEnd) {
          const dayOfWeek = currentDate.getDay();
          const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
          const dateToProcess = new Date(currentDate);

          currentDate.setDate(currentDate.getDate() + 1);

          if (isWeekend) continue;
          if (dateToProcess.getMonth() !== month) continue;
          if (dateToProcess.getFullYear() !== year) continue;

          const isFirstDay = dateToProcess.getTime() === leaveStart.getTime();
          const isLastDay = dateToProcess.getTime() === leaveEnd.getTime();
          const isSingleDay = isFirstDay && isLastDay;

          const effectiveDuration =
            parseInt(request.effective_date_duration) || 1;
          const endDuration = parseInt(request.end_date_duration) || 1;

          const isHalfDay = determineHalfDay(
            isSingleDay,
            isFirstDay,
            isLastDay,
            effectiveDuration,
            endDuration,
          );

          Logger.log(
            `Leave for ${empName} on day ${dateToProcess.getDate()}: effectiveDuration=${effectiveDuration}, endDuration=${endDuration}, isHalfDay=${isHalfDay}`,
          );

          leaveDays.push({
            date: dateToProcess.getDate(),
            leave_type: request.time_off?.name,
            is_half_day: isHalfDay,
          });
        }
      }

      if (leaveDays.length > 0) {
        leaveData[employeeId || empName] = {
          employee_id: employeeId,
          employee_name: empName,
          leave_requests: leaveDays,
        };
      }
    }
  }

  return leaveData;
}

/**
 * Update sheet with leave data
 * @param {Sheet} sheet - The sheet
 * @param {Object} leaveData - Leave data
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {boolean} useSheetHours - If true, read hours from sheet cells
 * @param {Set} holidayDays - Optional set of holiday day numbers to exclude from total
 */
function updateSheetWithLeaveData(
  sheet,
  leaveData,
  month,
  year,
  useSheetHours = false,
  holidayDays = new Set(),
) {
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const employeeLookup = buildEmployeeLookup(sheet);

  let attendanceByEmployee = {};
  if (!useSheetHours) {
    const attendanceList = getAttendanceData() || [];
    for (const att of attendanceList) {
      const key = att.empId;
      if (!attendanceByEmployee[key]) {
        attendanceByEmployee[key] = [];
      }
      attendanceByEmployee[key].push({
        project: att.project,
        hours: att.hours,
      });
    }
  }

  const { dayColumns, weekRanges } = calculateDayColumns(month, year);

  // Build day -> override column mapping
  const dayToOverrideCol = {};
  for (const weekRange of weekRanges) {
    for (const [dayStr, col] of Object.entries(dayColumns)) {
      if (col >= weekRange.startCol && col <= weekRange.endCol) {
        dayToOverrideCol[dayStr] = weekRange.overrideCol;
      }
    }
  }

  // Clear existing leave markings before applying new ones (respects Time off Override)
  Logger.log("Clearing existing leave markings...");
  clearLeaveCellsRespectingOverride(
    sheet,
    dayColumns,
    dayToOverrideCol,
    month,
    year,
    holidayDays,
  );

  // Set default hours for Operations team (8 hours on working days without existing values)
  Logger.log("Setting default hours for Operations team...");
  setOperationsDefaultHours(sheet, dayColumns, month, year, holidayDays);

  const fullDayCells = [];
  const halfDayCellsMap = {};
  let matchedEmployees = 0;

  for (const [key, empData] of Object.entries(leaveData)) {
    const { employee_id, employee_name, leave_requests } = empData;

    const rows = findEmployeeRows(employeeLookup, employee_id, employee_name);
    if (rows.length === 0) {
      Logger.log(`Employee not found: ${employee_name} (${employee_id})`);
      continue;
    }

    matchedEmployees++;
    Logger.log(
      `Found ${
        rows.length
      } rows for ${employee_name} (ID: ${employee_id}) at rows: ${rows.join(
        ", ",
      )}`,
    );

    // Build row -> hours mapping
    const rowHoursMap = {};

    const rowHoursSource = useSheetHours
      ? buildRowHoursFromSheet(sheet, rows, dayColumns)
      : buildRowHoursFromAttendance(
          sheet,
          rows,
          attendanceByEmployee[employee_id.toUpperCase()] || [],
        );

    for (const [row, hours] of Object.entries(rowHoursSource)) {
      rowHoursMap[row] = hours;
    }

    // Apply leaves
    for (const leave of leave_requests) {
      const col = dayColumns[leave.date];
      if (!col) continue;

      // Skip visual marking if it's a public holiday
      if (holidayDays.has(leave.date)) {
        Logger.log(
          `Skipping leave marking for ${employee_name} on day ${leave.date} - it's a public holiday`,
        );
        continue;
      }

      Logger.log(
        `Processing leave for ${employee_name} on day ${leave.date}, is_half_day: ${leave.is_half_day}`,
      );

      const overrideCol = dayToOverrideCol[leave.date];

      const activeRows = getActiveRows(sheet, rows, overrideCol, leave.date);

      if (activeRows.length === 0) {
        Logger.log(
          `All rows have Week Override checked for day ${leave.date}, skipping`,
        );
        continue;
      }

      assignLeaveCells(leave, activeRows, col, fullDayCells, halfDayCellsMap);
    }
  }

  // Apply leave values and colors
  if (fullDayCells.length > 0) {
    Logger.log(
      `Applying full-day leave to ${fullDayCells.length} cells (value 0, red)`,
    );
    const fullDayRanges = sheet.getRangeList(fullDayCells);
    fullDayRanges.setValue(0);
    fullDayRanges.setBackground(CONFIG.COLORS.FULL_DAY);
    fullDayRanges.setFontColor("#FFFFFF");
    fullDayRanges.setFontWeight("bold");
  }

  const halfDayCellEntries = Object.entries(halfDayCellsMap);
  if (halfDayCellEntries.length > 0) {
    Logger.log(
      `Applying half-day leave to ${halfDayCellEntries.length} cells (orange)`,
    );
    for (const [cellA1, value] of halfDayCellEntries) {
      const range = sheet.getRange(cellA1);
      range.setValue(parseFloat(value));
      range.setBackground(CONFIG.COLORS.HALF_DAY);
      range.setFontColor("#000000");
      range.setFontWeight("bold");
    }
  }

  // Calculate and update Total Days Off (column I) for each employee
  const totalDaysOffMap = {};
  for (const [key, empData] of Object.entries(leaveData)) {
    const { employee_id, employee_name, leave_requests } = empData;

    // Calculate total working days off (excluding weekends AND holidays, counting half-days as 0.5)
    let totalDaysOff = 0;
    for (const leave of leave_requests) {
      // Skip if leave falls on a public holiday
      if (holidayDays.has(leave.date)) {
        Logger.log(
          `Skipping leave on day ${leave.date} - it's a public holiday`,
        );
        continue;
      }

      // Skip if leave falls on a weekend
      const date = new Date(year, month, leave.date);
      const dayOfWeek = date.getDay();
      if (dayOfWeek === 0 || dayOfWeek === 6) {
        Logger.log(`Skipping leave on day ${leave.date} - it's a weekend`);
        continue;
      }

      if (!leave.is_half_day) {
        totalDaysOff += 1;
      } else {
        totalDaysOff += 0.5;
      }
    }

    const rows = findEmployeeRows(employeeLookup, employee_id, employee_name);
    for (const row of rows) {
      totalDaysOffMap[row] = totalDaysOff;
    }
  }

  // Update column I with calculated values for ALL employees
  // First, set 0 for all employee rows (to clear old values for employees without leave)
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;
  if (numRows > 0) {
    // Initialize all employees to 0 days off
    const zeroValues = Array(numRows).fill([0]);
    sheet.getRange(CONFIG.FIRST_DATA_ROW, 9, numRows, 1).setValues(zeroValues);
    Logger.log(`Initialized Total Days Off to 0 for ${numRows} employee rows`);
  }

  // Now update with actual leave values for employees who have leave
  if (Object.keys(totalDaysOffMap).length > 0) {
    Logger.log(
      `Updating Total Days Off for ${
        Object.keys(totalDaysOffMap).length
      } employees with leave`,
    );
    for (const [rowStr, daysOff] of Object.entries(totalDaysOffMap)) {
      const row = parseInt(rowStr);
      sheet.getRange(row, 9).setValue(daysOff);
    }
  }

  // Add conditional formatting
  const newLeaveCells = [...fullDayCells, ...Object.keys(halfDayCellsMap)];
  const existingLeaveCells = scanForLeaveCells(sheet, dayColumns);
  const allLeaveCells = [...new Set([...newLeaveCells, ...existingLeaveCells])];
  Logger.log(
    `Total leave cells to exclude: ${allLeaveCells.length} (${newLeaveCells.length} new, ${existingLeaveCells.length} existing)`,
  );
  addValidatedConditionalFormatting(sheet, month, year, allLeaveCells);

  SpreadsheetApp.flush();

  Logger.log(
    `Matched ${matchedEmployees} employees, updated ${
      fullDayCells.length + halfDayCellEntries.length
    } cells`,
  );
}

/**
 * Apply leave colors to sheet without modifying attendance data
 * @param {Sheet} sheet - The sheet
 * @param {Object} leaveData - Leave data
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 */
function applyLeaveColorsToSheet(sheet, leaveData, month, year) {
  const { dayColumns, weekRanges } = calculateDayColumns(month, year);

  const dayToOverrideCol = {};
  for (const weekRange of weekRanges) {
    for (const [dayStr, col] of Object.entries(dayColumns)) {
      if (col >= weekRange.startCol && col <= weekRange.endCol) {
        dayToOverrideCol[dayStr] = weekRange.overrideCol;
      }
    }
  }

  const lastRow = sheet.getLastRow();
  const employeeLookup = {};

  if (lastRow >= CONFIG.FIRST_DATA_ROW) {
    const employeeRange = sheet.getRange(
      CONFIG.FIRST_DATA_ROW,
      1,
      lastRow - CONFIG.FIRST_DATA_ROW + 1,
      3,
    );
    const employeeValues = employeeRange.getValues();

    for (let i = 0; i < employeeValues.length; i++) {
      const row = CONFIG.FIRST_DATA_ROW + i;
      const empId = String(employeeValues[i][0] || "")
        .trim()
        .toUpperCase();
      const empName = String(employeeValues[i][1] || "")
        .trim()
        .toUpperCase();

      if (empId) {
        if (!employeeLookup[empId]) employeeLookup[empId] = [];
        employeeLookup[empId].push(row);
      }
      if (empName) {
        if (!employeeLookup[empName]) employeeLookup[empName] = [];
        if (!employeeLookup[empName].includes(row)) {
          employeeLookup[empName].push(row);
        }
      }
    }
  }

  const fullDayCells = [];
  const halfDayCells = [];
  let matchedEmployees = 0;

  for (const [key, empData] of Object.entries(leaveData)) {
    const { employee_id, employee_name, leave_requests } = empData;

    const rows = findEmployeeRows(employeeLookup, employee_id, employee_name);
    if (rows.length === 0) {
      Logger.log(`Employee not found: ${employee_name} (${employee_id})`);
      continue;
    }

    matchedEmployees++;
    const numRows = rows.length;
    Logger.log(`Found ${numRows} rows for ${employee_name}`);

    const rowHoursMap = {};
    const allDayCols = Object.values(dayColumns);

    for (const row of rows) {
      let foundHours = null;

      for (const col of allDayCols) {
        const cellValue = sheet.getRange(row, col).getValue();
        const numValue = parseFloat(cellValue);

        if (!isNaN(numValue) && numValue > 0) {
          foundHours = numValue;
          break;
        }
      }

      rowHoursMap[row] = foundHours || CONFIG.DEFAULT_HOURS;
      Logger.log(`Row ${row}: found ${rowHoursMap[row]} hours`);
    }

    for (const leave of leave_requests) {
      const col = dayColumns[leave.date];
      if (!col) continue;

      const overrideCol = dayToOverrideCol[leave.date];

      const activeRows = getActiveRows(sheet, rows, overrideCol, leave.date);
      if (activeRows.length === 0) continue;

      assignLeaveCellsWithObjects(
        leave,
        activeRows,
        col,
        fullDayCells,
        halfDayCells,
      );
    }
  }

  if (fullDayCells.length > 0) {
    Logger.log(`Applying red to ${fullDayCells.length} full-day leave cells`);
    const fullDayRanges = sheet.getRangeList(fullDayCells);
    fullDayRanges.setValue(0);
    fullDayRanges.setBackground(CONFIG.COLORS.FULL_DAY);
    fullDayRanges.setFontColor("#FFFFFF");
    fullDayRanges.setFontWeight("bold");
  }

  if (halfDayCells.length > 0) {
    Logger.log(
      `Applying orange to ${halfDayCells.length} half-day leave cells`,
    );
    for (const { cell, value } of halfDayCells) {
      const range = sheet.getRange(cell);
      range.setValue(value);
      range.setBackground(CONFIG.COLORS.HALF_DAY);
      range.setFontColor("#000000");
      range.setFontWeight("bold");
    }
  }

  Logger.log(
    `Applied leave colors: ${matchedEmployees} employees, ${fullDayCells.length} full-day, ${halfDayCells.length} half-day`,
  );
}

/**
 * Add conditional formatting for validated weeks
 * @param {Sheet} sheet - The sheet
 * @param {number} month - Month (0-11)
 * @param {number} year - Year
 * @param {Array} leaveCells - Cells to exclude from green formatting
 */
function addValidatedConditionalFormatting(
  sheet,
  month,
  year,
  leaveCells = [],
) {
  const { dayColumns, weekRanges } = calculateDayColumns(month, year);
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - CONFIG.FIRST_DATA_ROW + 1;

  if (numRows <= 0) {
    Logger.log("addValidatedConditionalFormatting: Invalid range dimensions");
    return;
  }

  const leaveCellSet = new Set(leaveCells.map((c) => c.toUpperCase()));
  Logger.log(
    `Excluding ${leaveCellSet.size} leave cells from green formatting`,
  );

  sheet.setConditionalFormatRules([]);

  const rules = [];

  for (const weekRange of weekRanges) {
    const rule = buildWeekConditionalRule(
      sheet,
      weekRange,
      dayColumns,
      month,
      year,
      lastRow,
      numRows,
      leaveCellSet,
    );
    if (rule) rules.push(rule);
  }

  sheet.setConditionalFormatRules(rules);
  Logger.log(
    `Applied ${rules.length} conditional formatting rules (validated weeks only)`,
  );
}
