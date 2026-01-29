/**
 * Detect Hire and Termination Employees
 *
 * Mirrors the logic from google-appscript (EmployeeSync.gs, Api.gs):
 * - Fetches employees from /employee/list/ (hired_date)
 * - Fetches termination dates from /onboarding/workflow-dashboard/
 * - Excludes: Omni Support, People Culture
 * - Outputs: employees with hire date, terminated employees, and optional recent hires/terminations
 */

require("dotenv").config();
const fs = require("fs");
const path = require("path");
const OmniHRAPIClient = require("./apiClient");

const EXCLUDED_EMPLOYEES = ["Omni Support", "People Culture"];

/**
 * Parse date string in DD/MM/YYYY format (API format)
 * @param {string} dateStr
 * @returns {Date|null}
 */
function parseDateDMY(dateStr) {
  if (!dateStr) return null;
  const parts = String(dateStr).split("/");
  if (parts.length !== 3) return null;
  const day = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1;
  const year = parseInt(parts[2], 10);
  if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
  const d = new Date(year, month, day);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Format date for display
 * @param {Date} d
 * @returns {string}
 */
function formatDate(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return "—";
  return d.toISOString().slice(0, 10);
}

/**
 * Fetch all employees from /employee/list/ (paginated)
 * @param {OmniHRAPIClient} client
 * @returns {Promise<Array>}
 */
async function fetchAllEmployees(client) {
  const all = [];
  let page = 1;
  let hasMore = true;

  while (hasMore) {
    const response = await client.get("/employee/list/", {
      page: String(page),
      page_size: "100",
    });
    const results = Array.isArray(response) ? response : response.results || [];
    all.push(...results);
    hasMore = response.next != null;
    page++;
  }

  return all;
}

/**
 * Fetch termination dates from /onboarding/workflow-dashboard/ (paginated)
 * @param {OmniHRAPIClient} client
 * @returns {Promise<Object>} Map of user_id -> termination_date
 */
async function fetchTerminationDates(client) {
  const terminationDates = {};
  let page = 1;
  let hasMore = true;

  while (hasMore) {
    const response = await client.get("/onboarding/workflow-dashboard/", {
      page: String(page),
      page_size: "100",
    });
    const results = response.results || [];
    for (const emp of results) {
      if (emp.termination_date) {
        terminationDates[emp.id] = emp.termination_date;
      }
    }
    hasMore = response.next != null;
    page++;
  }

  return terminationDates;
}

/**
 * Build merged employee list with hired_date and termination_date
 * Excludes Omni Support, People Culture
 * @param {OmniHRAPIClient} client
 * @returns {Promise<Array>}
 */
async function fetchEmployeesWithHireAndTermination(client) {
  const [allEmployees, terminationDates] = await Promise.all([
    fetchAllEmployees(client),
    fetchTerminationDates(client),
  ]);

  const filtered = allEmployees.filter((emp) => {
    const fullName = (emp.full_name || emp.name || "").trim();
    const isExcluded = EXCLUDED_EMPLOYEES.some(
      (excluded) => fullName.toLowerCase() === excluded.toLowerCase(),
    );
    return !isExcluded;
  });

  return filtered.map((emp) => {
    const userId = emp.id || emp.user_id;
    return {
      user_id: userId,
      employee_id: emp.employee_id || "",
      full_name: emp.full_name || emp.name || `User ${userId}`,
      hired_date: emp.hired_date || null,
      termination_date: terminationDates[userId] || null,
    };
  });
}

/**
 * Detect hire and termination; return summary and lists
 * @param {Array} employees - from fetchEmployeesWithHireAndTermination
 * @param {Object} options - { recentDays: number for "recent" filter }
 * @returns {Object}
 */
function detectHireAndTermination(employees, options = {}) {
  const { recentDays = 30 } = options;
  const now = new Date();
  const recentStart = new Date(now);
  recentStart.setDate(recentStart.getDate() - recentDays);

  const withHireDate = [];
  const withTerminationDate = [];
  const recentHires = [];
  const recentTerminations = [];

  for (const emp of employees) {
    const hireDate = parseDateDMY(emp.hired_date);
    const termDate = parseDateDMY(emp.termination_date);

    if (hireDate) {
      withHireDate.push({ ...emp, hired_date_parsed: hireDate });
      if (hireDate >= recentStart && hireDate <= now) {
        recentHires.push({ ...emp, hired_date_parsed: hireDate });
      }
    }

    if (termDate) {
      withTerminationDate.push({ ...emp, termination_date_parsed: termDate });
      if (termDate >= recentStart && termDate <= now) {
        recentTerminations.push({ ...emp, termination_date_parsed: termDate });
      }
    }
  }

  return {
    totalEmployees: employees.length,
    withHireDate,
    withTerminationDate,
    recentHires,
    recentTerminations,
    recentDays,
  };
}

/**
 * Generate markdown report
 * @param {Object} result - from detectHireAndTermination
 * @param {Date} runAt
 * @returns {string}
 */
function toMarkdown(result, runAt) {
  const lines = [
    "# Hire & Termination Detection Report",
    "",
    `**Generated:** ${runAt.toISOString()}`,
    "",
    "## Summary",
    "",
    `| Metric | Count |`,
    `|--------|-------|`,
    `| Total employees (excl. Omni Support, People Culture) | ${result.totalEmployees} |`,
    `| Employees with hire date | ${result.withHireDate.length} |`,
    `| Employees with termination date | ${result.withTerminationDate.length} |`,
    `| Recent hires (last ${result.recentDays} days) | ${result.recentHires.length} |`,
    `| Recent terminations (last ${result.recentDays} days) | ${result.recentTerminations.length} |`,
    "",
    "---",
    "",
    "## Recent Hires (last " + result.recentDays + " days)",
    "",
  ];

  if (result.recentHires.length === 0) {
    lines.push("*None*", "");
  } else {
    lines.push(
      "| Employee ID | Full Name | Hired Date |",
      "|-------------|-----------|------------|",
    );
    for (const e of result.recentHires) {
      lines.push(
        `| ${e.employee_id || e.user_id} | ${e.full_name} | ${formatDate(e.hired_date_parsed)} |`,
      );
    }
    lines.push("");
  }

  lines.push(
    "---",
    "",
    "## Recent Terminations (last " + result.recentDays + " days)",
    "",
  );

  if (result.recentTerminations.length === 0) {
    lines.push("*None*", "");
  } else {
    lines.push(
      "| Employee ID | Full Name | Termination Date |",
      "|-------------|-----------|------------------|",
    );
    for (const e of result.recentTerminations) {
      lines.push(
        `| ${e.employee_id || e.user_id} | ${e.full_name} | ${formatDate(e.termination_date_parsed)} |`,
      );
    }
    lines.push("");
  }

  lines.push("---", "", "## All Employees with Hire Date", "");
  lines.push(
    "| Employee ID | Full Name | Hired Date |",
    "|-------------|-----------|------------|",
  );
  for (const e of result.withHireDate) {
    lines.push(
      `| ${e.employee_id || e.user_id} | ${e.full_name} | ${formatDate(e.hired_date_parsed)} |`,
    );
  }
  lines.push("");

  lines.push("---", "", "## All Employees with Termination Date", "");
  lines.push(
    "| Employee ID | Full Name | Termination Date |",
    "|-------------|-----------|------------------|",
  );
  for (const e of result.withTerminationDate) {
    lines.push(
      `| ${e.employee_id || e.user_id} | ${e.full_name} | ${formatDate(e.termination_date_parsed)} |`,
    );
  }
  lines.push("");

  return lines.join("\n");
}

async function main() {
  const outputDir = path.join(__dirname, "..");
  const outputPath = path.join(outputDir, "hire-termination-report.md");

  console.log("Detecting hire and termination employees...");
  const client = new OmniHRAPIClient();

  const employees = await fetchEmployeesWithHireAndTermination(client);
  const result = detectHireAndTermination(employees, { recentDays: 30 });
  const runAt = new Date();
  const md = toMarkdown(result, runAt);

  fs.mkdirSync(outputDir, { recursive: true });
  fs.writeFileSync(outputPath, md, "utf8");

  console.log("Summary:");
  console.log("  Total employees:", result.totalEmployees);
  console.log("  With hire date:", result.withHireDate.length);
  console.log("  With termination date:", result.withTerminationDate.length);
  console.log("  Recent hires (30 days):", result.recentHires.length);
  console.log(
    "  Recent terminations (30 days):",
    result.recentTerminations.length,
  );
  console.log("\nReport written to:", outputPath);

  return { result, outputPath };
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

module.exports = {
  parseDateDMY,
  fetchAllEmployees,
  fetchTerminationDates,
  fetchEmployeesWithHireAndTermination,
  detectHireAndTermination,
  toMarkdown,
};
