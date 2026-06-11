// Salary OA — business logic (browser port of salary_oa_core.py).
// All numeric rules are lifted from the original workbook's Power Query M code:
//   1. Govt Payment 3/4 (PF employee/employer) come from PF MIS "PF Ac. 1".
//   2. PF Admin Charge auto-derived from PF MIS "Admin Charges Ac. 2".
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

// ---------------------------------------------------------------------------
// Generic spreadsheet helpers
// ---------------------------------------------------------------------------

// Read the worksheet with the most data from an .xls/.xlsx/.xlsm buffer —
// handles workbooks where the data isn't on the first tab.
export function readSheet(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
  let best = null;
  let bestScore = -1;
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const ref = ws["!ref"];
    if (!ref) continue;
    const range = XLSX.utils.decode_range(ref);
    const score = (range.e.r + 1) * (range.e.c + 1);
    if (score > bestScore) {
      best = ws;
      bestScore = score;
    }
  }
  if (!best) throw new Error("Workbook has no data sheets.");
  return XLSX.utils.sheet_to_json(best, { header: 1, defval: null, raw: true });
}

function findHeaderRow(rows, mustContain) {
  const needles = mustContain.map((s) => s.trim().toLowerCase());
  for (let i = 0; i < rows.length; i++) {
    const cells = new Set(
      rows[i].filter((v) => v !== null && v !== undefined).map((v) => String(v).trim().toLowerCase())
    );
    if (needles.every((n) => cells.has(n))) return i;
  }
  throw new Error(`Could not find header row containing all of: ${mustContain.join(", ")}`);
}

function columnIndex(header, name) {
  const target = name.trim().toLowerCase();
  for (let i = 0; i < header.length; i++) {
    const v = header[i];
    if (v !== null && v !== undefined && String(v).trim().toLowerCase() === target) return i;
  }
  return null;
}

export function num(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const s = String(v).replace(/,/g, "").trim();
  if (!s) return 0;
  const f = parseFloat(s);
  return Number.isNaN(f) ? 0 : f;
}

export function cleanCode(v) {
  let s = String(v).trim();
  if (s.endsWith(".0")) s = s.slice(0, -2);
  return s;
}

function isTotalLabel(v) {
  if (v === null || v === undefined) return false;
  const s = String(v).trim().toLowerCase();
  return s === "total" || s === "grand total";
}

// Python's round() is banker's rounding (half to even) — match it so outputs
// tie exactly with the desktop tool.
export function pyRound(x) {
  const r = Math.round(x);
  if (Math.abs(x % 1) === 0.5) {
    const f = Math.floor(x);
    return f % 2 === 0 ? f : f + 1;
  }
  return r;
}

// ---------------------------------------------------------------------------
// Domain model (plain objects)
// ---------------------------------------------------------------------------

export function makeEmployee(fields) {
  return {
    code: "",
    name: "",
    department: "",
    totalEarnings: 0,
    incomeTax: 0,
    pfRegister: 0,
    professionTax: 0,
    netpayRegister: 0,
    pfMis: 0,
    inPfMis: false,
    loanDeduction: 0,
    bonus: 0,
    gratuity: 0,
    workedDays: 0,
    nComponents: 0,
    arrearsTotal: 0,
    recoveriesTotal: 0,
    leaveEncashment: 0,
    pfArrears: 0,
    ...fields,
  };
}

export function netPaid(e) {
  if (e.netpayRegister) return e.netpayRegister;
  return e.totalEarnings - e.incomeTax - e.pfRegister - e.professionTax;
}

export function totalOaCost(e) {
  return netPaid(e) + e.incomeTax + e.professionTax + 2 * e.pfMis;
}

export function makeExtraRow(fields) {
  return { code: "", name: "", department: "", netPaid: 0, tds: 0, pt: 0, pfEmployee: 0, pfEmployer: 0, ...fields };
}

export function extraTotal(x) {
  return x.netPaid + x.tds + x.pt + x.pfEmployee + x.pfEmployer;
}

// ---------------------------------------------------------------------------
// Run helpers
// ---------------------------------------------------------------------------

export const DEFAULT_SERVICE_DEPTS = ["service", "service department"];
export const DEFAULT_DEPT_RENAMES = { "human resources": "OPERATIONS" };
export const DEFAULT_TECH_EMPLOYEES = ["aditya kumar"];

export function isService(run, dept) {
  return run.serviceDepts.includes((dept || "").trim().toLowerCase());
}

export function loanExtras(run) {
  return run.employees
    .filter((e) => e.loanDeduction)
    .map((e) =>
      makeExtraRow({ code: "LOANRECOVERY", name: e.name, department: e.department, netPaid: -e.loanDeduction })
    );
}

export function splitEmployees(run) {
  const svc = [];
  const nonsvc = [];
  for (const e of run.employees) (isService(run, e.department) ? svc : nonsvc).push(e);
  return [svc, nonsvc];
}

export function splitExtras(run) {
  const svc = [];
  const nonsvc = [];
  for (const x of run.extras) (isService(run, x.department) ? svc : nonsvc).push(x);
  return [svc, nonsvc];
}

export function otherTotal(run) {
  return run.otherCosts.reduce((s, o) => s + o.amount, 0);
}

// ---------------------------------------------------------------------------
// Column classification for the APAC component breakdown
// ---------------------------------------------------------------------------

const N_FOLD_HEADERS = [
  "basic",
  "house rent allowance",
  "special allowance",
  "books & periodicals reimbursement",
  "leave travel allowance",
  "professional development allowance",
  "fuel expenses reimbursement",
];
const N_EXCLUDE = ["arrears", "recovery"];
const RESERVED = [
  "total earnings", "income tax", "provident fund", "profession tax",
  "netpay", "total deductions", "bonus", "gratuity on resignation",
  "emp worked days", "code", "name", "department", "remarks",
  "sl no", "bank ac no", "ifsc code", "bank name", "email",
  "father name", "branch", "pan", "lop days", "rlop",
];

export function classifyHeader(h) {
  if (h === null || h === undefined) return null;
  const s = String(h).trim().toLowerCase().replaceAll("  ", " ");
  if (!s) return null;
  if (RESERVED.some((r) => s === r || s.startsWith(r))) return null;
  if (s.includes("leave encashment")) return "T";
  if (s.includes("provident fund")) return null; // PF Arrears/Recovery handled separately
  if (s.includes("recovery")) return "R-";
  if (s.includes("arrears")) return "R+";
  if (N_FOLD_HEADERS.some((n) => s.includes(n)) && !N_EXCLUDE.some((x) => s.includes(x))) return "N";
  return null;
}

// ---------------------------------------------------------------------------
// Loaders
// ---------------------------------------------------------------------------

export function loadSalary(arrayBuffer, opts = {}) {
  const rows = readSheet(arrayBuffer);
  const headerIdx = findHeaderRow(rows, ["Code", "Name", "Total Earnings"]);
  const header = rows[headerIdx];

  const cols = {
    code: columnIndex(header, "Code"),
    name: columnIndex(header, "Name"),
    department: columnIndex(header, "Department"),
    totalEarnings: columnIndex(header, "Total Earnings"),
    incomeTax: columnIndex(header, "Income Tax"),
    providentFund: columnIndex(header, "Provident Fund"),
    pfArrears: columnIndex(header, "Provident Fund Arrears"),
    professionTax: columnIndex(header, "Profession Tax"),
    netpay: columnIndex(header, "NETPAY"),
    bonus: columnIndex(header, "Bonus"),
    gratuity: columnIndex(header, "GRATUITY ON RESIGNATION"),
    workedDays: columnIndex(header, "Emp Worked days"),
  };
  const missing = ["code", "name", "department", "totalEarnings", "incomeTax", "providentFund", "professionTax"]
    .filter((k) => cols[k] === null);
  if (missing.length) throw new Error(`Salary register is missing required columns: ${missing.join(", ")}`);

  const classified = [];
  header.forEach((h, ci) => {
    const kind = classifyHeader(h);
    if (kind) classified.push([ci, kind]);
  });

  const renames = {};
  for (const [k, v] of Object.entries(opts.deptRenames || DEFAULT_DEPT_RENAMES)) renames[k.toLowerCase()] = v;
  const techLc = (opts.techEmployees || DEFAULT_TECH_EMPLOYEES).map((n) => n.toLowerCase());
  const zeroTds = new Set(opts.zeroTdsCodes || []);

  const out = [];
  for (const row of rows.slice(headerIdx + 1)) {
    if (!row.some((v) => v !== null && v !== undefined && v !== "" && v !== 0) && !row.some(Boolean)) continue;
    const codeV = cols.code < row.length ? row[cols.code] : null;
    if (codeV === null || codeV === undefined || String(codeV).trim() === "") continue;
    if (isTotalLabel(codeV)) continue;
    const code = cleanCode(codeV);
    const name = String(row[cols.name] ?? "").trim();
    let dept = String(row[cols.department] ?? "").trim();
    if (renames[dept.toLowerCase()]) dept = renames[dept.toLowerCase()];
    if (techLc.includes(name.toLowerCase())) dept = "TECHNICAL";

    let nComponents = 0, arrearsTotal = 0, recoveriesTotal = 0, leaveEncashment = 0;
    for (const [ci, kind] of classified) {
      if (ci >= row.length) continue;
      const v = num(row[ci]);
      if (kind === "N") nComponents += v;
      else if (kind === "R+") arrearsTotal += v;
      else if (kind === "R-") recoveriesTotal += v;
      else if (kind === "T") leaveEncashment += v;
    }

    const totalEarnings = num(row[cols.totalEarnings]) - recoveriesTotal;
    let incomeTax = num(row[cols.incomeTax]);
    if (zeroTds.has(code)) incomeTax = 0;

    const pfRegular = num(row[cols.providentFund]);
    const pfArrears = cols.pfArrears !== null ? num(row[cols.pfArrears]) : 0;

    out.push(
      makeEmployee({
        code,
        name,
        department: dept,
        totalEarnings,
        incomeTax,
        pfRegister: pfRegular + pfArrears,
        professionTax: num(row[cols.professionTax]),
        netpayRegister: cols.netpay !== null ? num(row[cols.netpay]) : 0,
        bonus: cols.bonus !== null ? num(row[cols.bonus]) : 0,
        gratuity: cols.gratuity !== null ? num(row[cols.gratuity]) : 0,
        workedDays: cols.workedDays !== null ? num(row[cols.workedDays]) : 0,
        nComponents,
        arrearsTotal,
        recoveriesTotal,
        leaveEncashment,
        pfArrears,
      })
    );
  }
  return out;
}

export function loadPf(arrayBuffer) {
  const rows = readSheet(arrayBuffer);
  const headerIdx = findHeaderRow(rows, ["Code", "PF Ac. 1"]);
  const header = rows[headerIdx];
  const codeCol = columnIndex(header, "Code");
  const pfCol = columnIndex(header, "PF Ac. 1");
  const adminCol = columnIndex(header, "Admin Charges Ac. 2");
  const edliCol = columnIndex(header, "EDLI Charges Ac. 21");
  if (codeCol === null || pfCol === null) throw new Error("PF MIS missing 'Code' or 'PF Ac. 1' column");

  const pfLookup = {};
  let adminTotal = 0;
  let edliTotal = 0;
  for (const row of rows.slice(headerIdx + 1)) {
    if (!row.some((v) => v !== null && v !== undefined && v !== "")) continue;
    const codeV = codeCol < row.length ? row[codeCol] : null;
    if (!codeV) continue;
    const nextCell = codeCol + 1 < row.length ? row[codeCol + 1] : null;
    if (isTotalLabel(codeV) || isTotalLabel(nextCell)) continue;
    const code = cleanCode(codeV);
    pfLookup[code] = num(row[pfCol]);
    if (adminCol !== null && adminCol < row.length) adminTotal += num(row[adminCol]);
    if (edliCol !== null && edliCol < row.length) edliTotal += num(row[edliCol]);
  }
  return { pfLookup, adminTotal, edliTotal };
}

export function attachPf(employees, pfLookup) {
  for (const e of employees) {
    if (Object.prototype.hasOwnProperty.call(pfLookup, e.code)) {
      e.pfMis = pfLookup[e.code];
      e.inPfMis = true;
    }
  }
}

// ---------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------

function makeCheck(name, sourceValue, outputValue, tolerance = 1.0) {
  const diff = outputValue - sourceValue;
  return { name, sourceValue, outputValue, diff, ok: Math.abs(diff) <= tolerance };
}

export function validate(run, pfLookup, pfAdminSource, extraPfEmployees = {}) {
  const allEmp = run.employees;
  const loanRows = loanExtras(run);
  const userExtras = [...run.extras];
  const allExtra = [...userExtras, ...loanRows];

  const totalLoanDeductions = allEmp.reduce((s, e) => s + e.loanDeduction, 0);
  const sumRegisterNetpay = allEmp.reduce((s, e) => s + e.netpayRegister, 0) - totalLoanDeductions;
  const sumRegisterTds = allEmp.reduce((s, e) => s + e.incomeTax, 0);
  const sumRegisterPt = allEmp.reduce((s, e) => s + e.professionTax, 0);
  const sumRegisterPf = allEmp.reduce((s, e) => s + e.pfRegister, 0);
  const sumMisPf = Object.values(pfLookup).reduce((s, v) => s + v, 0);

  const ot = otherTotal(run);
  const outNetpay = allEmp.reduce((s, e) => s + netPaid(e), 0) + allExtra.reduce((s, x) => s + x.netPaid, 0);
  const outTds = sumRegisterTds + allExtra.reduce((s, x) => s + x.tds, 0);
  const outPt = sumRegisterPt + allExtra.reduce((s, x) => s + x.pt, 0);
  const outPfEmployee = allEmp.reduce((s, e) => s + e.pfMis, 0) + allExtra.reduce((s, x) => s + x.pfEmployee, 0);
  const outPfEmployer = allEmp.reduce((s, e) => s + e.pfMis, 0) + allExtra.reduce((s, x) => s + x.pfEmployer, 0) + ot;
  const outTotal =
    allEmp.reduce((s, e) => s + totalOaCost(e), 0) + allExtra.reduce((s, x) => s + extraTotal(x), 0) + ot;

  const identity = outNetpay + outTds + outPt + outPfEmployee + outPfEmployer;

  const sumUx = (f) => userExtras.reduce((s, x) => s + f(x), 0);

  const checks = [
    makeCheck("Net Pay  (register NETPAY − loan deductions → output Net Paid)", sumRegisterNetpay,
      outNetpay - sumUx((x) => x.netPaid)),
    makeCheck("TDS  (register Income Tax → output Govt 1)", sumRegisterTds, outTds - sumUx((x) => x.tds)),
    makeCheck("PT  (register Profession Tax → output Govt 2)", sumRegisterPt, outPt - sumUx((x) => x.pt)),
    makeCheck("PF Employee  (PF MIS PF Ac. 1 → output Govt 3)", sumMisPf, outPfEmployee - sumUx((x) => x.pfEmployee)),
    makeCheck("PF Employer  (PF MIS PF Ac. 1 → output Govt 4)", sumMisPf,
      outPfEmployer - ot - sumUx((x) => x.pfEmployer)),
    makeCheck("Other Costs  (PF MIS Admin + EDLI → 'Other Cost' rows)", pfAdminSource, ot, 1.0),
    makeCheck("Headcount  (salary register vs Final Combined employee rows)", allEmp.length, allEmp.length, 0),
    makeCheck("Identity  (Total OA Cost = Net Paid + 4 Govt Payments)", identity, outTotal, 1.0),
    makeCheck("Register PF vs PF MIS PF Ac. 1 (sanity — should match after manual reconciliation)",
      sumRegisterPf, sumMisPf),
  ];

  const expectedExtras = new Set(Object.keys(extraPfEmployees).map((k) => k.toLowerCase()));
  const registerCodes = new Set(allEmp.map((e) => e.code));
  const missingInPfMis = allEmp.filter((e) => !e.inPfMis && e.pfRegister !== 0);
  const extraInPfMis = Object.keys(pfLookup).filter(
    (c) => !registerCodes.has(c) && !expectedExtras.has(c.toLowerCase())
  );
  const pfAmountMismatches = allEmp
    .filter((e) => e.inPfMis && Math.abs(e.pfRegister - pfLookup[e.code]) > 1.0)
    .map((e) => [e, pfLookup[e.code]]);

  return { checks, missingInPfMis, extraInPfMis, pfAmountMismatches, allOk: checks.every((c) => c.ok) };
}

// ---------------------------------------------------------------------------
// End-to-end build
// ---------------------------------------------------------------------------

export function build(salaryBuffer, pfBuffer, opts = {}) {
  const employees = loadSalary(salaryBuffer, opts);
  const { pfLookup, adminTotal, edliTotal } = loadPf(pfBuffer);
  attachPf(employees, pfLookup);
  const pfAdmin = opts.pfAdminOverride ?? adminTotal;

  const otherCosts = [];
  if (pfAdmin > 0)
    otherCosts.push({ label: "Other Cost", subLabel: "PF Admin Charge", department: "Operations Department", amount: pfAdmin });
  if (edliTotal > 0)
    otherCosts.push({ label: "Other Cost", subLabel: "EDLI Charges", department: "Operations Department", amount: edliTotal });

  const run = {
    employees,
    extras: [...(opts.extras || [])],
    otherCosts,
    monthLabel: opts.monthLabel || "",
    serviceDepts: opts.serviceDepts?.length ? opts.serviceDepts : DEFAULT_SERVICE_DEPTS,
  };
  const report = validate(run, pfLookup, adminTotal + edliTotal, opts.extraPfEmployees || {});
  return { run, report, pfLookup, pfAdminSource: adminTotal + edliTotal, adminTotal, edliTotal };
}

// ---------------------------------------------------------------------------
// Workbook writers (ExcelJS)
// ---------------------------------------------------------------------------

export const OA_HEADERS = [
  "Code", "Name", "Department", "Net Paid to Employee",
  "Government Payment 1 - TDS", "Government Payment 2 - PT",
  "Government Payment 3 - PF", "Government Payment 4 - PF Employer",
  "Total OA Cost",
];
const COMBINED_HEADERS = [
  "Code", "Name", "Department", "Total Earnings", "Income Tax",
  "Provident Fund (Register)", "PF (MIS — Ac. 1)", "Profession Tax", "NETPAY (Register)",
];

const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF305496" } };
const HEADER_FONT = { bold: true, color: { argb: "FFFFFFFF" } };
const TOTAL_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9E1F2" } };
const OK_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC6EFCE" } };
const FAIL_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
export const INT_FMT = "#,##0;[Red]-#,##0;-";

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function empRow(e) {
  return [e.code, e.name, e.department, pyRound(netPaid(e)), pyRound(e.incomeTax), pyRound(e.professionTax),
    pyRound(e.pfMis), pyRound(e.pfMis), pyRound(totalOaCost(e))];
}
function extraRowVals(x) {
  return [x.code, x.name, x.department, pyRound(x.netPaid), pyRound(x.tds), pyRound(x.pt),
    pyRound(x.pfEmployee), pyRound(x.pfEmployer), pyRound(extraTotal(x))];
}
function otherRowVals(o) {
  return [o.label, o.subLabel, o.department, 0, 0, 0, 0, pyRound(o.amount), pyRound(o.amount)];
}

function writeOaSheet(ws, employeeRows, extraRows, otherCosts, useFormulas = false) {
  ws.addRow(OA_HEADERS);
  for (let c = 1; c <= OA_HEADERS.length; c++) {
    const cell = ws.getCell(1, c);
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.alignment = { horizontal: "center" };
  }
  const allRows = [...employeeRows, ...extraRows, ...otherCosts.map(otherRowVals)];
  for (const r of allRows) ws.addRow(r);

  const last = ws.rowCount;
  let totals;
  if (useFormulas) {
    totals = ["Total", "", ""];
    for (let c = 4; c <= OA_HEADERS.length; c++) {
      totals.push({ formula: `SUBTOTAL(109,${colLetter(c)}2:${colLetter(c)}${last})` });
    }
  } else {
    const colTotals = new Array(OA_HEADERS.length).fill(0);
    for (const r of allRows) {
      for (let i = 3; i < OA_HEADERS.length; i++) {
        if (typeof r[i] === "number") colTotals[i] += r[i];
      }
    }
    totals = ["Total", "", "", ...colTotals.slice(3).map(pyRound)];
  }
  ws.addRow(totals);
  const totalRow = ws.rowCount;
  for (let c = 1; c <= OA_HEADERS.length; c++) {
    ws.getCell(totalRow, c).fill = TOTAL_FILL;
    ws.getCell(totalRow, c).font = { bold: true };
  }
  for (let c = 4; c <= OA_HEADERS.length; c++) {
    for (let r = 2; r <= totalRow; r++) ws.getCell(r, c).numFmt = INT_FMT;
  }
  [12, 32, 22, 16, 16, 14, 16, 18, 16].forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });
  ws.views = [{ state: "frozen", ySplit: 1 }];
}

function splitForWrite(run) {
  const [svc, nonsvc] = splitEmployees(run);
  const [svcX, nonsvcX] = splitExtras(run);
  const combinedSorted = [...run.employees].sort((a, b) => b.incomeTax - a.incomeTax);
  const loanRows = loanExtras(run);
  for (const lx of loanRows) (isService(run, lx.department) ? svcX : nonsvcX).push(lx);
  const allExtras = [...run.extras, ...loanRows];
  return { svc, nonsvc, svcX, nonsvcX, combinedSorted, allExtras };
}

export async function writeWorkbook(run, report) {
  const { svc, nonsvc, svcX, nonsvcX, combinedSorted, allExtras } = splitForWrite(run);
  const wb = new ExcelJS.Workbook();

  // ---- Validation tab ----
  const val = wb.addWorksheet("Validation");
  val.addRow([`Validation Report${run.monthLabel ? " — " + run.monthLabel : ""}`]);
  val.getCell(1, 1).font = { bold: true, size: 14 };
  val.addRow([]);
  val.addRow(["Check", "Source value", "Output value", "Diff", "Status"]);
  for (let c = 1; c <= 5; c++) {
    val.getCell(3, c).fill = HEADER_FILL;
    val.getCell(3, c).font = HEADER_FONT;
  }
  for (const ck of report.checks) {
    val.addRow([ck.name, pyRound(ck.sourceValue), pyRound(ck.outputValue), pyRound(ck.diff), ck.ok ? "OK" : "FAIL"]);
    const r = val.rowCount;
    val.getCell(r, 5).fill = ck.ok ? OK_FILL : FAIL_FILL;
    val.getCell(r, 5).font = { bold: true };
    for (const c of [2, 3, 4]) val.getCell(r, c).numFmt = INT_FMT;
  }
  if (report.missingInPfMis.length) {
    val.addRow([]);
    val.addRow([`Employees in salary register missing from PF MIS (${report.missingInPfMis.length}):`]);
    val.getCell(val.rowCount, 1).font = { bold: true };
    for (const e of report.missingInPfMis) val.addRow([e.code, e.name, e.department, pyRound(e.pfRegister)]);
  }
  if (report.extraInPfMis.length) {
    val.addRow([]);
    val.addRow([`Codes in PF MIS missing from salary register (${report.extraInPfMis.length}):`]);
    val.getCell(val.rowCount, 1).font = { bold: true };
    for (const c of report.extraInPfMis) val.addRow([c]);
  }
  if (report.pfAmountMismatches.length) {
    val.addRow([]);
    val.addRow([`PF amount differs (register vs MIS) — output uses MIS (${report.pfAmountMismatches.length}):`]);
    val.getCell(val.rowCount, 1).font = { bold: true };
    val.addRow(["Code", "Name", "Register PF", "MIS PF", "Diff"]);
    for (const [e, mis] of report.pfAmountMismatches) {
      val.addRow([e.code, e.name, pyRound(e.pfRegister), pyRound(mis), pyRound(mis - e.pfRegister)]);
    }
  }
  [60, 18, 18, 14, 10].forEach((w, i) => {
    val.getColumn(i + 1).width = w;
  });

  // ---- Summary ----
  const summary = wb.addWorksheet("Summary");
  summary.addRow([`Salary OA Summary${run.monthLabel ? " — " + run.monthLabel : ""}`]);
  summary.getCell(1, 1).font = { bold: true, size: 14 };
  summary.addRow([]);
  summary.addRow(["Sum of Total OA Cost"]);
  summary.addRow(["Department Group", "Total"]);
  const svcTotal = svc.reduce((s, e) => s + totalOaCost(e), 0) + svcX.reduce((s, x) => s + extraTotal(x), 0);
  const nsvcTotal =
    nonsvc.reduce((s, e) => s + totalOaCost(e), 0) + nonsvcX.reduce((s, x) => s + extraTotal(x), 0) + otherTotal(run);
  summary.addRow(["Service", pyRound(svcTotal)]);
  summary.addRow(["Non service", pyRound(nsvcTotal)]);
  summary.addRow(["Grand Total", pyRound(svcTotal + nsvcTotal)]);
  for (let r = 4; r <= 7; r++) summary.getCell(r, 2).numFmt = INT_FMT;
  summary.getRow(4).font = { bold: true };
  summary.getRow(7).font = { bold: true };
  summary.getColumn(1).width = 26;
  summary.getColumn(2).width = 18;

  // ---- OA tabs ----
  writeOaSheet(wb.addWorksheet("Final Combined (Working)"), combinedSorted.map(empRow), allExtras.map(extraRowVals),
    run.otherCosts, true);
  writeOaSheet(wb.addWorksheet("Service Cost"), svc.map(empRow), svcX.map(extraRowVals), []);
  writeOaSheet(wb.addWorksheet("Non Service Cost"), nonsvc.map(empRow), nonsvcX.map(extraRowVals), run.otherCosts);
  writeOaSheet(wb.addWorksheet("Final Combined"), combinedSorted.map(empRow), allExtras.map(extraRowVals), run.otherCosts);

  // ---- COMBINED (raw extract) ----
  const comb = wb.addWorksheet("COMBINED");
  comb.addRow(COMBINED_HEADERS);
  for (let c = 1; c <= COMBINED_HEADERS.length; c++) {
    comb.getCell(1, c).fill = HEADER_FILL;
    comb.getCell(1, c).font = HEADER_FONT;
  }
  for (const e of run.employees) {
    comb.addRow([e.code, e.name, e.department, pyRound(e.totalEarnings), pyRound(e.incomeTax),
      pyRound(e.pfRegister), pyRound(e.pfMis), pyRound(e.professionTax), pyRound(e.netpayRegister)]);
  }
  for (let c = 4; c <= COMBINED_HEADERS.length; c++) {
    for (let r = 2; r <= comb.rowCount; r++) comb.getCell(r, c).numFmt = INT_FMT;
  }
  [12, 32, 22, 16, 14, 16, 16, 14, 16].forEach((w, i) => {
    comb.getColumn(i + 1).width = w;
  });
  comb.views = [{ state: "frozen", ySplit: 1 }];

  return wb.xlsx.writeBuffer();
}

// 3 standalone single-sheet workbooks (Service / Non Service / Final Combined)
export async function splitExports(run) {
  const { svc, nonsvc, svcX, nonsvcX, combinedSorted, allExtras } = splitForWrite(run);
  const specs = [
    ["Service Cost", svc.map(empRow), svcX.map(extraRowVals), []],
    ["Non Service Cost", nonsvc.map(empRow), nonsvcX.map(extraRowVals), run.otherCosts],
    ["Final Combined", combinedSorted.map(empRow), allExtras.map(extraRowVals), run.otherCosts],
  ];
  const out = {};
  for (const [name, empRows, exRows, oc] of specs) {
    const wb = new ExcelJS.Workbook();
    writeOaSheet(wb.addWorksheet(name), empRows, exRows, oc, false);
    out[name] = await wb.xlsx.writeBuffer();
  }
  return out;
}
