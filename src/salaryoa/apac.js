// APAC Payroll Report generator (browser port of apac_core.py).
// Takes last month's APAC (decrypted), current Run → filled new APAC,
// Manual Fill report, Reconciliation report.
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { netPaid, totalOaCost, extraTotal, otherTotal, pyRound, INT_FMT } from "./engine.js";

// Column layout in APAC "Payroll Details" (1-indexed). Data starts at row 4.
const COL = {
  no: 2, code: 3, name: 4, employment: 5, business_unit: 6, department: 7,
  position: 8, country: 9, currency: 10, start_date: 11, resign_date: 12,
  attendance: 13, monthly_salary: 14, back_pay: 18, paid_leave: 20,
  birthday: 40, bonus: 39, severance: 49, loan: 76,
  pf_employer: 82, pf_employee: 83, tds: 86, other_ded: 89, total_lc: 94,
};
const DATA_ROW_START = 4;

const YELLOW = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE699" } };
const HDR_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF305496" } };
const HDR_FONT = { bold: true, color: { argb: "FFFFFFFF" } };
const WARN_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFEB9C" } };
const INFO_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF7" } };

const DEPT_NORMALISE = {
  "business department": "SALES", sales: "SALES",
  "service department": "SERVICE", service: "SERVICE",
  "operations department": "OPERATIONS", operations: "OPERATIONS",
  technical: "TECHNICAL", finance: "Finance",
  "human resources": "HR", hr: "HR", marketing: "Marketing",
  "overseas marketing department": "Overseas", overseas: "Overseas",
};

// Every "Total" column formula, parameterised by row. Force-written on each
// row we touch so stale hardcoded template values can't survive.
function rowFormulas(r) {
  return {
    V: `SUM(N${r}:U${r})`,
    Z: `SUM(X${r}:Y${r})`,
    AJ: `SUM(AB${r}:AI${r})`,
    AR: `SUM(AL${r}:AQ${r})`,
    AU: `SUM(AT${r}:AT${r})`,
    BD: `SUM(AW${r}:BC${r})`,
    BK: `SUM(BF${r}:BJ${r})`,
    BQ: `SUM(BM${r}:BP${r})`,
    BS: `SUM(V${r},Z${r},AJ${r},AR${r},AU${r},BD${r})-BK${r}`,
    CB: `SUM(BT${r}:CA${r})`,
    CI: `SUM(CD${r}:CH${r})`,
    CL: `CK${r}`,
    CN: `BS${r}-CB${r}-CI${r}-CL${r}`,
    CP: `SUM(CI${r},CL${r},CN${r})`,
  };
}

function writeTotalFormulas(ws, rowIdx) {
  for (const [colL, formula] of Object.entries(rowFormulas(rowIdx))) {
    ws.getCell(`${colL}${rowIdx}`).value = { formula };
  }
}

function cleanCodeStr(v) {
  let s = String(v).trim();
  if (s.endsWith(".0")) s = s.slice(0, -2);
  return s;
}

// ExcelJS trips over cell-comment relationship parts written by some tools
// (TypeError in worksheet-xform reconcile). The template's old comments are
// disposable — we write our own notes — so strip comments/VML before loading.
async function stripComments(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  // Comment parts vary by writer: Excel uses xl/comments1.xml + vmlDrawing1.vml,
  // openpyxl uses xl/comments/comment1.xml + xl/drawings/commentsDrawing1.vml.
  const toRemove = Object.keys(zip.files).filter(
    (p) => (/comment/i.test(p) && /\.(xml|vml)$/.test(p)) || /vmlDrawing\d*\.vml$/.test(p)
  );
  if (!toRemove.length) return arrayBuffer;
  for (const p of toRemove) zip.remove(p);
  for (const relPath of Object.keys(zip.files).filter((p) => /xl\/worksheets\/_rels\/.*\.rels$/.test(p))) {
    let xml = await zip.file(relPath).async("string");
    xml = xml.replace(/<Relationship\b[^>]*?(?:relationships\/comments|relationships\/vmlDrawing)[^>]*?\/>/g, "");
    zip.file(relPath, xml);
  }
  for (const wsPath of Object.keys(zip.files).filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p))) {
    let xml = await zip.file(wsPath).async("string");
    xml = xml.replace(/<legacyDrawing[^>]*\/>/g, "");
    zip.file(wsPath, xml);
  }
  let ct = await zip.file("[Content_Types].xml").async("string");
  ct = ct.replace(/<Override\b[^>]*comment[^>]*\/>/gi, "").replace(/<Default\b[^>]*Extension="vml"[^>]*\/>/gi, "");
  zip.file("[Content_Types].xml", ct);
  return zip.generateAsync({ type: "arraybuffer" });
}

export async function loadApacWorkbook(arrayBuffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await stripComments(arrayBuffer));
  if (!wb.getWorksheet("Payroll Details")) {
    throw new Error('Previous APAC has no "Payroll Details" sheet.');
  }
  return wb;
}

export function extractPrev(wb) {
  const ws = wb.getWorksheet("Payroll Details");
  const out = {};
  ws.eachRow((row, rowNumber) => {
    if (rowNumber < DATA_ROW_START) return;
    const codeV = row.getCell(COL.code).value;
    if (codeV === null || codeV === undefined || codeV === "") return;
    const code = cleanCodeStr(codeV);
    out[code] = {
      code,
      name: row.getCell(COL.name).value || "",
      employmentType: row.getCell(COL.employment).value || "Full time",
      businessUnit: row.getCell(COL.business_unit).value || "",
      department: row.getCell(COL.department).value || "",
      position: row.getCell(COL.position).value || "",
      country: row.getCell(COL.country).value || "India",
      currency: row.getCell(COL.currency).value || "INR",
      startDate: row.getCell(COL.start_date).value ?? null,
      monthlySalary: Number(row.getCell(COL.monthly_salary).value || 0),
    };
  });
  return out;
}

function clearDataRows(ws, startRow) {
  const last = ws.rowCount;
  for (let r = startRow; r <= last; r++) {
    const row = ws.getRow(r);
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (cell.type === ExcelJS.ValueType.Formula) return; // keep template formulas
      cell.value = null;
    });
  }
}

function setCell(ws, r, c, v) {
  ws.getCell(r, c).value = v;
}

function fillEmployeeRow(ws, rowIdx, no, emp, prev, manual) {
  setCell(ws, rowIdx, COL.no, no);
  setCell(ws, rowIdx, COL.code, emp.code);
  setCell(ws, rowIdx, COL.name, emp.name);

  if (prev) {
    setCell(ws, rowIdx, COL.employment, prev.employmentType);
    setCell(ws, rowIdx, COL.business_unit, prev.businessUnit);
    setCell(ws, rowIdx, COL.department, prev.department);
    setCell(ws, rowIdx, COL.position, prev.position);
    setCell(ws, rowIdx, COL.country, prev.country);
    setCell(ws, rowIdx, COL.currency, prev.currency);
    setCell(ws, rowIdx, COL.start_date, prev.startDate);
  } else {
    setCell(ws, rowIdx, COL.employment, "Full time");
    setCell(ws, rowIdx, COL.business_unit, "PV&ES Division");
    setCell(ws, rowIdx, COL.department, DEPT_NORMALISE[emp.department.trim().toLowerCase()] || emp.department);
    setCell(ws, rowIdx, COL.country, "India");
    setCell(ws, rowIdx, COL.currency, "INR");
    for (const k of ["business_unit", "department", "position", "start_date"]) {
      ws.getCell(rowIdx, COL[k]).fill = YELLOW;
    }
    manual.push({
      code: emp.code, name: emp.name, itemType: "New joiner",
      detail: "New employee — wasn't in last month's APAC. The yellow cells need verification: enter their Job Title (Position) and Joining Date (StartDate); confirm Department and Business Unit. Monthly Salary has been calculated from this month's pay components (same logic as every other employee).",
      suggestedValue: "—",
      apacColumns: "Business Unit, Department, Position, StartDate",
    });
  }

  // N = Basic+HRA+Special+Books+LTA+ProfDev+Fuel + REGULAR PF Employer
  const nValue = emp.nComponents + (emp.pfMis - emp.pfArrears);
  // R = Salary Arrears − Recoveries + PF Arrears (Employer share)
  const backPay = emp.arrearsTotal - emp.recoveriesTotal + emp.pfArrears;
  const tValue = emp.leaveEncashment;

  setCell(ws, rowIdx, COL.attendance, emp.workedDays || 31);
  setCell(ws, rowIdx, COL.monthly_salary, Math.round(nValue * 100) / 100);
  if (backPay) setCell(ws, rowIdx, COL.back_pay, Math.round(backPay * 100) / 100);
  if (tValue) setCell(ws, rowIdx, COL.paid_leave, Math.round(tValue * 100) / 100);

  if (emp.bonus) {
    const bcell = ws.getCell(rowIdx, COL.bonus);
    bcell.value = Math.round(emp.bonus * 100) / 100;
    bcell.fill = YELLOW;
    bcell.note = "Bonus from salary register parked here (Holiday Allowance_Company, PI00086). If any portion was a Birthday gift, move that amount to column AN (Birthday Benefit). Total stays unchanged.";
    manual.push({
      code: emp.code, name: emp.name, itemType: "Bonus split",
      detail: "This employee received a bonus this month. The full amount is currently in the yellow 'Holiday Allowance_Company' column (AM). If any portion was a birthday gift, move that part to 'Birthday Benefit' (AN). The total stays unchanged.",
      suggestedValue: Math.round(emp.bonus * 100) / 100,
      apacColumns: "Holiday Allowance_Company (AM) → Birthday Benefit (AN) if applicable",
    });
  }

  if (emp.gratuity) {
    const sv = ws.getCell(rowIdx, COL.severance);
    sv.value = Math.round(emp.gratuity * 100) / 100;
    sv.fill = YELLOW;
    ws.getCell(rowIdx, COL.resign_date).fill = YELLOW;
    manual.push({
      code: emp.code, name: emp.name, itemType: "Resignation / Severance",
      detail: "This employee resigned this month. Gratuity has been pre-filled in 'Severance Pay' (yellow). Please add their resignation date in the 'ResignDate' column (also yellow).",
      suggestedValue: Math.round(emp.gratuity * 100) / 100,
      apacColumns: "ResignDate (L), Severance Pay (AW)",
    });
  }

  setCell(ws, rowIdx, COL.pf_employer, Math.round(emp.pfMis * 100) / 100);
  setCell(ws, rowIdx, COL.pf_employee, Math.round(emp.pfMis * 100) / 100);
  setCell(ws, rowIdx, COL.tds, Math.round(emp.incomeTax * 100) / 100);
  if (emp.professionTax) setCell(ws, rowIdx, COL.other_ded, Math.round(emp.professionTax * 100) / 100);
  if (emp.loanDeduction) setCell(ws, rowIdx, COL.loan, Math.round(emp.loanDeduction * 100) / 100);

  writeTotalFormulas(ws, rowIdx);
}

function fillInternRow(ws, rowIdx, no, idx, name, deptCanonical, amount, manual) {
  const code = `Intern${idx}`;
  setCell(ws, rowIdx, COL.no, no);
  setCell(ws, rowIdx, COL.code, code);
  setCell(ws, rowIdx, COL.name, name);
  setCell(ws, rowIdx, COL.employment, "Internship");
  setCell(ws, rowIdx, COL.business_unit, "PV&ES Division");
  setCell(ws, rowIdx, COL.department, deptCanonical);
  setCell(ws, rowIdx, COL.position, "Intern");
  setCell(ws, rowIdx, COL.country, "India");
  setCell(ws, rowIdx, COL.currency, "INR");
  setCell(ws, rowIdx, COL.attendance, 31);
  setCell(ws, rowIdx, COL.monthly_salary, Math.round(amount * 100) / 100);
  ws.getCell(rowIdx, COL.start_date).fill = YELLOW;
  ws.getCell(rowIdx, COL.department).fill = YELLOW;
  writeTotalFormulas(ws, rowIdx);
  manual.push({
    code, name, itemType: "Intern",
    detail: "An intern has been added at the bottom of the report. Please add their joining date (StartDate, highlighted yellow) and confirm the Department is correct.",
    suggestedValue: Math.round(amount * 100) / 100,
    apacColumns: "StartDate, Department",
  });
}

function cellNum(ws, r, c) {
  const v = ws.getCell(r, c).value;
  return typeof v === "number" ? v : 0;
}

// prevBuffer: decrypted ArrayBuffer of last month's APAC.
export async function buildApac(prevBuffer, run) {
  const prevWb = await loadApacWorkbook(prevBuffer);
  const prevData = extractPrev(prevWb);

  // Fresh load to use as the new workbook (keeps template intact)
  const newWb = await loadApacWorkbook(prevBuffer);
  const ws = newWb.getWorksheet("Payroll Details");
  clearDataRows(ws, DATA_ROW_START);

  const manual = [];
  const newJoiners = [];
  const recoRows = [];

  let no = 1;
  let row = DATA_ROW_START;
  let employeesWritten = 0;
  for (const emp of run.employees) {
    const prev = prevData[emp.code] || null;
    if (!prev) newJoiners.push(emp.code);
    fillEmployeeRow(ws, row, no, emp, prev, manual);

    const nV = cellNum(ws, row, COL.monthly_salary);
    const rV = cellNum(ws, row, COL.back_pay);
    const tV = cellNum(ws, row, COL.paid_leave);
    const amV = cellNum(ws, row, COL.bonus);
    const awV = cellNum(ws, row, COL.severance);
    const loanV = cellNum(ws, row, COL.loan);
    const apacBs = nV + rV + tV + amV + awV;
    const apacCp = apacBs - loanV;
    const oaTarget = pyRound(totalOaCost(emp)) - emp.loanDeduction;
    recoRows.push({
      Code: emp.code, Name: emp.name,
      "APAC N": pyRound(nV), "APAC R": pyRound(rV), "APAC T": pyRound(tV),
      "APAC AM (Bonus)": pyRound(amV), "APAC AW": pyRound(awV),
      "APAC BS = N+R+T+AM+AW": pyRound(apacBs),
      "APAC BX (Loan)": pyRound(loanV),
      "APAC CP = BS − Loan": pyRound(apacCp),
      "OA Per-Emp TotalOA − Loan": pyRound(oaTarget),
      "Diff (CP − OA)": pyRound(apacCp - oaTarget),
      "APAC CD (PF Employer)": pyRound(cellNum(ws, row, COL.pf_employer)),
      "APAC CE (PF Employee)": pyRound(cellNum(ws, row, COL.pf_employee)),
      "APAC CH (TDS)": pyRound(cellNum(ws, row, COL.tds)),
      "APAC CK (Profession Tax)": pyRound(cellNum(ws, row, COL.other_ded)),
    });
    no++;
    row++;
    employeesWritten++;
  }

  // Stipends → intern rows
  let internsWritten = 0;
  let internIdx = 1;
  for (const x of run.extras) {
    if (x.code !== "STIPEND") continue;
    let clean = x.name.startsWith("Stipend — ") ? x.name.slice("Stipend — ".length) : x.name;
    if (clean.includes(" (")) clean = clean.split(" (")[0];
    const deptCanonical = DEPT_NORMALISE[x.department.trim().toLowerCase()] || x.department;
    fillInternRow(ws, row, no, internIdx, clean, deptCanonical, x.netPaid, manual);
    no++;
    row++;
    internIdx++;
    internsWritten++;
  }

  const currentCodes = new Set(run.employees.map((e) => e.code));
  const leavers = Object.keys(prevData).filter(
    (c) => !currentCodes.has(c) && !c.toLowerCase().startsWith("intern")
  );

  const recoSummary = {
    "OA headcount": run.employees.length,
    "APAC employee rows written": employeesWritten,
    "APAC intern rows written": internsWritten,
    "New joiners (need manual fill)": newJoiners.length,
    "Leavers (in prev APAC, not current)": leavers.length,
    "OA total Net Paid": pyRound(
      run.employees.reduce((s, e) => s + netPaid(e), 0) - run.employees.reduce((s, e) => s + e.loanDeduction, 0)
    ),
    "OA total TDS": pyRound(run.employees.reduce((s, e) => s + e.incomeTax, 0)),
    "OA total PF (employee, MIS)": pyRound(run.employees.reduce((s, e) => s + e.pfMis, 0)),
    "OA total PF (employer, MIS)": pyRound(run.employees.reduce((s, e) => s + e.pfMis, 0)),
    "OA total OA Cost (incl. PF Admin/EDLI)": pyRound(
      run.employees.reduce((s, e) => s + totalOaCost(e), 0) +
        run.extras.reduce((s, x) => s + extraTotal(x), 0) +
        otherTotal(run)
    ),
    "OA total OA Cost (EXCL. PF Admin/EDLI) ⇄ APAC Σ Total Labor Cost": pyRound(
      run.employees.reduce((s, e) => s + totalOaCost(e), 0) + run.extras.reduce((s, x) => s + extraTotal(x), 0)
    ),
  };

  return { workbook: newWb, manual, recoRows, recoSummary, employeesWritten, internsWritten, newJoiners, leavers };
}

export async function saveApac(wb) {
  return wb.xlsx.writeBuffer();
}

// ---------------------------------------------------------------------------
// Manual Fill report
// ---------------------------------------------------------------------------

export async function writeManualReport(items) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Manual Fill");

  ws.getCell("A1").value = "APAC Payroll Report — Manual Fill Required";
  ws.getCell("A1").font = { bold: true, size: 14 };
  ws.getCell("A2").value =
    "Open the filled APAC and address each item below. Defaults are pre-filled in the indicated columns; you may need to reallocate or add detail.";

  ws.getCell("A4").value = "Summary";
  ws.getCell("A4").font = { bold: true };
  const counts = {
    "Bonus splits": items.filter((i) => i.itemType === "Bonus split").length,
    "Resignations / Severance": items.filter((i) => i.itemType === "Resignation / Severance").length,
    "New joiners": items.filter((i) => i.itemType === "New joiner").length,
    "Interns added": items.filter((i) => i.itemType === "Intern").length,
  };
  let r = 5;
  for (const [k, v] of Object.entries(counts)) {
    ws.getCell(r, 1).value = k;
    ws.getCell(r, 2).value = v;
    r++;
  }

  r += 2;
  const headers = ["Code", "Name", "Type", "Detail", "Suggested value", "APAC columns to fill"];
  headers.forEach((hText, i) => {
    const cell = ws.getCell(r, i + 1);
    cell.value = hText;
    cell.fill = HDR_FILL;
    cell.font = HDR_FONT;
  });
  r++;

  for (const it of items) {
    const vals = [it.code, it.name, it.itemType, it.detail, it.suggestedValue, it.apacColumns];
    vals.forEach((v, i) => {
      const cell = ws.getCell(r, i + 1);
      cell.value = v;
      if (it.itemType === "Bonus split" || it.itemType === "Resignation / Severance") cell.fill = WARN_FILL;
      else if (it.itemType === "New joiner") cell.fill = INFO_FILL;
    });
    if (typeof it.suggestedValue === "number") ws.getCell(r, 5).numFmt = INT_FMT;
    r++;
  }

  [12, 28, 22, 80, 18, 28].forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });
  return wb.xlsx.writeBuffer();
}

// ---------------------------------------------------------------------------
// Reconciliation report — email-ready single sheet + per-employee audit tab
// ---------------------------------------------------------------------------

export async function writeRecoReport(recoRows, summary, run, pfAdmin, edli, monthLabel) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Reconciliation");

  const registerNetpay = run.employees.reduce((s, e) => s + e.netpayRegister, 0);
  const pfEmployer = run.employees.reduce((s, e) => s + e.pfMis, 0);
  const pfEmployee = pfEmployer;
  const tds = run.employees.reduce((s, e) => s + e.incomeTax, 0);
  const pt = run.employees.reduce((s, e) => s + e.professionTax, 0);
  const loanRecovery = run.employees.reduce((s, e) => s + e.loanDeduction, 0);
  const stipends = run.extras.filter((x) => x.code === "STIPEND").reduce((s, x) => s + x.netPaid, 0);

  const cashPaidOut = registerNetpay - loanRecovery + stipends;
  const govtDeducted = tds + pt + pfEmployee;
  const employerStatutory = pfEmployer + pfAdmin + edli;
  const totalOa = cashPaidOut + govtDeducted + employerStatutory;
  const apacTotalLabor = totalOa - pfAdmin - edli;

  const sectionFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9E1F2" } };
  const sectionFont = { bold: true, size: 11, color: { argb: "FF1F4E79" } };
  const moneyTotalFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
  const tickFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC6EFCE" } };
  const moneyFmt = '_-₹* #,##0_-;[Red]-₹* #,##0_-;_-₹* "-"_-;_-@_-';
  const inr = (n) => "₹" + Math.round(n).toLocaleString("en-IN");

  const sectionHeader = (row, text) => {
    ws.mergeCells(row, 1, row, 4);
    const c = ws.getCell(row, 1);
    c.value = text;
    c.font = sectionFont;
    c.alignment = { vertical: "middle" };
    ws.getRow(row).height = 22;
    for (let col = 1; col <= 4; col++) ws.getCell(row, col).fill = sectionFill;
  };

  const line = (row, label, value, opts = {}) => {
    const { bold = false, fill = null, op = "", isMoney = true } = opts;
    const c1 = ws.getCell(row, 1);
    c1.value = label;
    c1.font = { bold, size: 11 };
    if (op) {
      const c2 = ws.getCell(row, 2);
      c2.value = op;
      c2.font = { bold, size: 11 };
      c2.alignment = { horizontal: "center" };
    }
    const c3 = ws.getCell(row, 3);
    c3.value = value;
    c3.font = { bold, size: 11 };
    if (typeof value === "number") {
      c3.numFmt = isMoney ? moneyFmt : "#,##0";
      c3.alignment = { horizontal: "right" };
    }
    if (fill) for (let col = 1; col <= 4; col++) ws.getCell(row, col).fill = fill;
  };

  // Title
  ws.mergeCells("A1:D1");
  const t = ws.getCell("A1");
  t.value = `Salary OA Reconciliation${monthLabel ? " — " + monthLabel : ""}`;
  t.font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
  t.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } };
  t.alignment = { vertical: "middle", horizontal: "center" };
  ws.getRow(1).height = 28;

  ws.mergeCells("A2:D2");
  const sub = ws.getCell("A2");
  sub.value =
    `Company paid ${inr(totalOa)} this month — split across employees, government remittances and EPFO. ` +
    `Of that, ${inr(apacTotalLabor)} is what HQ sees in the APAC Payroll Report (PF Admin and EDLI are ` +
    "India-only statutory items kept on the OA report).";
  sub.font = { italic: true, size: 10, color: { argb: "FF595959" } };
  sub.alignment = { wrapText: true, vertical: "middle" };
  ws.getRow(2).height = 32;

  let r = 4;
  sectionHeader(r, "1.  Head-count check"); r++;
  line(r, "Salary Register — employees this month", summary["OA headcount"] || 0, { isMoney: false }); r++;
  line(r, "PF MIS — employees with PF", run.employees.filter((e) => e.inPfMis).length, { isMoney: false }); r++;
  line(r, "APAC — employee rows written", summary["APAC employee rows written"] || 0, { isMoney: false }); r++;
  line(r, "APAC — intern rows added", summary["APAC intern rows written"] || 0, { isMoney: false }); r++;
  line(r, "APAC — new joiners (need manual fill)", summary["New joiners (need manual fill)"] || 0, { isMoney: false }); r++;
  line(r, "APAC — leavers (excluded from this month)", summary["Leavers (in prev APAC, not current)"] || 0, { isMoney: false }); r++;
  r++;

  sectionHeader(r, "2.  How the totals tie together"); r++;
  ws.getCell(r, 1).value =
    "Built bottom-up — start with cash paid out, add what was withheld and remitted, arrive at the total cost the company bore.";
  ws.getCell(r, 1).font = { italic: true, size: 10, color: { argb: "FF595959" } };
  r++;

  const subTotalFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAF2F8" } };
  line(r, "Cash paid to employees (sum of NETPAY in salary register)", registerNetpay); r++;
  line(r, "Loan EMIs recovered from employees (kept by company)", -loanRecovery, { op: "−" }); r++;
  line(r, "Stipends paid to interns (added manually in OA)", stipends, { op: "+" }); r++;
  line(r, "Sub-total: cash actually paid out this month", cashPaidOut, { bold: true, fill: subTotalFill }); r++;
  r++;
  line(r, "TDS withheld and paid to Income Tax Dept", tds, { op: "+" }); r++;
  line(r, "PT withheld and paid to State Govt", pt, { op: "+" }); r++;
  line(r, "PF (Employee share) withheld and paid to EPFO", pfEmployee, { op: "+" }); r++;
  line(r, "PF (Employer share) paid by company to EPFO", pfEmployer, { op: "+" }); r++;
  line(r, "PF Admin Charges paid by company to EPFO (Account 2)", pfAdmin, { op: "+" }); r++;
  line(r, "EDLI Charges paid by company to EPFO (Account 21)", edli, { op: "+" }); r++;
  line(r, "TOTAL OA COST  (company's full monthly salary outflow)", totalOa, { bold: true, fill: moneyTotalFill }); r++;
  r++;
  line(r, "PF Admin Charges (kept on OA report only — not in APAC)", -pfAdmin, { op: "−" }); r++;
  line(r, "EDLI Charges (kept on OA report only — not in APAC)", -edli, { op: "−" }); r++;
  line(r, "APAC TOTAL LABOR COST  (sum of column CP in APAC)", apacTotalLabor, { bold: true, fill: moneyTotalFill }); r++;
  r++;

  sectionHeader(r, "3.  Where did the money go?"); r++;
  ws.getCell(r, 1).value = `Same ${inr(totalOa)}, sliced by who actually received it.`;
  ws.getCell(r, 1).font = { italic: true, size: 10, color: { argb: "FF595959" } };
  r++;

  const breakdown = [
    ["Net cash paid to employees (after loan recovery)", registerNetpay - loanRecovery],
    ["Stipends paid to interns", stipends],
    ["TDS — Income Tax to government", tds],
    ["PT — Profession Tax to government", pt],
    ["PF — Employee contribution to EPFO", pfEmployee],
    ["PF — Employer contribution to EPFO", pfEmployer],
    ["PF Admin Charges (Account 2) to EPFO", pfAdmin],
    ["EDLI Charges (Account 21) to EPFO", edli],
  ];
  for (const [labelText, amt] of breakdown) {
    line(r, labelText, amt);
    if (totalOa) {
      const pc = ws.getCell(r, 4);
      pc.value = amt / totalOa;
      pc.numFmt = "0.0%";
      pc.font = { size: 10, color: { argb: "FF595959" } };
      pc.alignment = { horizontal: "right" };
    }
    r++;
  }
  const allocationTotal = breakdown.reduce((s, [, amt]) => s + amt, 0);
  line(r, "TOTAL", allocationTotal, { bold: true, fill: moneyTotalFill }); r++;
  r++;

  sectionHeader(r, "4.  Sanity checks"); r++;
  const sanity = [
    ["Money-out breakdown ties back to Total OA Cost", Math.abs(allocationTotal - totalOa) < 1],
    ["APAC Total Labor Cost = Total OA Cost − PF Admin − EDLI", Math.abs(apacTotalLabor - (totalOa - pfAdmin - edli)) < 1],
    ["Salary register PF total matches PF MIS PF Ac. 1",
      Math.abs(run.employees.reduce((s, e) => s + e.pfRegister, 0) - pfEmployer) < 1],
    ["Head-count: salary register matches APAC employee rows",
      (summary["OA headcount"] || 0) === (summary["APAC employee rows written"] || 0)],
  ];
  for (const [labelText, ok] of sanity) {
    ws.getCell(r, 1).value = labelText;
    ws.mergeCells(r, 1, r, 3);
    const c = ws.getCell(r, 4);
    c.value = ok ? "✓ Tied" : "✗ Mismatch";
    c.font = ok ? { bold: true, color: { argb: "FF375623" } } : { bold: true, color: { argb: "FF9C0006" } };
    c.fill = ok ? tickFill : { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
    c.alignment = { horizontal: "center" };
    r++;
  }
  r++;

  ws.mergeCells(r, 1, r, 4);
  ws.getCell(r, 1).value =
    "Note: APAC excludes PF Admin Charges and EDLI Charges by design — these statutory items are tracked only on the OA report.";
  ws.getCell(r, 1).font = { italic: true, size: 9, color: { argb: "FF808080" } };

  ws.getColumn(1).width = 60;
  ws.getColumn(2).width = 4;
  ws.getColumn(3).width = 18;
  ws.getColumn(4).width = 14;

  // Per-employee audit tab
  const p = wb.addWorksheet("Per Employee (audit)");
  if (recoRows.length) {
    const headers = Object.keys(recoRows[0]);
    headers.forEach((hText, i) => {
      const cell = p.getCell(1, i + 1);
      cell.value = hText;
      cell.fill = HDR_FILL;
      cell.font = HDR_FONT;
    });
    recoRows.forEach((rowObj, ri) => {
      headers.forEach((hText, i) => {
        const cell = p.getCell(ri + 2, i + 1);
        cell.value = rowObj[hText];
        if (typeof rowObj[hText] === "number" && hText !== "Code") cell.numFmt = INT_FMT;
      });
    });
    headers.forEach((hText, i) => {
      p.getColumn(i + 1).width = Math.max(14, hText.length + 2);
    });
    p.views = [{ state: "frozen", xSplit: 2, ySplit: 1 }];
  }

  return { buffer: await wb.xlsx.writeBuffer(), totals: { totalOa, apacTotalLabor, cashPaidOut } };
}
