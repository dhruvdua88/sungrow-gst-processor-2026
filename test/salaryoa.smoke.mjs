// Smoke test for the Salary OA / APAC tab engine.
// The Drive folder ships the April 2026 OUTPUT files (no raw inputs), so we
// reconstruct a salary register + PF MIS from the "COMBINED" raw-extract tab
// of the main TEST workbook, re-run the engine, and assert the totals match
// the desktop tool's outputs exactly. The "APAC Filled" TEST file doubles as
// the previous-month APAC template input.
//
// Run: node test/salaryoa.smoke.mjs [path-to-test-folder]
import { readFileSync } from "node:fs";
import { join } from "node:path";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { build, validate, writeWorkbook, splitExports, makeExtraRow } from "../src/salaryoa/engine.js";
import { buildApac, saveApac, writeManualReport, writeRecoReport } from "../src/salaryoa/apac.js";

const DIR = process.argv[2] || "/Users/dhruvdua88/Documents/Gog CLI/apac-salary-oa";
const buf = (name) => {
  const b = readFileSync(join(DIR, name));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
};

let failures = 0;
function check(label, actual, expected, tol = 0) {
  const ok = typeof expected === "number" ? Math.abs(actual - expected) <= tol : actual === expected;
  console.log(`${ok ? "  OK " : "  FAIL"}  ${label}: ${actual}${ok ? "" : `  (expected ${expected})`}`);
  if (!ok) failures++;
}

// ---------------------------------------------------------------------------
// 1. Reconstruct inputs from the COMBINED tab of the TEST main workbook
// ---------------------------------------------------------------------------
const mainWb = XLSX.read(buf("Salary OA - April 2026 TEST.xlsx"), { type: "array" });
const combined = XLSX.utils.sheet_to_json(mainWb.Sheets["COMBINED"], { header: 1 });
const [, ...empRows] = combined;

// Salary register sheet (same headers the loader expects)
const salaryHeader = ["Code", "Name", "Department", "Total Earnings", "Income Tax",
  "Provident Fund", "Profession Tax", "NETPAY"];
const salaryRows = empRows.map((r) => [r[0], r[1], r[2], r[3], r[4], r[5], r[7], r[8]]);
const salaryWs = XLSX.utils.aoa_to_sheet([salaryHeader, ...salaryRows]);
const salaryWb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(salaryWb, salaryWs, "Salary Sheet");
const salaryBuf = XLSX.write(salaryWb, { type: "array", bookType: "xlsx" });

// PF MIS sheet — per-employee PF Ac. 1; Admin/EDLI totals spread on first row
const PF_ADMIN = 18090;
const EDLI = 3450;
const pfHeader = ["Code", "PF Ac. 1", "Admin Charges Ac. 2", "EDLI Charges Ac. 21"];
const pfRows = empRows
  .filter((r) => Number(r[6]) !== 0)
  .map((r, i) => [r[0], r[6], i === 0 ? PF_ADMIN : 0, i === 0 ? EDLI : 0]);
const pfWs = XLSX.utils.aoa_to_sheet([pfHeader, ...pfRows]);
const pfWb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(pfWb, pfWs, "PF MIS");
const pfBuf = XLSX.write(pfWb, { type: "array", bookType: "xlsx" });

// ---------------------------------------------------------------------------
// 2. Run the OA engine with April's stipends + loan
// ---------------------------------------------------------------------------
const extras = [
  makeExtraRow({ code: "STIPEND", name: "Stipend — Itesh Hasija (Marketing, Gurgaon)", department: "OPERATIONS", netPaid: 10000 }),
  makeExtraRow({ code: "STIPEND", name: "Stipend — Chirag Dhankhar (Operations, Gurgaon)", department: "OPERATIONS", netPaid: 14000 }),
];
const res = build(salaryBuf, pfBuf, { extras, monthLabel: "April 2026" });

const divya = res.run.employees.find((e) => e.name === "Divya Kukreja");
if (!divya) throw new Error("Divya Kukreja not found in reconstructed register");
divya.loanDeduction += 7693;
const report = validate(res.run, res.pfLookup, res.pfAdminSource);

console.log("\n--- OA validation checks (vs desktop tool's April values) ---");
check("Headcount", res.run.employees.length, 47);
const byName = Object.fromEntries(report.checks.map((c) => [c.name.split("  ")[0], c]));
check("Net Pay source", byName["Net Pay"].sourceValue, 6656478, 1);
check("Net Pay ok", byName["Net Pay"].ok, true);
check("TDS source", byName["TDS"].sourceValue, 1090290, 1);
check("TDS ok", byName["TDS"].ok, true);
check("PT source", byName["PT"].sourceValue, 3000, 1);
check("PF Employee source", byName["PF Employee"].sourceValue, 434168, 1);
check("PF Employee ok", byName["PF Employee"].ok, true);
check("PF Employer ok", byName["PF Employer"].ok, true);
check("Other Costs source", byName["Other Costs"].sourceValue, 21540, 1);
check("Other Costs ok", byName["Other Costs"].ok, true);
check("Identity ok", byName["Identity"].ok, true);
check("All checks pass", report.allOk, true);

// ---------------------------------------------------------------------------
// 3. Write workbooks, re-read, compare totals against the TEST outputs
// ---------------------------------------------------------------------------
const mainBuffer = await writeWorkbook(res.run, report);
const splits = await splitExports(res.run);

async function lastTotalsRow(buffer, sheetName) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = sheetName ? wb.getWorksheet(sheetName) : wb.worksheets[0];
  let totals = null;
  ws.eachRow((row) => {
    if (row.getCell(1).value === "Total") {
      totals = [4, 5, 6, 7, 8, 9].map((c) => {
        const v = row.getCell(c).value;
        return typeof v === "object" && v !== null ? v.result ?? 0 : v;
      });
    }
  });
  return totals;
}

console.log("\n--- Output workbook totals (vs TEST output files) ---");
const fc = await lastTotalsRow(splits["Final Combined"]);
check("Final Combined: Net Paid", fc[0], 6680478);
check("Final Combined: TDS", fc[1], 1090290);
check("Final Combined: PT", fc[2], 3000);
check("Final Combined: PF Govt3", fc[3], 434168);
check("Final Combined: PF Govt4", fc[4], 455708);
check("Final Combined: Total OA Cost", fc[5], 8663644);

const sc = await lastTotalsRow(splits["Service Cost"]);
check("Service Cost: Total OA Cost", sc[5], 1749057);
check("Service Cost: Net Paid", sc[0], 1454870);

const nsc = await lastTotalsRow(splits["Non Service Cost"]);
check("Non Service: Total OA Cost", nsc[5], 6914587);
check("Non Service: Net Paid", nsc[0], 5225608);

// Summary tab of the main workbook
{
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(mainBuffer);
  const s = wb.getWorksheet("Summary");
  check("Summary: Service", s.getCell("B5").value, 1749057);
  check("Summary: Non service", s.getCell("B6").value, 6914587);
  check("Summary: Grand Total", s.getCell("B7").value, 8663644);
  check("Main workbook sheets", wb.worksheets.map((w) => w.name).join("|"),
    "Validation|Summary|Final Combined (Working)|Service Cost|Non Service Cost|Final Combined|COMBINED");
}

// ---------------------------------------------------------------------------
// 4. APAC build — previous APAC = the April Filled TEST file (unencrypted)
// ---------------------------------------------------------------------------
console.log("\n--- APAC build (prev = APAC Filled TEST, unencrypted) ---");
const apacRes = await buildApac(buf("APAC Filled - April 2026 TEST.xlsx"), res.run);
check("APAC employees written", apacRes.employeesWritten, 47);
check("APAC interns written", apacRes.internsWritten, 2);
check("APAC new joiners", apacRes.newJoiners.length, 0);
// Python's summary line sums employees + stipends + other costs WITHOUT the
// negative LOANRECOVERY rows → Final Combined total (8663644) + loan (7693).
check("OA total OA Cost (incl admin/EDLI)",
  apacRes.recoSummary["OA total OA Cost (incl. PF Admin/EDLI)"], 8671337);

const apacBuffer = await saveApac(apacRes.workbook);
const manualBuffer = await writeManualReport(apacRes.manual);
const reco = await writeRecoReport(apacRes.recoRows, apacRes.recoSummary, res.run, PF_ADMIN, EDLI, "April 2026");
check("APAC xlsx bytes > 50k", apacBuffer.byteLength > 50000, true);
check("Manual report bytes > 4k", manualBuffer.byteLength > 4000, true);
check("Reco report bytes > 6k", reco.buffer.byteLength > 6000, true);
check("Reco Total OA Cost", Math.round(reco.totals.totalOa), 8663644);
check("Reco APAC Total Labor", Math.round(reco.totals.apacTotalLabor), 8663644 - PF_ADMIN - EDLI);

// Re-open the filled APAC and verify structure
{
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(apacBuffer);
  const ws = wb.getWorksheet("Payroll Details");
  check("APAC row 4 code", String(ws.getCell("C4").value), res.run.employees[0].code);
  check("APAC row 4 CP formula", ws.getCell("CP4").formula || "", "SUM(CI4,CL4,CN4)");
  const internRow = 4 + 47;
  check("APAC intern row code", String(ws.getCell(`C${internRow}`).value), "Intern1");
  check("APAC intern name (suffix stripped)", String(ws.getCell(`D${internRow}`).value), "Itesh Hasija");
}

console.log(failures === 0 ? "\nSMOKE TEST PASSED — all assertions OK." : `\nSMOKE TEST FAILED — ${failures} assertion(s).`);
process.exit(failures === 0 ? 0 : 1);
