// End-to-end engine test with synthetic data (node, no browser needed).
// Run: node test/engine.test.mjs
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { writeFileSync } from "node:fs";
import { buildWorkbook } from "../src/engine/index.js";

let failures = 0;
function check(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  console.log(`${ok ? "PASS" : "FAIL"}  ${label}: ${JSON.stringify(actual)}${ok ? "" : ` (expected ${JSON.stringify(expected)})`}`);
  if (!ok) failures += 1;
}

// ---- Synthetic GST workbook: PO / Non PO / GL Summary ----
const wb = XLSX.utils.book_new();

const poRows = [
  // Exact 2B match (B2B)
  { "Supplier Name": "Acme Industries Private Limited", "SAP Invoice No.": "5100000001",
    "Original Invoice No.": "INV-001", "SAP Invoice Date": "01-04-2026", "GST No.": "29AAACA1111A1Z5",
    "Country": "IN", "Description": "Karnataka", "Basic Amount": 100000, "IGST Amount": 0,
    "CGST Amount": 9000, "SGST Amount": 9000, "RCM": "N" },
  // Tax mismatch vs 2B (books higher by 100)
  { "Supplier Name": "Beta Logistics Pvt Ltd", "SAP Invoice No.": "5100000002",
    "Original Invoice No.": "INV-002", "SAP Invoice Date": "05-04-2026", "GST No.": "29AAACB2222B1Z6",
    "Country": "IN", "Description": "Karnataka", "Basic Amount": 50000, "IGST Amount": 9100,
    "CGST Amount": 0, "SGST Amount": 0, "RCM": "N" },
  // Wrong RCM flag in PO; 2B says normal B2B (correction flag expected)
  { "Supplier Name": "Gamma Tech Solutions Ltd", "SAP Invoice No.": "5100000003",
    "Original Invoice No.": "INV-003", "SAP Invoice Date": "08-04-2026", "GST No.": "29AAACG3333C1Z7",
    "Country": "IN", "Description": "Karnataka", "Basic Amount": 20000, "IGST Amount": 3600,
    "CGST Amount": 0, "SGST Amount": 0, "RCM": "Y" },
  // NA supplier — skipped
  { "Supplier Name": "NA", "SAP Invoice No.": "5100000099", "Original Invoice No.": "INV-099",
    "SAP Invoice Date": "08-04-2026", "GST No.": "", "Country": "IN", "Description": "",
    "Basic Amount": 999, "IGST Amount": 1, "CGST Amount": 0, "SGST Amount": 0, "RCM": "N" },
];
XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(poRows), "PO");

const nonPoRows = [
  // Unresolved B2B (no GSTIN, not in 2B)
  { "Company Code": "2081", "Supplier Name": "Delta Services", "Document No": "1900000001",
    "Reference": "DS-44", "Country": "IN", "State Name": "Karnataka", "GST No.": "",
    "Basic Amount": 10000, "IGST Amount": 1800, "CGST Amount": 0, "SGST Amount": 0 },
  // Import (matches 2B IMPG BoE by IGST)
  { "Company Code": "2081", "Supplier Name": "Overseas Supplier GmbH", "Document No": "1900000002",
    "Reference": "OS-IMP-1", "Country": "DE", "State Name": "", "GST No.": "",
    "Basic Amount": 0, "IGST Amount": 25000, "CGST Amount": 0, "SGST Amount": 0 },
  // Other company code — filtered out
  { "Company Code": "9999", "Supplier Name": "Other Co", "Document No": "1900000003",
    "Reference": "OC-1", "Country": "IN", "State Name": "", "GST No.": "",
    "Basic Amount": 5000, "IGST Amount": 900, "CGST Amount": 0, "SGST Amount": 0 },
];
XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(nonPoRows), "Non PO");

const glRows = [
  // REC postings for INV-001 (ties: 9000 + 9000)
  { "Company Code": "2081", "Account": "2221013220", "G/L Name": "CGST REC", "Tax code": "",
    "Document Number": "5100000001", "Reference": "INV-001", "Invoice reference": "",
    "Document Date": "01-04-2026", "Posting Date": "01-04-2026", "Amount in local currency": 9000 },
  { "Company Code": "2081", "Account": "2221013210", "G/L Name": "SGST REC", "Tax code": "",
    "Document Number": "5100000001", "Reference": "INV-001", "Invoice reference": "",
    "Document Date": "01-04-2026", "Posting Date": "01-04-2026", "Amount in local currency": 9000 },
  // REC posting for INV-002 — short by 100 => ledger tie mismatch
  { "Company Code": "2081", "Account": "2221013230", "G/L Name": "IGST REC", "Tax code": "",
    "Document Number": "5100000002", "Reference": "INV-002", "Invoice reference": "",
    "Document Date": "05-04-2026", "Posting Date": "05-04-2026", "Amount in local currency": 9000 },
  // REC posting for INV-003 ties (3600)
  { "Company Code": "2081", "Account": "2221013230", "G/L Name": "IGST REC", "Tax code": "",
    "Document Number": "5100000003", "Reference": "INV-003", "Invoice reference": "",
    "Document Date": "08-04-2026", "Posting Date": "08-04-2026", "Amount in local currency": 3600 },
  // REC posting for Delta Services 1800 ties
  { "Company Code": "2081", "Account": "2221013230", "G/L Name": "IGST REC", "Tax code": "",
    "Document Number": "1900000001", "Reference": "DS-44", "Invoice reference": "",
    "Document Date": "10-04-2026", "Posting Date": "10-04-2026", "Amount in local currency": 1800 },
  // Import REC 25000 ties
  { "Company Code": "2081", "Account": "2221013230", "G/L Name": "IGST REC", "Tax code": "",
    "Document Number": "1900000002", "Reference": "OS-IMP-1", "Invoice reference": "",
    "Document Date": "12-04-2026", "Posting Date": "12-04-2026", "Amount in local currency": 25000 },
];
XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(glRows), "GL Summary");

const gstWorkbookBuffer = XLSX.write(wb, { type: "array", bookType: "xlsx" });

// Tracker template (content unused, presence validated — same as the Python tool)
const trackerWb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(trackerWb, XLSX.utils.json_to_sheet([{ note: "template" }]), "Sheet1");
const trackerBuffer = XLSX.write(trackerWb, { type: "array", bookType: "xlsx" });

// ---- Synthetic GSTR-2B JSON ----
const r2bJson = {
  data: {
    rtnprd: "042026",
    docdata: {
      b2b: [
        { ctin: "29AAACA1111A1Z5", trdnm: "ACME INDUSTRIES PVT LTD", supprd: "042026",
          inv: [{ inum: "INV-001", dt: "01-04-2026", txval: 100000, igst: 0, cgst: 9000, sgst: 9000,
                  cess: 0, rev: "N", itcavl: "Y", pos: "29", typ: "R" }] },
        { ctin: "29AAACB2222B1Z6", trdnm: "BETA LOGISTICS PVT LTD", supprd: "042026",
          inv: [{ inum: "INV-002", dt: "05-04-2026", txval: 50000, igst: 9000, cgst: 0, sgst: 0,
                  cess: 0, rev: "N", itcavl: "Y", pos: "29", typ: "R" }] },
        { ctin: "29AAACG3333C1Z7", trdnm: "GAMMA TECH SOLUTIONS LTD", supprd: "042026",
          inv: [{ inum: "INV-003", dt: "08-04-2026", txval: 20000, igst: 3600, cgst: 0, sgst: 0,
                  cess: 0, rev: "N", itcavl: "Y", pos: "29", typ: "R" }] },
        // Only in 2B
        { ctin: "29AAACZ9999Z1Z9", trdnm: "ZETA UNSEEN SUPPLIES", supprd: "042026",
          inv: [{ inum: "INV-ZZ", dt: "20-04-2026", txval: 7000, igst: 1260, cgst: 0, sgst: 0,
                  cess: 0, rev: "N", itcavl: "Y", pos: "29", typ: "R" }] },
      ],
      cdnr: [],
      impg: [
        { boenum: "7654321", boedt: "11-04-2026", txval: 138888, igst: 25000, cgst: 0, sgst: 0,
          cess: 0, portcode: "INMAA1" },
      ],
      isd: [],
    },
  },
};

// ---- Synthetic ISD JSON (2 suppliers; supplier A eligible via subset sum) ----
const isdJson = {
  b2b: [
    { ctin: "27AAACI1111I1Z1",
      inv: [{ inum: "ISD-B1", idt: "02-04-2026", val: 11800, pos: "27",
              itms: [{ itm_det: { rt: 18, txval: 10000, iamt: 1800, camt: 0, samt: 0, csamt: 0 } }] }] },
    { ctin: "27AAACJ2222J1Z2",
      inv: [{ inum: "ISD-B2", idt: "03-04-2026", val: 5900, pos: "27",
              itms: [{ itm_det: { rt: 18, txval: 5000, iamt: 900, camt: 0, samt: 0, csamt: 0 } }] }] },
  ],
  isd: {
    elglst: [
      { cpty: "SIPL KARNATAKA", statecd: "29",
        doclst: [{ docnum: "ISD0001", docdt: "30-04-2026", iamt: 1800, camt: 0, samt: 0, csamt: 0 }] },
    ],
    inelglst: [
      { cpty: "SIPL KARNATAKA", statecd: "29",
        doclst: [{ docnum: "ISD0002", docdt: "30-04-2026", iamt: 900, camt: 0, samt: 0, csamt: 0 }] },
    ],
  },
};

const res = await buildWorkbook({
  trackerBuffer,
  gstWorkbookBuffer,
  r2bJson,
  priorFiles: [],
  boeFiles: [],
  isdJson,
  periodLabel: "April",
  fyLabel: "2026-27",
  onProgress: (m) => console.log(`  · ${m}`),
});

console.log("\n--- assertions ---");
check("invoice_count", res.invoice_count, 5);
check("r2b_count", res.r2b_count, 5);
check("books_total_tax", res.books_total_tax, 18000 + 9100 + 3600 + 1800 + 25000);
check("r2b_total_tax", res.r2b_total_tax, 18000 + 9000 + 3600 + 1260 + 25000);
check("match Match count", res.match_counts["Match"], 3);
check("match Tax mismatch count", res.match_counts["Tax mismatch"], 1);
check("unresolved count", res.unresolved.length, 1);
check("correction_count >= 2 (tax mismatch + RCM flag)", res.correction_count >= 2, true);
check("ledger ITC-REC diff (INV-002 short 100)", res.ledger_itc_rec_difference, 100);
check("isd_bill_count", res.isd_bill_count, 2);
check("isd_bill_tax", res.isd_bill_tax, 2700);
check("isd_eligible_tax (ISD-002)", res.isd_eligible_tax, 1800);
check("isd eligible classifier method", res.isd_eligibility_method, "Supplier-level subset sum (exact)");
check("isd eligible supplier", res.isd_eligible_supplier_gstins, ["27AAACI1111I1Z1"]);
check("isd_eligible_bill_tax", res.isd_eligible_bill_tax, 1800);
check("boe_count (none supplied)", res.boe_count, 0);
check("output name", res.output, "SIPL_GST_ITC_April_2026_27_Focused_Output.xlsx");

// ---- Output workbook integrity ----
writeFileSync("test/out.xlsx", Buffer.from(res.buffer));
const outWb = new ExcelJS.Workbook();
await outWb.xlsx.readFile("test/out.xlsx");
const names = outWb.worksheets.map((s) => s.name);
check("sheet names", names, [
  "Dashboard", "Sanitized ITC", "2B", "ISD", "ISD Eligible",
  "2B vs Books Summary", "2B vs Books Detail",
  "ITC vs Ledger", "Invoice Ledger Detail", "Book Correction Flags",
]);
const itcSheet = outWb.getWorksheet("Sanitized ITC");
check("Sanitized ITC rows (incl header)", itcSheet.rowCount, 6);
check("Sanitized ITC header A1", itcSheet.getCell(1, 1).value, "Party GSTIN/UIN");
const dash = outWb.getWorksheet("Dashboard");
check("Dashboard title", dash.getCell("A1").value, "SIPL GST ITC to Ledger Tie-Out");


// Parity dump (compare with Python build_sipl_gst_itc_tracker.py on same data)
console.log("\n--- parity dump ---");
console.log(JSON.stringify({
  invoice_count: res.invoice_count,
  r2b_count: res.r2b_count,
  books_total_tax: res.books_total_tax,
  r2b_total_tax: res.r2b_total_tax,
  match_counts: res.match_counts,
  gstin_counts: res.gstin_counts,
  unresolved: res.unresolved.length,
  correction_count: res.correction_count,
  correction_amount: res.correction_amount,
  ledger_itc_rec_difference: res.ledger_itc_rec_difference,
  ledger_mismatch_count: res.ledger_mismatch_count,
  isd_bill_count: res.isd_bill_count,
  isd_bill_tax: res.isd_bill_tax,
  isd_eligible_tax: res.isd_eligible_tax,
  isd_eligibility_method: res.isd_eligibility_method,
  isd_eligible_supplier_gstins: res.isd_eligible_supplier_gstins,
  isd_eligible_bill_tax: res.isd_eligible_bill_tax,
}, null, 2));

console.log(failures ? `\n${failures} FAILURES` : "\nALL TESTS PASSED");
process.exit(failures ? 1 : 0);
