// Styled XLSX output via ExcelJS — mirrors _style_header/_set_widths/_format_numbers/
// _write_sheet/build_dashboard and the sheet assembly in build_workbook
import ExcelJS from "exceljs";
import { asText, asFloat, round2, TARGET_COMPANY_CODE, RECIPIENT_GSTIN } from "./helpers.js";

const BLUE = "FF1F4E78";
const LIGHT = "FFEAF2F8";
const WARN = "FFFCE4D6";
const RED = "FF9C0006";
const PURPLE = "FF7030A0";

const solid = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function styleHeader(ws, maxCol, row = 1, fillColor = BLUE) {
  for (let c = 1; c <= maxCol; c++) {
    const cell = ws.getCell(row, c);
    cell.fill = solid(fillColor);
    cell.font = { color: { argb: "FFFFFFFF" }, bold: true };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = { bottom: { style: "thin", color: { argb: "FFD9E2F3" } } };
  }
}

function cellText(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) {
    return `${String(v.getDate()).padStart(2, "0")}/${String(v.getMonth() + 1).padStart(2, "0")}/${v.getFullYear()}`;
  }
  return String(v);
}

function setWidths(ws, maxWidth = 38) {
  const colCount = ws.columnCount;
  const rowLimit = Math.min(ws.rowCount, 150);
  for (let c = 1; c <= colCount; c++) {
    let maxLen = 0;
    for (let r = 1; r <= rowLimit; r++) {
      const v = ws.getCell(r, c).value;
      const text = cellText(v && typeof v === "object" && v.richText ? "" : v);
      maxLen = Math.max(maxLen, text.length);
    }
    ws.getColumn(c).width = Math.min(Math.max(maxLen + 2, 10), maxWidth);
  }
}

function formatNumbers(ws, amountCols, dateCols = []) {
  for (let r = 2; r <= ws.rowCount; r++) {
    for (const c of amountCols) {
      if (c <= ws.columnCount) ws.getCell(r, c).numFmt = "#,##0.00";
    }
    for (const c of dateCols) {
      if (c <= ws.columnCount) ws.getCell(r, c).numFmt = "dd/mm/yyyy";
    }
  }
}

function writeSheet(ws, headers, rows, tableName = null) {
  ws.addRow(headers);
  for (const row of rows) {
    ws.addRow(row.map((v) => (v === undefined ? "" : v)));
  }
  styleHeader(ws, headers.length);
  if (rows.length) {
    ws.autoFilter = { from: "A1", to: `${colLetter(headers.length)}${rows.length + 1}` };
  }
  ws.views = [{ state: "frozen", ySplit: 1 }];
  setWidths(ws);
  // ExcelJS tables conflict with manually written rows; the auto-filter + striped
  // header replicate the visual result of openpyxl's TableStyleMedium2 closely.
  if (tableName && rows.length) {
    for (let r = 2; r <= rows.length + 1; r++) {
      if (r % 2 === 0) {
        for (let c = 1; c <= headers.length; c++) {
          const cell = ws.getCell(r, c);
          if (!cell.fill || cell.fill.pattern !== "solid") cell.fill = solid("FFF2F7FB");
        }
      }
    }
  }
}

export function buildDashboard(wb, ctx) {
  const {
    invoices, docs, ledgerSummary, ledgerDetails, corrections,
    boeDocs, isdDocs, isdDist, periodLabel, fyLabel,
  } = ctx;
  const ws = wb.addWorksheet("Dashboard", { views: [{ showGridLines: false }] });

  ws.getCell("A1").value = "SIPL GST ITC to Ledger Tie-Out";
  ws.getCell("A1").font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
  ws.getCell("A1").fill = solid(BLUE);
  ws.mergeCells("A1:L1");
  ws.getCell("A2").value =
    `Period: ${periodLabel} FY ${fyLabel} | Company Code: ${TARGET_COMPANY_CODE} only | ` +
    "Focus: sanitized ITC, 2B, invoice recon, GST ledger movement, ISD bill-wise";
  ws.getCell("A2").fill = solid(LIGHT);
  ws.mergeCells("A2:L2");

  const booksTax = round2(invoices.reduce((s, i) => s + i.total_tax, 0));
  const booksNonIsd = round2(
    invoices.filter((i) => i.support.isd_flag !== "ISD").reduce((s, i) => s + i.total_tax, 0)
  );
  const r2bTax = round2(docs.filter((d) => d.itc_available === "Y").reduce((s, d) => s + d.total_tax, 0));
  const ledgerRec = round2(ledgerSummary.reduce((s, r) => s + r[13], 0));
  const ledgerRcm = round2(ledgerSummary.reduce((s, r) => s + r[17], 0));
  const tieDiff = round2(booksTax - ledgerRec);
  const booksVs2b = round2(booksNonIsd - r2bTax);
  const corrAmt = round2(corrections.reduce((s, r) => s + asFloat(r[9]), 0));

  const unresolvedB2b = invoices.filter((i) => i.itc_type === "B2B" && i.gstin_match === "Unresolved");
  const rcmUr = invoices.filter((i) => i.itc_type === "RCM-UR");
  const isdEligible = isdDist.filter((d) => d.eligibility === "Eligible");
  const isdEligibleTax = round2(isdEligible.reduce((s, d) => s + d.total_tax, 0));
  const isdBillTax = round2(isdDocs.reduce((s, d) => s + d.total_tax, 0));
  const isdGstinsInItc = new Set(isdDocs.filter((d) => d.gstin_in_itc).map((d) => d.gstin));
  const isdGstinsAbsent = new Set(isdDocs.filter((d) => !d.gstin_in_itc).map((d) => d.gstin));

  const cards = [
    ["Sanitized ITC Tax", booksTax],
    ["GST REC Ledger Movement", ledgerRec],
    ["ITC - Ledger Difference", tieDiff],
    ["2B Available Tax", r2bTax],
    ["Books Non-ISD - 2B Diff", booksVs2b],
    ["RCM Payable Movement", ledgerRcm],
    ["Immediate Correction Flags", corrections.length],
    ["Correction Amount", corrAmt],
    ["Unresolved B2B Rows", unresolvedB2b.length],
    ["Unresolved B2B Tax", round2(unresolvedB2b.reduce((s, i) => s + i.total_tax, 0))],
    ["RCM-UR Rows", rcmUr.length],
    ["RCM-UR Tax", round2(rcmUr.reduce((s, i) => s + i.total_tax, 0))],
    ["ISD Bill Rows", isdDocs.length],
    ["ISD Bill Tax", isdBillTax],
    ["Eligible ISD Tax (ISD-002)", isdEligibleTax],
    ["ISD Bill - Eligible Diff", isdDocs.length || isdEligible.length ? round2(isdBillTax - isdEligibleTax) : 0],
    ["Eligible Bills (Classified)", isdDocs.filter((d) => d.eligibility === "Eligible").length],
    ["Eligible Bill Tax (Classified)", round2(isdDocs.filter((d) => d.eligibility === "Eligible").reduce((s, d) => s + d.total_tax, 0))],
    ["BOE PDFs Parsed", boeDocs.length],
    ["BOE Matched Imports", boeDocs.filter((d) => d.matched).length],
    ["BOE GSTIN Mismatch", boeDocs.filter((d) => d.gstin_check !== "OK").length],
    ["BOE Not in 2B", boeDocs.filter((d) => d.twob_check === "Not in 2B").length],
    ["ISD GSTINs in ITC", isdGstinsInItc.size],
    ["ISD GSTINs Absent", isdGstinsAbsent.size],
  ];
  cards.forEach(([label, value], idx) => {
    const r = 4 + Math.floor(idx / 4) * 3;
    const c = 1 + (idx % 4) * 3;
    ws.getCell(r, c).value = label;
    ws.getCell(r + 1, c).value = value;
    ws.getCell(r, c).font = { bold: true, color: { argb: BLUE } };
    const isWarn =
      (label.includes("Correction") && value) ||
      (label.includes("Diff") && value) ||
      (label.includes("Absent") && value) ||
      (label.includes("Mismatch") && value) ||
      (label.includes("Not in 2B") && value);
    ws.getCell(r, c).fill = isWarn ? solid(WARN) : solid(LIGHT);
    ws.getCell(r + 1, c).font = { bold: true, size: 13 };
    ws.getCell(r + 1, c).numFmt = ["Rows", "Flags", "Parsed", "Matched", "in ITC", "Absent"].some((k) =>
      label.includes(k)
    )
      ? "#,##0"
      : "#,##0.00";
    ws.mergeCells(r, c, r, c + 1);
    ws.mergeCells(r + 1, c, r + 1, c + 1);
  });

  let cursor = 4 + Math.floor((cards.length - 1) / 4) * 3 + 4;

  const section = (title, fillColor = BLUE) => {
    ws.getCell(cursor, 1).value = title;
    ws.getCell(cursor, 1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    ws.getCell(cursor, 1).fill = solid(fillColor);
    ws.mergeCells(cursor, 1, cursor, 12);
    cursor += 1;
    return cursor;
  };

  const writeBlock = (headers, rows, emptyMsg, warnRows = false, headerColor = BLUE) => {
    headers.forEach((h, i) => {
      ws.getCell(cursor, i + 1).value = h;
    });
    styleHeader(ws, headers.length, cursor, headerColor);
    cursor += 1;
    if (!rows.length) {
      ws.getCell(cursor, 1).value = emptyMsg;
      cursor += 1;
      return;
    }
    for (const row of rows) {
      row.forEach((val, i) => {
        const cell = ws.getCell(cursor, i + 1);
        cell.value = val === undefined ? "" : val;
        if (typeof val === "number" && !Number.isNaN(val)) cell.numFmt = "#,##0.00";
        if (val instanceof Date) cell.numFmt = "dd/mm/yyyy";
        if (warnRows) cell.fill = solid(WARN);
      });
      cursor += 1;
    }
  };

  // Block: ledger mismatches at invoice level (this is where the 52,125 lives)
  cursor = section("Invoices where Sanitized ITC does not tie to GST REC ledger", BLUE);
  const ledgerAction = (diff, rcmDiff, reverse) => {
    if (Math.abs(diff) > 2 && reverse !== "Y") {
      return (
        "ITC booked in source but GST REC ledger movement missing/short. " +
        "Check SAP doc linkage (AP invoice vs GR/IR doc) and reference typos."
      );
    }
    if (Math.abs(rcmDiff) > 2 && reverse === "Y") {
      return "RCM expected from source but RCM payable ledger does not match — verify liability posting.";
    }
    return "Review ledger posting against source ITC classification.";
  };
  const mismatchRows = ledgerDetails
    .filter((d) => d[0] !== "OK")
    .map((d) => [
      d[1], d[2], d[3], d[6], d[7], d[8], d[4] || "", d[5] || "",
      d[13], d[17], d[22], d[21], d[24], d[25] || "", d[26] || "",
      ledgerAction(d[22], d[24], d[3]),
    ]);
  writeBlock(
    ["Source", "Type", "RCM", "Vch No.", "SAP Doc", "Date", "GSTIN", "Party",
      "ITC Tax", "GST REC Ledger", "ITC - REC Diff", "RCM Payable Ledger", "RCM Diff",
      "GL Docs", "GL Refs", "Suggested Action"],
    mismatchRows,
    "No invoice-level mismatch between Sanitized ITC and GST REC ledger.",
    true
  );
  cursor += 2;

  // Unresolved B2B
  cursor = section("Unresolved B2B bills to ask accountant", BLUE);
  writeBlock(
    ["Source", "Type", "Vch No.", "SAP Doc", "Date", "Party as per books",
      "Taxable", "IGST", "CGST", "SGST", "Total Tax", "Question for accountant"],
    unresolvedB2b.map((i) => [
      i.source, i.itc_type, i.invoice_no, i.sap_invoice_no, i.invoice_date, i.supplier_name,
      i.taxable, i.igst, i.cgst, i.sgst, i.total_tax,
      "Provide GSTIN/vendor master correction or confirm why this B2B ITC is not in current 2B.",
    ]),
    "No unresolved B2B bills after current 2B/source/prior matching.",
    true, RED
  );
  cursor += 2;

  // ISD GSTIN check
  cursor = section("ISD GSTIN check after Sanitized ITC", BLUE);
  const isdByGstin = new Map();
  for (const d of isdDocs) {
    if (!isdByGstin.has(d.gstin)) isdByGstin.set(d.gstin, []);
    isdByGstin.get(d.gstin).push(d);
  }
  const isdBlock = [];
  const sortedIsd = [...isdByGstin.entries()].sort(
    (a, b) => b[1].reduce((s, x) => s + x.total_tax, 0) - a[1].reduce((s, x) => s + x.total_tax, 0)
  );
  for (const [gstin, bills] of sortedIsd) {
    const largest = bills.reduce((m, b) => (b.total_tax > m.total_tax ? b : m), bills[0]);
    const matched = bills.filter((b) => b.matched);
    const sanitizedTax = bills[0].sanitized_gstin_tax;
    const isdTax = round2(bills.reduce((s, b) => s + b.total_tax, 0));
    const diff = round2(sanitizedTax - isdTax);
    let action;
    if (!bills[0].gstin_in_itc) {
      action = "Ask: ISD supplier GSTIN booked elsewhere (other GSTIN/branch) or not booked in 2081?";
    } else if (diff) {
      action = "GSTIN exists in Sanitized ITC, but ISD bill tax does not tie — check missing invoice numbers.";
    } else {
      action = "GSTIN-level ISD tax ties to Sanitized ITC.";
    }
    isdBlock.push([
      gstin, bills.length, isdTax, matched.length,
      round2(matched.reduce((s, b) => s + b.matched_books_tax, 0)),
      bills[0].gstin_in_itc ? "Present" : "Absent",
      sanitizedTax, diff,
      largest.doc_no, largest.total_tax, action,
    ]);
  }
  writeBlock(
    ["Supplier GSTIN", "ISD Bill Rows", "ISD Bill Tax", "Invoice No. Matches",
      "Matched Books Tax", "GSTIN Status", "Sanitized GSTIN Tax",
      "Sanitized - ISD Tax", "Largest ISD Bill", "Largest ISD Tax", "Action"],
    isdBlock,
    "No ISD bill-wise JSON selected.",
    false, PURPLE
  );
  cursor += 2;

  // BOE support
  cursor = section("Import BOE support", BLUE);
  const importRows = invoices.filter((i) => i.source === "IMPORT");
  const boeDashboardRows = importRows.map((i) => {
    const boe = i.support.boe_doc;
    return [
      i.invoice_no, i.sap_invoice_no,
      i.matched_doc ? i.matched_doc.doc_no : "",
      boe ? boe.be_no : "",
      boe ? boe.be_date : "",
      i.igst,
      boe ? boe.igst : "",
      boe ? boe.taxable_for_igst : "",
      boe ? boe.challan_no : "",
      boe ? "Matched" : "No BOE PDF support",
      boe ? boe.match_method : "",
      boe ? boe.file_name : "",
    ];
  });
  writeBlock(
    ["Import Ref", "SAP Doc", "2B BoE", "BOE PDF", "BOE Date",
      "Books IGST", "BOE IGST", "Taxable/IGST Base", "Challan No.", "Status", "Method", "File"],
    boeDashboardRows,
    "No import bills in sanitized ITC."
  );
  cursor += 2;

  // BOE validation flags
  cursor = section("BOE validation flags (GSTIN check + GSTR-2B reporting check)", BLUE);
  const flaggedBoes = boeDocs.filter((d) => d.gstin_check !== "OK" || d.twob_check === "Not in 2B");
  writeBlock(
    ["BOE No.", "BOE Date", "Port", "Importer GSTIN on BOE", "Recipient GSTIN",
      "GSTIN Check", "2B Check", "Matched SAP Doc", "BOE IGST", "Suggested Action"],
    flaggedBoes.map((d) => [
      d.be_no, d.be_date, d.port_code, d.importer_gstin, RECIPIENT_GSTIN,
      d.gstin_check, d.twob_check, d.matched_source_doc, d.igst, d.validation_action,
    ]),
    "All BOEs pass GSTIN and 2B reporting checks.",
    true, RED
  );
  cursor += 2;

  // Category-wise ledger tie
  cursor = section("Category-wise ITC vs SAP GST ledger movement", BLUE);
  const catHeaders = ["Status", "Source", "Type", "RCM", "Rows", "ITC Tax", "GST REC Ledger",
    "RCM Payable Ledger", "ITC - REC Diff", "RCM Diff"];
  catHeaders.forEach((h, i) => {
    ws.getCell(cursor, i + 1).value = h;
  });
  styleHeader(ws, catHeaders.length, cursor);
  cursor += 1;
  for (const item of ledgerSummary) {
    const values = [item[0], item[1], item[2], item[3], item[4], item[9], item[13], item[17], item[18], item[19]];
    values.forEach((v, i) => {
      const cell = ws.getCell(cursor, i + 1);
      cell.value = v;
      if (typeof v === "number") cell.numFmt = "#,##0.00";
      if (item[0] !== "OK") cell.fill = solid(WARN);
    });
    cursor += 1;
  }
  cursor += 2;

  // Correction flags
  cursor = section("Immediate book correction flags", RED);
  const corrHeaders = ["Flag", "Source", "Type", "Invoice", "GSTIN", "Party", "Amount", "Required Action"];
  corrHeaders.forEach((h, i) => {
    ws.getCell(cursor, i + 1).value = h;
  });
  styleHeader(ws, corrHeaders.length, cursor, RED);
  cursor += 1;
  if (corrections.length) {
    for (const item of corrections.slice(0, 20)) {
      const values = [item[0], item[1], item[2], item[3], item[5], item[6], item[9], item[10]];
      values.forEach((v, i) => {
        const cell = ws.getCell(cursor, i + 1);
        cell.value = v;
        cell.fill = solid(WARN);
        if (typeof v === "number") cell.numFmt = "#,##0.00";
      });
      cursor += 1;
    }
  } else {
    ws.getCell(cursor, 1).value = "No immediate book correction flags after sanitization.";
    cursor += 1;
  }

  setWidths(ws, 34);
  ws.eachRow((row) => {
    row.eachCell((cell) => {
      cell.alignment = { ...(cell.alignment || {}), vertical: "top", wrapText: true };
    });
  });
}

export async function assembleWorkbook(ctx) {
  const {
    invoices, r2bDocs, ledgerSummary, ledgerDetails, corrections,
    boeDocs, isdDocs, isdDist, isdEligibilitySummary,
    periodLabel, fyLabel,
    itcRowsData, r2bRowsData, isdRowsData, isdEligibleRowsData, boeRowsData,
    twoBSummaryRows, twoBDetailRows,
  } = ctx;

  const wb = new ExcelJS.Workbook();
  buildDashboard(wb, {
    invoices, docs: r2bDocs, ledgerSummary, ledgerDetails, corrections,
    boeDocs, isdDocs, isdDist, periodLabel, fyLabel,
  });

  // Sanitized ITC
  const itcHeaders = ["Party GSTIN/UIN", "Party name", "Vch No.", "Date",
    "Taxable", "IGST", "CGST", "SGST", "Tax",
    "Place of Supply", "Reverse Charge", "ITC Availability", "Type",
    "3B Month", "Books Month", "Invoice Match", "GSTIN Match", "FY",
    "Support Reference",
    "ISD Flag", "ISD Supplier GSTIN", "ISD Bill Tax", "Books Tax - ISD Bill Tax"];
  let ws = wb.addWorksheet("Sanitized ITC");
  writeSheet(ws, itcHeaders, itcRowsData, "SanitizedITC");
  formatNumbers(ws, [5, 6, 7, 8, 9, 22, 23], [4]);

  // 2B
  const r2bHeaders = ["Party GSTIN/UIN", "Particulars", "Vch No.", "Date",
    "Taxable", "IGST", "CGST", "SGST", "Tax",
    "Place of Supply", "Reverse Charge", "ITC Availability", "Type",
    "Month", "Invoice Match", "GSTIN Match", "FY"];
  ws = wb.addWorksheet("2B");
  writeSheet(ws, r2bHeaders, r2bRowsData, "GSTR2B");
  formatNumbers(ws, [5, 6, 7, 8, 9], [4]);

  // ISD — same 17 columns as 2B + 6 ISD-specific tail columns + Eligibility
  const isdHeaders = [...r2bHeaders,
    "Invoice Value", "Rate(s)",
    "Matched SAP/Books Doc", "Matched Books Tax", "Books Tax - ISD Bill Tax",
    "Sanitized GSTIN Tax", "Eligibility"];
  ws = wb.addWorksheet("ISD");
  writeSheet(ws, isdHeaders, isdRowsData, "ISDBillDetail");
  formatNumbers(ws, [5, 6, 7, 8, 9, 18, 21, 22, 23], [4]);

  // ISD Eligible — bill-wise rows that constitute the eligible ISD distribution only
  if (isdDocs.length) {
    ws = wb.addWorksheet("ISD Eligible");
    writeSheet(ws, isdHeaders, isdEligibleRowsData, "ISDEligibleBills");
    formatNumbers(ws, [5, 6, 7, 8, 9, 18, 21, 22, 23], [4]);
    const eligibleBills = isdDocs.filter((d) => d.eligibility === "Eligible");
    const last = ws.rowCount + 2;
    ws.getCell(last, 1).value = "Classification method";
    ws.getCell(last, 1).font = { bold: true };
    ws.getCell(last, 2).value = isdEligibilitySummary.method;
    ws.getCell(last + 1, 1).value = "Eligible supplier GSTINs";
    ws.getCell(last + 1, 2).value = isdEligibilitySummary.eligible_gstins.join(", ") || "(none)";
    const eligTotal = round2(eligibleBills.reduce((s, d) => s + d.total_tax, 0));
    const target = isdEligibilitySummary.elig_total_tax;
    ws.getCell(last + 2, 1).value = "Eligible bill tax (this sheet)";
    ws.getCell(last + 2, 2).value = eligTotal;
    ws.getCell(last + 2, 2).numFmt = "#,##0.00";
    ws.getCell(last + 3, 1).value = "ISD-002 eligible distribution";
    ws.getCell(last + 3, 2).value = target;
    ws.getCell(last + 3, 2).numFmt = "#,##0.00";
    ws.getCell(last + 4, 1).value = "Tie-out diff (bills - distribution)";
    ws.getCell(last + 4, 2).value = round2(eligTotal - target);
    ws.getCell(last + 4, 2).numFmt = "#,##0.00";
    if (isdEligibilitySummary.warning) {
      ws.getCell(last + 5, 1).value = "Warning";
      ws.getCell(last + 5, 2).value = isdEligibilitySummary.warning;
      ws.getCell(last + 5, 2).fill = solid(WARN);
    }
  }

  // BOE Detail (only if BOE input present)
  if (boeDocs.length) {
    const boeHeaders = ["Status", "BOE No.", "BOE Date", "Port Code", "Importer GSTIN", "Importer",
      "Invoice No.", "Invoice Amount", "Currency",
      "Assessable Value", "BCD", "SWS", "Taxable/IGST Base", "IGST",
      "Total Duty", "Total Amount", "Challan No.", "Payment Amount",
      "Matched SAP Doc", "Matched ITC IGST", "ITC IGST - BOE IGST",
      "Match Method",
      "GSTIN Check", "2B Check", "Validation Action",
      "File", "Path"];
    ws = wb.addWorksheet("BOE Detail");
    writeSheet(ws, boeHeaders, boeRowsData, "BOEDetail");
    formatNumbers(ws, [8, 10, 11, 12, 13, 14, 15, 16, 18, 20, 21], [3]);
    boeDocs.forEach((d, i) => {
      if (d.gstin_check !== "OK" || d.twob_check !== "Reported in 2B") {
        for (let c = 1; c <= boeHeaders.length; c++) {
          ws.getCell(i + 2, c).fill = solid(WARN);
        }
      }
    });
  }

  // 2B vs Books Summary
  const summaryHeaders = ["Status", "GSTIN", "Party", "Type", "Reverse Charge",
    "Books Invoice Count", "2B Invoice Count",
    "Books Taxable", "Books IGST", "Books CGST", "Books SGST", "Books Total Tax",
    "2B Taxable", "2B IGST", "2B CGST", "2B SGST", "2B Total Tax",
    "Books - 2B Difference", "Suggested Action"];
  ws = wb.addWorksheet("2B vs Books Summary");
  writeSheet(ws, summaryHeaders, twoBSummaryRows, "TwoBBooksSummary");
  formatNumbers(ws, Array.from({ length: 11 }, (_, i) => i + 8));

  // 2B vs Books Detail
  const detailHeaders = ["Status", "Source", "Type", "GSTIN", "Party",
    "Books Vch No.", "Books Date",
    "Books Taxable", "Books IGST", "Books CGST", "Books SGST", "Books Total Tax",
    "2B Section", "2B Doc No.", "2B Date",
    "2B Taxable", "2B IGST", "2B CGST", "2B SGST", "2B Total Tax",
    "Taxable Diff", "Tax Diff", "Remarks",
    "Book Correction Flag", "Correction Amount", "Correction Action"];
  ws = wb.addWorksheet("2B vs Books Detail");
  writeSheet(ws, detailHeaders, twoBDetailRows, "TwoBBooksDetail");
  formatNumbers(ws, [8, 9, 10, 11, 12, 16, 17, 18, 19, 20, 21, 22, 25], [7, 15]);

  // ITC vs Ledger (category summary)
  const ledgerSummaryHeaders = ["Status", "Source", "Type", "RCM", "Invoice Count",
    "Taxable", "ITC IGST", "ITC CGST", "ITC SGST", "ITC Total",
    "Ledger IGST REC", "Ledger CGST REC", "Ledger SGST REC", "Ledger REC Total",
    "Ledger RCM IGST", "Ledger RCM CGST", "Ledger RCM SGST", "Ledger RCM Total",
    "ITC - REC Diff", "Expected RCM - Ledger RCM Diff", "Invoice Exceptions"];
  ws = wb.addWorksheet("ITC vs Ledger");
  writeSheet(ws, ledgerSummaryHeaders, ledgerSummary, "ITCVsLedger");
  formatNumbers(ws, Array.from({ length: 15 }, (_, i) => i + 6));

  // Invoice Ledger Detail
  const ledgerDetailHeaders = ["Status", "Source", "Type", "RCM", "GSTIN", "Party",
    "Vch No.", "SAP Doc", "Date",
    "Taxable", "ITC IGST", "ITC CGST", "ITC SGST", "ITC Total",
    "Ledger IGST REC", "Ledger CGST REC", "Ledger SGST REC", "Ledger REC Total",
    "Ledger RCM IGST", "Ledger RCM CGST", "Ledger RCM SGST", "Ledger RCM Total",
    "ITC - REC Diff", "Expected RCM", "Expected RCM - Ledger RCM Diff",
    "GL Docs", "GL References"];
  ws = wb.addWorksheet("Invoice Ledger Detail");
  writeSheet(ws, ledgerDetailHeaders, ledgerDetails, "InvoiceLedgerDetail");
  formatNumbers(ws, Array.from({ length: 16 }, (_, i) => i + 10), [9]);

  // Book Correction Flags
  const correctionHeaders = ["Flag", "Source", "Type", "Vch No.", "SAP Doc", "GSTIN", "Party",
    "Invoice Match", "Tax Difference vs 2B",
    "Correction Amount", "Required Action", "Notes"];
  ws = wb.addWorksheet("Book Correction Flags");
  writeSheet(ws, correctionHeaders, corrections, "BookCorrectionFlags");
  formatNumbers(ws, [9, 10]);

  // Match openpyxl pass: body cells top-aligned, no wrap
  for (const sheet of wb.worksheets) {
    if (sheet.name === "Dashboard") continue;
    for (let r = 2; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      row.eachCell((cell) => {
        cell.alignment = { vertical: "top", wrapText: false };
      });
    }
  }

  const buffer = await wb.xlsx.writeBuffer();
  return buffer;
}

export function outputFileName(periodLabel, fyLabel) {
  return `SIPL_GST_ITC_${periodLabel.replace(/ /g, "_")}_${fyLabel.replace(/-/g, "_")}_Focused_Output.xlsx`;
}
