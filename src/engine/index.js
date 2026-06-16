// Main entry point — mirrors build_workbook(), browser-native:
// File inputs in, styled XLSX ArrayBuffer + run summary out.
import { asFloat, round2, fmtDate } from "./helpers.js";
import { workbookFromArrayBuffer, readGlRows } from "./readers.js";
import { loadR2bDocs } from "./r2b.js";
import { aggregateSourceRows } from "./sources.js";
import { resolveAll } from "./resolve.js";
import { loadPriorAliases, applyPriorAliases } from "./prior.js";
import { loadBoeDocs, validateBoeDocs, applyBoeMatches } from "./boe.js";
import {
  loadIsdDistributions, loadIsdBillDocs, classifyIsdBillEligibility, applyIsdMatches,
} from "./isd.js";
import { buildLedgerTie } from "./ledger.js";
import {
  build2bBooksSummary, build2bBooksDetail, buildCorrectionRows,
  itcRows, r2bRows, isdRows, boeRows,
} from "./recon.js";
import { assembleWorkbook, outputFileName } from "./workbook.js";

function counter(items, keyFn) {
  const c = {};
  for (const it of items) {
    const k = keyFn(it);
    c[k] = (c[k] || 0) + 1;
  }
  return c;
}

/**
 * @param {Object} inputs
 * @param {ArrayBuffer} inputs.trackerBuffer     SIPL tracker/template xlsx (validated present, like the Python tool)
 * @param {ArrayBuffer} inputs.gstWorkbookBuffer Current GST workbook xlsx (PO / Non PO / GL Summary)
 * @param {Object}      inputs.r2bJson           Parsed GSTR-2B JSON
 * @param {Array}       inputs.priorFiles        Optional [{name, buffer}] prior Focused_Output xlsx files
 * @param {Array}       inputs.boeFiles          Optional [{name, data, relativePath}] BOE zip/pdf inputs
 * @param {Object|null} inputs.isdJson           Optional parsed ISD JSON
 * @param {string}      inputs.periodLabel
 * @param {string}      inputs.fyLabel
 * @param {Function}    inputs.onProgress        Status logger
 */
export async function buildWorkbook(inputs) {
  const {
    trackerBuffer, gstWorkbookBuffer, r2bJson,
    priorFiles = [], boeFiles = [], isdJson = null,
    periodLabel = "April", fyLabel = "2026-27",
    onProgress = () => {},
  } = inputs;

  if (!trackerBuffer) throw new Error("SIPL tracker/template not found.");
  if (!gstWorkbookBuffer) throw new Error("GST workbook not found.");
  if (!r2bJson) throw new Error("GSTR-2B JSON not found.");

  onProgress("Reading GST workbook (PO / Non PO / GL Summary)...");
  const gstWorkbook = workbookFromArrayBuffer(gstWorkbookBuffer);
  const glRows = readGlRows(gstWorkbook);

  onProgress("Loading GSTR-2B JSON...");
  const r2bDocs = loadR2bDocs(r2bJson);

  onProgress("Scanning BOE input...");
  const [boeDocs, boeInspection] = await loadBoeDocs(boeFiles, onProgress);
  const boeValidation = validateBoeDocs(boeDocs, r2bJson);

  const isdDocs = loadIsdBillDocs(isdJson);
  const isdDist = loadIsdDistributions(isdJson);

  onProgress("Aggregating source invoices (PO + Non PO + Import)...");
  const warnings = [];
  const invoices = aggregateSourceRows(glRows, gstWorkbook, warnings);
  for (const w of warnings) onProgress(`⚠ ${w.sheet}: ${w.title}`);

  onProgress("Resolving invoices against 2B...");
  resolveAll(invoices, r2bDocs);

  onProgress("Applying prior focused-file aliases...");
  const priorWbs = priorFiles.map((f) => ({ name: f.name, workbook: workbookFromArrayBuffer(f.buffer) }));
  const prior = loadPriorAliases(priorWbs);
  applyPriorAliases(invoices, r2bDocs, prior);

  applyBoeMatches(invoices, boeDocs);
  applyIsdMatches(invoices, isdDocs);
  const isdEligibilitySummary = classifyIsdBillEligibility(isdDocs, isdDist);

  onProgress("Building ledger tie-out...");
  const [ledgerSummary, ledgerDetails] = buildLedgerTie(invoices, glRows);
  const corrections = buildCorrectionRows(invoices);
  const nonIsd = invoices.filter((i) => i.support.isd_flag !== "ISD");

  onProgress("Assembling focused workbook...");
  const eligibleBills = isdDocs.filter((d) => d.eligibility === "Eligible");
  const buffer = await assembleWorkbook({
    invoices, r2bDocs, ledgerSummary, ledgerDetails, corrections,
    boeDocs, isdDocs, isdDist, isdEligibilitySummary,
    periodLabel, fyLabel,
    itcRowsData: itcRows(invoices, periodLabel, fyLabel),
    r2bRowsData: r2bRows(r2bDocs, periodLabel, fyLabel),
    isdRowsData: isdRows(isdDocs, periodLabel, fyLabel),
    isdEligibleRowsData: isdRows(eligibleBills, periodLabel, fyLabel),
    boeRowsData: boeRows(boeDocs),
    twoBSummaryRows: build2bBooksSummary(nonIsd, r2bDocs),
    twoBDetailRows: build2bBooksDetail(nonIsd, r2bDocs),
  });

  // ---- In-browser preview data (mirrors the Dashboard sheet blocks) ----
  const booksTax = round2(invoices.reduce((s, i) => s + i.total_tax, 0));
  const booksNonIsd = round2(nonIsd.reduce((s, i) => s + i.total_tax, 0));
  const r2bTaxAvail = round2(r2bDocs.filter((d) => d.itc_available === "Y").reduce((s, d) => s + d.total_tax, 0));
  const ledgerRec = round2(ledgerSummary.reduce((s, r) => s + r[13], 0));
  const ledgerRcm = round2(ledgerSummary.reduce((s, r) => s + r[17], 0));
  const corrAmt = round2(corrections.reduce((s, r) => s + asFloat(r[9]), 0));
  const unresolvedB2b = invoices.filter((i) => i.itc_type === "B2B" && i.gstin_match === "Unresolved");
  const rcmUr = invoices.filter((i) => i.itc_type === "RCM-UR");
  const isdEligibleDist = isdDist.filter((d) => d.eligibility === "Eligible");
  const isdEligibleTax = round2(isdEligibleDist.reduce((s, d) => s + d.total_tax, 0));
  const isdBillTax = round2(isdDocs.reduce((s, d) => s + d.total_tax, 0));
  const isdInItc = new Set(isdDocs.filter((d) => d.gstin_in_itc).map((d) => d.gstin));
  const isdAbsent = new Set(isdDocs.filter((d) => !d.gstin_in_itc).map((d) => d.gstin));

  const countCard = (label, value) => ({ label, value, count: true });
  const amtCard = (label, value) => ({ label, value, count: false });
  const previewCards = [
    amtCard("Sanitized ITC Tax", booksTax),
    amtCard("GST REC Ledger Movement", ledgerRec),
    amtCard("ITC - Ledger Difference", round2(booksTax - ledgerRec)),
    amtCard("2B Available Tax", r2bTaxAvail),
    amtCard("Books Non-ISD - 2B Diff", round2(booksNonIsd - r2bTaxAvail)),
    amtCard("RCM Payable Movement", ledgerRcm),
    countCard("Immediate Correction Flags", corrections.length),
    amtCard("Correction Amount", corrAmt),
    countCard("Unresolved B2B Rows", unresolvedB2b.length),
    amtCard("Unresolved B2B Tax", round2(unresolvedB2b.reduce((s, i) => s + i.total_tax, 0))),
    countCard("RCM-UR Rows", rcmUr.length),
    amtCard("RCM-UR Tax", round2(rcmUr.reduce((s, i) => s + i.total_tax, 0))),
    countCard("ISD Bill Rows", isdDocs.length),
    amtCard("ISD Bill Tax", isdBillTax),
    amtCard("Eligible ISD Tax (ISD-002)", isdEligibleTax),
    amtCard("ISD Bill - Eligible Diff", isdDocs.length || isdEligibleDist.length ? round2(isdBillTax - isdEligibleTax) : 0),
    countCard("BOE PDFs Parsed", boeDocs.length),
    countCard("BOE Matched Imports", boeDocs.filter((d) => d.matched).length),
    countCard("BOE GSTIN Mismatch", boeDocs.filter((d) => d.gstin_check !== "OK").length),
    countCard("BOE Not in 2B", boeDocs.filter((d) => d.twob_check === "Not in 2B").length),
    countCard("ISD GSTINs in ITC", isdInItc.size),
    countCard("ISD GSTINs Absent", isdAbsent.size),
  ].map((c) => ({
    ...c,
    warn: Boolean(
      ((c.label.includes("Correction") || c.label.includes("Diff") || c.label.includes("Absent") ||
        c.label.includes("Mismatch") || c.label.includes("Not in 2B") || c.label.includes("Unresolved")) && c.value)
    ),
  }));

  const ledgerMismatchRows = ledgerDetails
    .filter((d) => d[0] !== "OK")
    .map((d) => [
      d[1], d[2], d[3], d[6], d[7], fmtDate(d[8]), d[4] || "", d[5] || "",
      d[13], d[17], d[22], d[21], d[24],
    ]);

  const preview = {
    cards: previewCards,
    ledgerMismatches: {
      title: "Invoices where Sanitized ITC does not tie to GST REC ledger",
      headers: ["Source", "Type", "RCM", "Vch No.", "SAP Doc", "Date", "GSTIN", "Party",
        "ITC Tax", "GST REC Ledger", "ITC - REC Diff", "RCM Payable", "RCM Diff"],
      rows: ledgerMismatchRows,
      empty: "No invoice-level mismatch between Sanitized ITC and GST REC ledger.",
    },
    unresolvedB2b: {
      title: "Unresolved B2B bills to ask accountant",
      headers: ["Source", "Type", "Vch No.", "SAP Doc", "Date", "Party as per books",
        "Taxable", "IGST", "CGST", "SGST", "Total Tax"],
      rows: unresolvedB2b.map((i) => [
        i.source, i.itc_type, i.invoice_no, i.sap_invoice_no, fmtDate(i.invoice_date),
        i.supplier_name, round2(i.taxable), round2(i.igst), round2(i.cgst), round2(i.sgst),
        round2(i.total_tax),
      ]),
      empty: "No unresolved B2B bills after current 2B/source/prior matching.",
    },
    corrections: {
      title: "Immediate book correction flags",
      headers: ["Flag", "Source", "Type", "Vch No.", "SAP Doc", "GSTIN", "Party",
        "Invoice Match", "Tax Diff vs 2B", "Correction Amount", "Required Action"],
      rows: corrections.map((r) => [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10]]),
      empty: "No immediate book correction flags after sanitization.",
    },
    ledgerSummary: {
      title: "Category-wise ITC vs SAP GST ledger movement",
      headers: ["Status", "Source", "Type", "RCM", "Rows", "ITC Tax", "GST REC Ledger",
        "RCM Payable Ledger", "ITC - REC Diff", "RCM Diff"],
      rows: ledgerSummary.map((r) => [r[0], r[1], r[2], r[3], r[4], r[9], r[13], r[17], r[18], r[19]]),
      empty: "No source invoices.",
    },
  };

  const isdGstins = new Set(isdDocs.filter((d) => d.gstin).map((d) => d.gstin));
  const isdGstinsInItc = new Set(isdDocs.filter((d) => d.gstin && d.gstin_in_itc).map((d) => d.gstin));
  const isdGstinsAbsent = new Set([...isdGstins].filter((g) => !isdGstinsInItc.has(g)));

  return {
    buffer,
    preview,
    warnings,
    output: outputFileName(periodLabel, fyLabel),
    invoice_count: invoices.length,
    r2b_count: r2bDocs.length,
    source_counts: counter(invoices, (i) => i.source),
    match_counts: counter(invoices, (i) => i.invoice_match),
    gstin_counts: counter(invoices, (i) => i.gstin_match),
    unresolved: invoices.filter((i) => i.gstin_match === "Unresolved"),
    books_total_tax: round2(invoices.reduce((s, i) => s + i.total_tax, 0)),
    r2b_total_tax: round2(r2bDocs.filter((d) => d.itc_available === "Y").reduce((s, d) => s + d.total_tax, 0)),
    import_unmatched: invoices.filter((i) => i.source === "IMPORT" && i.invoice_match !== "Match"),
    ledger_summary: ledgerSummary,
    ledger_mismatch_count: ledgerSummary.reduce((s, r) => s + asFloat(r[20]), 0),
    ledger_itc_rec_difference: round2(ledgerSummary.reduce((s, r) => s + asFloat(r[18]), 0)),
    correction_count: corrections.length,
    correction_amount: round2(corrections.reduce((s, r) => s + asFloat(r[9]), 0)),
    boe_count: boeDocs.length,
    boe_matched_count: boeDocs.filter((d) => d.matched).length,
    boe_inspection: boeInspection,
    boe_validation: boeValidation,
    boe_gstin_mismatch_count: boeValidation.gstin_mismatch + boeValidation.gstin_missing,
    boe_not_in_2b_count: boeValidation.not_in_2b,
    isd_bill_count: isdDocs.length,
    isd_bill_tax: round2(isdDocs.reduce((s, d) => s + d.total_tax, 0)),
    isd_eligible_count: isdDist.filter((d) => d.eligibility === "Eligible").length,
    isd_eligible_tax: round2(isdDist.filter((d) => d.eligibility === "Eligible").reduce((s, d) => s + d.total_tax, 0)),
    isd_gstin_present_count: isdGstinsInItc.size,
    isd_gstin_absent_count: isdGstinsAbsent.size,
    isd_gstin_absent_tax: round2(
      isdDocs.filter((d) => isdGstinsAbsent.has(d.gstin)).reduce((s, d) => s + d.total_tax, 0)
    ),
    isd_eligible_bill_count: isdDocs.filter((d) => d.eligibility === "Eligible").length,
    isd_eligible_bill_tax: round2(
      isdDocs.filter((d) => d.eligibility === "Eligible").reduce((s, d) => s + d.total_tax, 0)
    ),
    isd_eligibility_method: isdEligibilitySummary.method,
    isd_eligible_supplier_gstins: isdEligibilitySummary.eligible_gstins,
    prior_alias_count: prior.aliases.size,
    prior_conflict_count: prior.conflicts.size,
  };
}
