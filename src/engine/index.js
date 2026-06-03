// Main entry point — mirrors build_workbook(), browser-native:
// File inputs in, styled XLSX ArrayBuffer + run summary out.
import { asFloat, round2 } from "./helpers.js";
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
  const invoices = aggregateSourceRows(glRows, gstWorkbook);

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

  const isdGstins = new Set(isdDocs.filter((d) => d.gstin).map((d) => d.gstin));
  const isdGstinsInItc = new Set(isdDocs.filter((d) => d.gstin && d.gstin_in_itc).map((d) => d.gstin));
  const isdGstinsAbsent = new Set([...isdGstins].filter((g) => !isdGstinsInItc.has(g)));

  return {
    buffer,
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
