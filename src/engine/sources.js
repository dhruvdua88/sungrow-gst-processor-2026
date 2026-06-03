// Source-row aggregation (PO + Non PO + Import) — mirrors aggregate_source_rows
import {
  asText, asFloat, normDoc, cleanName, parseDate, taxTotal, round2,
  TARGET_COMPANY_CODE, RCM_ACCOUNTS_LEDGER,
} from "./helpers.js";
import { readSheetRows } from "./readers.js";

export function makeSourceInvoice(fields) {
  return {
    source: "", company_code: "", source_docs: new Set(),
    invoice_no: "", sap_invoice_no: "", invoice_date: null,
    supplier_name: "", country: "", place_of_supply: "", source_gstin: "",
    taxable: 0, igst: 0, cgst: 0, sgst: 0, total_tax: 0,
    source_rcm: false, source_rcm_flag: false, ledger_rcm_evidence: false,
    source_rcm_tax: 0,
    matched_doc: null, resolved_gstin: "", resolved_party: "",
    itc_available: "Y", reverse_charge: "N", itc_type: "B2B",
    invoice_match: "Only in Books", gstin_match: "Unresolved",
    mapping_method: "", mapping_confidence: 0, mapping_notes: "",
    tax_diff_vs_2b: null, taxable_diff_vs_2b: null,
    correction_flag: "", correction_action: "", correction_amount: 0,
    support: {},
    ...fields,
  };
}

function buildDocDateLookup(glRows) {
  const lookup = new Map();
  for (const row of glRows) {
    const doc = normDoc(row["Document Number"]);
    if (!doc) continue;
    const date = parseDate(row["Document Date"]) || parseDate(row["Posting Date"]);
    if (date && !lookup.has(doc)) lookup.set(doc, date);
  }
  return lookup;
}

function buildReverseDocSet(glRows) {
  const reverse = new Set();
  for (const row of glRows) {
    const account = asText(row["Account"]);
    const name = asText(row["G/L Name"]).toUpperCase();
    const taxCode = asText(row["Tax code"]).toUpperCase();
    if (name.includes("REVE CH") || RCM_ACCOUNTS_LEDGER.has(account) || taxCode === "3R") {
      for (const k of [
        normDoc(row["Document Number"]),
        normDoc(row["Reference"]),
        normDoc(row["Invoice reference"]),
      ]) {
        if (k) reverse.add(k);
      }
    }
  }
  return reverse;
}

export function aggregateSourceRows(glRows, gstWorkbook) {
  const glDate = buildDocDateLookup(glRows);
  const reverseDocs = buildReverseDocSet(glRows);
  const invoices = new Map(); // key string -> SourceInvoice

  const TRUTHY_RCM = new Set(["Y", "YES", "TRUE", "1"]);

  // --- PO sheet ---
  for (const row of readSheetRows(gstWorkbook, "PO")) {
    if (asText(row["Supplier Name"]).toUpperCase() === "NA") continue;
    if (!asText(row["SAP Invoice No."]) && !asText(row["Original Invoice No."])) continue;
    const tax = taxTotal(row["IGST Amount"], row["CGST Amount"], row["SGST Amount"]);
    if (Math.abs(tax) < 0.005) continue;
    const invoiceNo = asText(row["Original Invoice No."]) || asText(row["SAP Invoice No."]);
    const gstin = asText(row["GST No."]).toUpperCase();
    const supplier = asText(row["Supplier Name"]);
    const rcm = TRUTHY_RCM.has(asText(row["RCM"]).toUpperCase());
    const ledgerRcm =
      reverseDocs.has(normDoc(invoiceNo)) || reverseDocs.has(normDoc(row["SAP Invoice No."]));
    const key = JSON.stringify(["PO", normDoc(invoiceNo), gstin, cleanName(supplier)]);
    let item = invoices.get(key);
    if (!item) {
      item = makeSourceInvoice({
        source: "PO", company_code: TARGET_COMPANY_CODE, source_docs: new Set(),
        invoice_no: invoiceNo, sap_invoice_no: asText(row["SAP Invoice No."]),
        invoice_date: parseDate(row["SAP Invoice Date"]),
        supplier_name: supplier, country: asText(row["Country"]),
        place_of_supply: asText(row["Description"]),
        source_gstin: gstin,
        source_rcm: rcm || ledgerRcm, source_rcm_flag: rcm, ledger_rcm_evidence: ledgerRcm,
      });
      invoices.set(key, item);
    }
    item.source_docs.add(asText(row["SAP Invoice No."]));
    item.source_rcm = item.source_rcm || rcm || ledgerRcm;
    item.source_rcm_flag = item.source_rcm_flag || rcm;
    item.ledger_rcm_evidence = item.ledger_rcm_evidence || ledgerRcm;
    if (rcm) item.source_rcm_tax += tax;
    item.taxable += asFloat(row["Basic Amount"]);
    item.igst += asFloat(row["IGST Amount"]);
    item.cgst += asFloat(row["CGST Amount"]);
    item.sgst += asFloat(row["SGST Amount"]);
    item.total_tax = taxTotal(item.igst, item.cgst, item.sgst);
  }

  // --- Non PO sheet ---
  for (const row of readSheetRows(gstWorkbook, "Non PO")) {
    const company = asText(row["Company Code"]);
    if (company && company !== TARGET_COMPANY_CODE) continue;
    if (asText(row["Supplier Name"]).toUpperCase() === "NA") continue;
    if (!asText(row["Document No"])) continue;
    const tax = taxTotal(row["IGST Amount"], row["CGST Amount"], row["SGST Amount"]);
    if (Math.abs(tax) < 0.005) continue;
    const docNo = asText(row["Document No"]);
    const ref = asText(row["Reference"]) || docNo;
    const supplier = asText(row["Supplier Name"]);
    const country = asText(row["Country"]);
    const gstin = asText(row["GST No."]).toUpperCase();
    const docNorm = normDoc(docNo);
    const ledgerRcm = reverseDocs.has(docNorm) || reverseDocs.has(normDoc(ref));
    const isImport = country.toUpperCase() !== "IN";
    const key = isImport
      ? JSON.stringify(["IMPORT", docNorm, ref, round2(tax)])
      : JSON.stringify(["Non PO", docNorm, normDoc(ref), cleanName(supplier), gstin, ledgerRcm]);
    let item = invoices.get(key);
    if (!item) {
      item = makeSourceInvoice({
        source: isImport ? "IMPORT" : "Non PO",
        company_code: TARGET_COMPANY_CODE, source_docs: new Set(),
        invoice_no: ref, sap_invoice_no: docNo,
        invoice_date: glDate.get(docNorm) || null,
        supplier_name: supplier, country,
        place_of_supply: asText(row["State Name"]),
        source_gstin: gstin,
        source_rcm: ledgerRcm, source_rcm_flag: ledgerRcm, ledger_rcm_evidence: ledgerRcm,
      });
      invoices.set(key, item);
    }
    item.source_docs.add(docNo);
    item.source_rcm = item.source_rcm || ledgerRcm;
    item.source_rcm_flag = item.source_rcm_flag || ledgerRcm;
    item.ledger_rcm_evidence = item.ledger_rcm_evidence || ledgerRcm;
    if (ledgerRcm) item.source_rcm_tax += tax;
    item.taxable += asFloat(row["Basic Amount"]);
    item.igst += asFloat(row["IGST Amount"]);
    item.cgst += asFloat(row["CGST Amount"]);
    item.sgst += asFloat(row["SGST Amount"]);
    item.total_tax = taxTotal(item.igst, item.cgst, item.sgst);
  }

  return Array.from(invoices.values());
}
