// 2B vs Books reconciliation + row builders — mirrors _suggested_action /
// build_2b_books_summary / build_2b_books_detail / build_correction_rows /
// _r2b_type / itc_rows / r2b_rows / isd_rows / boe_rows
import { asText, round2 } from "./helpers.js";

export function r2bType(doc) {
  if (doc.section === "IMPG") return "IMPORTGOODS";
  if (doc.section === "ISD") return "ISD";
  if (doc.reverse_charge === "Y") return "RCM-B2B";
  if (doc.section === "CDNR") return "B2B-CDNR";
  return "B2B";
}

function suggestedAction(status, difference, reverse) {
  if (status === "Match" && Math.abs(difference) <= 2) return "No action.";
  if (status === "Only in Books") return "Check vendor/BoE availability in 2B before claiming.";
  if (status === "Only in 2B") return "Check whether invoice is missing in books or should be deferred.";
  if (status === "GSTIN mismatch") return "Correct GSTIN master/books or verify 2B supplier GSTIN.";
  if (status === "Ambiguous prior match") return "Review prior alias conflict before auto-mapping GSTIN.";
  if (status === "Tax mismatch") return "Correct books immediately before 3B claim.";
  if (reverse === "Y") return "Review RCM classification and liability posting.";
  return "Review difference.";
}

export function build2bBooksSummary(invoices, docs) {
  const buckets = new Map();

  const get = (gstin, party, reverse, typ) => {
    const k = JSON.stringify([gstin, party, reverse, typ]);
    if (!buckets.has(k)) {
      buckets.set(k, {
        gstin, party, reverse, type: typ,
        books_taxable: 0, books_igst: 0, books_cgst: 0, books_sgst: 0,
        books_tax: 0, books_count: 0,
        r2b_taxable: 0, r2b_igst: 0, r2b_cgst: 0, r2b_sgst: 0,
        r2b_tax: 0, r2b_count: 0, statuses: new Map(),
      });
    }
    return buckets.get(k);
  };
  const bump = (counter, key) => counter.set(key, (counter.get(key) || 0) + 1);

  for (const inv of invoices) {
    let status = inv.invoice_match;
    if (inv.gstin_match === "GSTIN mismatch") status = "GSTIN mismatch";
    else if (inv.gstin_match === "Ambiguous prior match") status = "Ambiguous prior match";
    const b = get(inv.resolved_gstin, inv.resolved_party || inv.supplier_name, inv.reverse_charge, inv.itc_type);
    b.books_taxable += inv.taxable; b.books_igst += inv.igst; b.books_cgst += inv.cgst;
    b.books_sgst += inv.sgst; b.books_tax += inv.total_tax; b.books_count += 1;
    bump(b.statuses, status);
    if (inv.matched_doc) {
      const d = inv.matched_doc;
      b.r2b_taxable += d.taxable; b.r2b_igst += d.igst; b.r2b_cgst += d.cgst;
      b.r2b_sgst += d.sgst; b.r2b_tax += d.total_tax; b.r2b_count += 1;
    }
  }

  for (const d of docs) {
    if (d.matched) continue;
    const b = get(d.gstin, d.party, d.reverse_charge, r2bType(d));
    b.r2b_taxable += d.taxable; b.r2b_igst += d.igst; b.r2b_cgst += d.cgst;
    b.r2b_sgst += d.sgst; b.r2b_tax += d.total_tax; b.r2b_count += 1;
    bump(b.statuses, "Only in 2B");
  }

  const priority = ["GSTIN mismatch", "Ambiguous prior match", "Tax mismatch", "Only in Books", "Only in 2B", "Match"];
  const rows = [];
  for (const b of buckets.values()) {
    const diff = round2(b.books_tax - b.r2b_tax);
    let status = priority.find((s) => b.statuses.get(s)) || "Match";
    if (status === "Match" && Math.abs(diff) > 2) status = "Tax mismatch";
    rows.push([
      status, b.gstin, b.party, b.type, b.reverse,
      b.books_count, b.r2b_count,
      round2(b.books_taxable), round2(b.books_igst), round2(b.books_cgst),
      round2(b.books_sgst), round2(b.books_tax),
      round2(b.r2b_taxable), round2(b.r2b_igst), round2(b.r2b_cgst),
      round2(b.r2b_sgst), round2(b.r2b_tax),
      diff, suggestedAction(status, diff, b.reverse),
    ]);
  }
  return rows.sort((a, b) => {
    const ka = a[0] === "Match" ? 1 : 0;
    const kb = b[0] === "Match" ? 1 : 0;
    if (ka !== kb) return ka - kb;
    const g = asText(a[1]).localeCompare(asText(b[1]));
    if (g !== 0) return g;
    return asText(a[2]).localeCompare(asText(b[2]));
  });
}

export function build2bBooksDetail(invoices, docs) {
  const rows = [];
  for (const inv of invoices) {
    const doc = inv.matched_doc;
    let status = inv.invoice_match;
    if (inv.gstin_match === "GSTIN mismatch") status = "GSTIN mismatch";
    else if (inv.gstin_match === "Ambiguous prior match") status = "Ambiguous prior match";
    rows.push([
      status, inv.source, inv.itc_type, inv.resolved_gstin,
      inv.resolved_party || inv.supplier_name,
      inv.invoice_no, inv.invoice_date,
      round2(inv.taxable), round2(inv.igst), round2(inv.cgst),
      round2(inv.sgst), round2(inv.total_tax),
      doc ? doc.section : "", doc ? doc.doc_no : "", doc ? doc.doc_date : "",
      doc ? round2(doc.taxable) : "", doc ? round2(doc.igst) : "",
      doc ? round2(doc.cgst) : "", doc ? round2(doc.sgst) : "",
      doc ? round2(doc.total_tax) : "",
      inv.taxable_diff_vs_2b !== null ? inv.taxable_diff_vs_2b : "",
      inv.tax_diff_vs_2b !== null ? inv.tax_diff_vs_2b : "",
      inv.mapping_notes,
      inv.correction_flag, inv.correction_flag ? inv.correction_amount : "", inv.correction_action,
    ]);
  }
  for (const d of docs) {
    if (d.matched) continue;
    rows.push([
      "Only in 2B", "2B JSON", r2bType(d), d.gstin, d.party,
      "", "", "", "", "", "", "",
      d.section, d.doc_no, d.doc_date,
      round2(d.taxable), round2(d.igst), round2(d.cgst), round2(d.sgst), round2(d.total_tax),
      round2(-d.taxable), round2(-d.total_tax),
      "Available in 2B JSON but not found in source ITC file.", "", "", "",
    ]);
  }
  return rows;
}

export function buildCorrectionRows(invoices) {
  return invoices
    .filter((inv) => inv.correction_flag)
    .map((inv) => [
      inv.correction_flag, inv.source, inv.itc_type, inv.invoice_no, inv.sap_invoice_no,
      inv.resolved_gstin, inv.resolved_party || inv.supplier_name,
      inv.invoice_match,
      inv.tax_diff_vs_2b !== null ? inv.tax_diff_vs_2b : "",
      inv.correction_amount, inv.correction_action, inv.mapping_notes,
    ]);
}

export function itcRows(invoices, periodLabel, fyLabel) {
  return invoices.map((inv) => [
    inv.resolved_gstin, inv.resolved_party || inv.supplier_name,
    inv.invoice_no, inv.invoice_date,
    round2(inv.taxable), round2(inv.igst), round2(inv.cgst),
    round2(inv.sgst), round2(inv.total_tax),
    inv.place_of_supply, inv.reverse_charge, inv.itc_available,
    inv.itc_type, periodLabel, periodLabel,
    inv.invoice_match, inv.gstin_match, fyLabel,
    inv.support.support_reference || "",
    inv.support.isd_flag || "",
    inv.support.isd_supplier_gstin || "",
    inv.support.isd_bill_tax !== undefined ? inv.support.isd_bill_tax : "",
    inv.support.isd_tax_diff !== undefined ? inv.support.isd_tax_diff : "",
  ]);
}

export function r2bRows(docs, periodLabel, fyLabel) {
  return docs.map((d) => [
    d.gstin, d.party, d.doc_no, d.doc_date,
    round2(d.taxable), round2(d.igst), round2(d.cgst), round2(d.sgst), round2(d.total_tax),
    d.pos, d.reverse_charge, d.itc_available, r2bType(d),
    periodLabel,
    d.matched ? "Matched to Books" : "Only in 2B",
    d.matched ? "Matched to Books" : "",
    fyLabel,
  ]);
}

export function isdRows(isdDocs, periodLabel, fyLabel) {
  // ISD bill-wise rows in the same 17-column shape as the 2B sheet,
  // plus ISD-specific status columns at the end.
  return isdDocs.map((d) => [
    d.gstin, d.party, d.doc_no, d.doc_date,
    round2(d.taxable), round2(d.igst), round2(d.cgst), round2(d.sgst), round2(d.total_tax),
    d.pos, "N", "Y", "ISD",
    periodLabel,
    d.matched ? "Matched to Sanitized ITC" : "Only in ISD JSON",
    d.gstin_in_itc ? "GSTIN present in Sanitized ITC" : "GSTIN absent from Sanitized ITC",
    fyLabel,
    // ISD-specific tail
    round2(d.invoice_value), d.rates,
    d.matched_source_doc,
    d.matched ? d.matched_books_tax : "",
    d.matched ? d.tax_diff : "",
    round2(d.sanitized_gstin_tax),
    d.eligibility,
  ]);
}

export function boeRows(docs) {
  return docs.map((d) => [
    d.matched ? "Matched" : "Unmatched",
    d.be_no, d.be_date, d.port_code, d.importer_gstin, d.importer_name,
    d.invoice_no, d.invoice_amount, d.currency,
    d.assessable_value, d.bcd, d.sws, d.taxable_for_igst, d.igst,
    d.total_duty, d.total_amount, d.challan_no, d.payment_amount,
    d.matched_source_doc, d.matched_itc_igst, d.igst_diff, d.match_method,
    d.gstin_check, d.twob_check, d.validation_action,
    d.file_name, d.source_path,
  ]);
}
