// 2B matching / invoice resolution — mirrors _build_r2b_indexes/_best_doc/_resolve_invoice/resolve_all
import { normDoc, cleanName, nameScore, amountDiff, round2 } from "./helpers.js";

function buildR2bIndexes(docs) {
  const byDoc = new Map();
  const bySupplier = new Map();
  for (const doc of docs) {
    const nd = normDoc(doc.doc_no);
    if (nd) {
      if (!byDoc.has(nd)) byDoc.set(nd, []);
      byDoc.get(nd).push(doc);
    }
    const cn = cleanName(doc.party);
    if (cn) {
      if (!bySupplier.has(cn)) bySupplier.set(cn, []);
      bySupplier.get(cn).push(doc);
    }
  }
  return { byDoc, bySupplier };
}

export function bestDoc(invoice, candidates) {
  if (!candidates.length) return null;
  return [...candidates].sort((a, b) => {
    const k1 = Math.abs(invoice.total_tax - a.total_tax) - Math.abs(invoice.total_tax - b.total_tax);
    if (k1 !== 0) return k1;
    const k2 = Math.abs(invoice.taxable - a.taxable) - Math.abs(invoice.taxable - b.taxable);
    if (k2 !== 0) return k2;
    return nameScore(invoice.supplier_name, b.party) - nameScore(invoice.supplier_name, a.party);
  })[0];
}

function resolveInvoice(invoice, docs, indexes) {
  // IMPORT: match by IGST exactly to an IMPG BoE in 2B
  if (invoice.source === "IMPORT") {
    const candidates = docs.filter(
      (d) => d.section === "IMPG" && !d.matched && Math.abs(invoice.igst - d.igst) <= 1
    );
    if (candidates.length) {
      const cand = [...candidates].sort(
        (a, b) => Math.abs(invoice.igst - a.igst) - Math.abs(invoice.igst - b.igst)
      )[0];
      invoice.matched_doc = cand;
      cand.matched = true;
      invoice.resolved_gstin = "IMPORTGOODS";
      invoice.resolved_party = "Import Goods";
      invoice.invoice_no = cand.doc_no;
      invoice.invoice_date = cand.doc_date || invoice.invoice_date;
      if (!invoice.taxable) invoice.taxable = cand.taxable;
      invoice.itc_available = cand.itc_available;
      invoice.itc_type = "IMPORTGOODS";
      invoice.invoice_match = "Match";
      invoice.gstin_match = "Import";
      invoice.mapping_method = "Import IGST exact match to 2B BoE";
      invoice.mapping_confidence = 98;
      invoice.tax_diff_vs_2b = amountDiff(invoice.total_tax, cand.total_tax);
      invoice.taxable_diff_vs_2b = amountDiff(invoice.taxable, cand.taxable);
      invoice.mapping_notes = `Matched BoE ${cand.doc_no} by IGST amount.`;
    } else {
      invoice.resolved_gstin = "IMPORTGOODS";
      invoice.resolved_party = "Import Goods";
      invoice.itc_type = "IMPORTGOODS";
      invoice.gstin_match = "Import";
      invoice.mapping_method = "Import marker - no 2B BoE match";
      invoice.mapping_confidence = 60;
      invoice.mapping_notes = "Country is non-India; no exact BoE/IGST match found in the 2B JSON.";
    }
    return;
  }

  // Exact invoice match
  const exactCands = (indexes.byDoc.get(normDoc(invoice.invoice_no)) || []).filter((d) => !d.matched);
  let cand = bestDoc(invoice, exactCands);
  if (cand) {
    invoice.matched_doc = cand;
    cand.matched = true;
    invoice.resolved_gstin = invoice.source_gstin || cand.gstin;
    invoice.resolved_party = cand.party || invoice.supplier_name;
    invoice.itc_available = cand.itc_available || "Y";
    const finalReverse = cand.reverse_charge === "Y" || invoice.ledger_rcm_evidence;
    invoice.reverse_charge = finalReverse ? "Y" : "N";
    invoice.itc_type = finalReverse ? "RCM-B2B" : "B2B";
    invoice.tax_diff_vs_2b = amountDiff(invoice.total_tax, cand.total_tax);
    invoice.taxable_diff_vs_2b = amountDiff(invoice.taxable, cand.taxable);
    invoice.invoice_match = Math.abs(invoice.tax_diff_vs_2b) <= 2 ? "Match" : "Tax mismatch";
    invoice.gstin_match = invoice.source_gstin ? "Source GSTIN" : "Resolved from 2B";
    invoice.mapping_method = "Exact invoice/reference match to 2B";
    invoice.mapping_confidence = invoice.invoice_match === "Match" ? 99 : 92;
    const notes = [];
    if (invoice.invoice_match === "Tax mismatch") {
      invoice.correction_flag = "Correct books immediately";
      invoice.correction_amount = round2(invoice.tax_diff_vs_2b || 0);
      invoice.correction_action = "Exact 2B invoice exists but tax differs from books.";
    }
    if (invoice.source_gstin && cand.gstin && invoice.source_gstin !== cand.gstin) {
      invoice.gstin_match = "GSTIN mismatch";
      notes.push(`Source GSTIN ${invoice.source_gstin} differs from 2B GSTIN ${cand.gstin}.`);
    }
    if (invoice.source_rcm_flag && cand.reverse_charge !== "Y" && !invoice.ledger_rcm_evidence) {
      invoice.correction_flag = "Correct books immediately";
      invoice.correction_amount = round2(invoice.source_rcm_tax || invoice.total_tax);
      invoice.correction_action =
        "Source PO marks invoice as RCM, but 2B and GST ledger support normal B2B ITC. Remove RCM flag.";
      notes.push("Sanitized as B2B; neither 2B nor GL supports RCM.");
    }
    invoice.mapping_notes = notes.join(" ");
    return;
  }

  // Supplier-name fallback
  const supplierHits = [];
  for (const doc of docs) {
    if (doc.matched || doc.section === "IMPG" || doc.section === "ISD") continue;
    const score = nameScore(invoice.supplier_name, doc.party);
    if (score >= 0.78) {
      const taxClose = Math.abs(invoice.total_tax - doc.total_tax) <= Math.max(2, invoice.total_tax * 0.01);
      const taxableClose = Math.abs(invoice.taxable - doc.taxable) <= Math.max(2, invoice.taxable * 0.02);
      if (taxClose || taxableClose || score >= 0.93) supplierHits.push([score, doc]);
    }
  }
  if (supplierHits.length) {
    supplierHits.sort((x, y) => {
      if (y[0] !== x[0]) return y[0] - x[0];
      const k1 = Math.abs(invoice.total_tax - x[1].total_tax) - Math.abs(invoice.total_tax - y[1].total_tax);
      if (k1 !== 0) return k1;
      return Math.abs(invoice.taxable - x[1].taxable) - Math.abs(invoice.taxable - y[1].taxable);
    });
    const [score, scand] = supplierHits[0];
    const sameTax = Math.abs(invoice.total_tax - scand.total_tax) <= 2;
    if (score >= 0.86 && (sameTax || !invoice.source_gstin)) {
      if (sameTax) {
        scand.matched = true;
        invoice.matched_doc = scand;
      }
      invoice.resolved_gstin = invoice.source_gstin || scand.gstin;
      invoice.resolved_party = scand.party || invoice.supplier_name;
      invoice.itc_available = scand.itc_available || "Y";
      const finalReverse = scand.reverse_charge === "Y" || invoice.ledger_rcm_evidence;
      invoice.reverse_charge = finalReverse ? "Y" : "N";
      invoice.itc_type = finalReverse ? "RCM-B2B" : "B2B";
      invoice.tax_diff_vs_2b = amountDiff(invoice.total_tax, scand.total_tax);
      invoice.taxable_diff_vs_2b = amountDiff(invoice.taxable, scand.taxable);
      invoice.invoice_match = sameTax ? "Match" : "Only in Books";
      invoice.gstin_match = invoice.source_gstin ? "Source GSTIN" : "Resolved from supplier name";
      invoice.mapping_method = "Supplier-name best match to 2B";
      invoice.mapping_confidence = sameTax ? 82 : 72;
      invoice.mapping_notes =
        `Best 2B supplier match: ${scand.party}; inv ${scand.doc_no}; score ${score.toFixed(2)}.`;
      if (invoice.source_rcm_flag && scand.reverse_charge !== "Y" && !invoice.ledger_rcm_evidence) {
        invoice.correction_flag = "Correct books immediately";
        invoice.correction_amount = round2(invoice.source_rcm_tax || invoice.total_tax);
        invoice.correction_action =
          "Source marks invoice as RCM, but 2B and GST ledger support normal B2B ITC. Remove RCM flag.";
      }
      return;
    }
  }

  // Unresolved
  invoice.resolved_gstin = invoice.source_gstin;
  invoice.resolved_party = invoice.supplier_name;
  invoice.reverse_charge = invoice.source_rcm ? "Y" : "N";
  invoice.itc_type = invoice.source_rcm && invoice.source_gstin
    ? "RCM-B2B"
    : invoice.source_rcm
      ? "RCM-UR"
      : "B2B";
  invoice.gstin_match = invoice.source_gstin ? "Source GSTIN" : "Unresolved";
  invoice.mapping_method = "No 2B match";
  invoice.mapping_confidence = invoice.source_gstin ? 40 : 0;
  invoice.mapping_notes = "No exact invoice or reliable supplier reference found in the 2B JSON.";
}

export function resolveAll(invoices, r2bDocs) {
  const indexes = buildR2bIndexes(r2bDocs);
  invoices.sort((a, b) => {
    const ka = a.source === "IMPORT" ? 0 : 1;
    const kb = b.source === "IMPORT" ? 0 : 1;
    if (ka !== kb) return ka - kb;
    if (a.source !== b.source) return a.source < b.source ? -1 : 1;
    if (a.invoice_no !== b.invoice_no) return a.invoice_no < b.invoice_no ? -1 : 1;
    return 0;
  });
  for (const inv of invoices) resolveInvoice(inv, r2bDocs, indexes);
}
