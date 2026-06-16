// ISD JSON loading + eligibility classifier + ITC matching —
// mirrors _isd_component / load_isd_distributions / load_isd_bill_docs /
// classify_isd_bill_eligibility / apply_isd_matches
import {
  asText, asFloat, parseDate, taxTotal, round2, normDoc, amountDiff, STATE_CODE,
} from "./helpers.js";

function isdComponent(doc, prefix) {
  let sum = 0;
  for (const [k, v] of Object.entries(doc)) {
    if (asText(k).toLowerCase().startsWith(prefix)) sum += asFloat(v);
  }
  return round2(sum);
}

export function loadIsdDistributions(payload) {
  if (!payload) return [];
  const rows = [];
  for (const [label, key] of [["Eligible", "elglst"], ["Ineligible", "inelglst"]]) {
    for (const supplier of (payload.isd || {})[key] || []) {
      const state = asText(supplier.statecd);
      for (const item of supplier.doclst || []) {
        const igst = isdComponent(item, "iamt");
        const cgst = isdComponent(item, "camt");
        const sgst = isdComponent(item, "samt");
        const cess = round2(isdComponent(item, "csamt") + isdComponent(item, "cess"));
        rows.push({
          eligibility: label,
          cpty: asText(supplier.cpty).toUpperCase(),
          state: STATE_CODE[state] || state,
          doc_no: asText(item.docnum || item.inum),
          doc_date: parseDate(item.docdt || item.dt),
          igst, cgst, sgst, cess,
          total_tax: taxTotal(igst, cgst, sgst, cess),
        });
      }
    }
  }
  return rows;
}

export function loadIsdBillDocs(payload) {
  if (!payload) return [];
  const rows = [];
  for (const supplier of payload.b2b || []) {
    const gstin = asText(supplier.ctin).toUpperCase();
    for (const invoice of supplier.inv || []) {
      let taxable = 0, igst = 0, cgst = 0, sgst = 0, cess = 0;
      const rates = [];
      for (const item of invoice.itms || []) {
        const d = item.itm_det || {};
        taxable += asFloat(d.txval);
        igst += asFloat(d.iamt);
        cgst += asFloat(d.camt);
        sgst += asFloat(d.samt);
        cess += asFloat(d.csamt) + asFloat(d.cess);
        if (asText(d.rt)) rates.push(asText(d.rt));
      }
      const posCode = asText(invoice.pos);
      rows.push({
        gstin,
        party: `ISD source bill - ${gstin}`,
        doc_no: asText(invoice.inum),
        doc_date: parseDate(invoice.idt),
        taxable: round2(taxable),
        igst: round2(igst), cgst: round2(cgst), sgst: round2(sgst),
        cess: round2(cess),
        total_tax: taxTotal(igst, cgst, sgst, cess),
        invoice_value: asFloat(invoice.val),
        pos: STATE_CODE[posCode] || posCode,
        rates: [...new Set(rates)].join(", "),
        matched: false,
        matched_source_doc: "",
        matched_books_tax: 0,
        tax_diff: 0,
        gstin_in_itc: false,
        sanitized_gstin_tax: 0,
        eligibility: "Unclassified",
      });
    }
  }
  return rows;
}

export function classifyIsdBillEligibility(isdDocs, isdDist) {
  const eligIgst = round2(isdDist.filter((d) => d.eligibility === "Eligible").reduce((s, d) => s + d.igst, 0));
  const eligCgst = round2(isdDist.filter((d) => d.eligibility === "Eligible").reduce((s, d) => s + d.cgst, 0));
  const eligSgst = round2(isdDist.filter((d) => d.eligibility === "Eligible").reduce((s, d) => s + d.sgst, 0));

  const summary = {
    method: "Unable to classify",
    eligible_gstins: [],
    elig_total_tax: round2(eligIgst + eligCgst + eligSgst),
    elig_target: [eligIgst, eligCgst, eligSgst],
    elig_achieved: [0, 0, 0],
    warning: "",
  };

  if (!isdDocs.length) return summary;
  if (!isdDist.length || summary.elig_total_tax === 0) {
    for (const d of isdDocs) d.eligibility = "Unclassified";
    summary.warning = "No ISD distribution voucher in JSON; bills left Unclassified.";
    return summary;
  }

  // Per-supplier totals
  const supTotals = new Map();
  for (const d of isdDocs) {
    const [i, c, s] = supTotals.get(d.gstin) || [0, 0, 0];
    supTotals.set(d.gstin, [round2(i + d.igst), round2(c + d.cgst), round2(s + d.sgst)]);
  }

  const suppliers = [...supTotals.keys()];
  const n = suppliers.length;
  const tol = 0.5;

  // Brute-force subset sum on supplier-level totals (n usually <= 20)
  const matches = [];
  if (n <= 22) {
    const limit = 1 << n;
    for (let mask = 1; mask < limit; mask++) {
      let ti = 0, tc = 0, ts = 0;
      for (let k = 0; k < n; k++) {
        if (mask & (1 << k)) {
          const [a, b, c] = supTotals.get(suppliers[k]);
          ti += a; tc += b; ts += c;
        }
      }
      if (Math.abs(ti - eligIgst) <= tol && Math.abs(tc - eligCgst) <= tol && Math.abs(ts - eligSgst) <= tol) {
        const chosen = new Set(suppliers.filter((_, k) => mask & (1 << k)));
        matches.push([chosen, [round2(ti), round2(tc), round2(ts)]]);
      }
    }
  }

  if (matches.length === 1) {
    const [chosen, achieved] = matches[0];
    for (const d of isdDocs) d.eligibility = chosen.has(d.gstin) ? "Eligible" : "Ineligible";
    summary.method = "Supplier-level subset sum (exact)";
    summary.eligible_gstins = [...chosen].sort();
    summary.elig_achieved = achieved;
    return summary;
  }

  if (matches.length > 1) {
    // Pick the subset with the smallest count of suppliers
    matches.sort((a, b) => a[0].size - b[0].size);
    const [chosen, achieved] = matches[0];
    for (const d of isdDocs) d.eligibility = chosen.has(d.gstin) ? "Eligible" : "Ineligible";
    summary.method = `Supplier-level subset sum (${matches.length} candidates, smallest picked)`;
    summary.eligible_gstins = [...chosen].sort();
    summary.elig_achieved = achieved;
    summary.warning = "Multiple supplier subsets tie to eligible target; review HQ working papers.";
    return summary;
  }

  // No supplier-level match - fall back to pro-rated tag (visible warning)
  const eligTotal = eligIgst + eligCgst + eligSgst;
  const allTax = isdDocs.reduce((s, d) => s + d.total_tax, 0) || 1.0;
  const ratio = eligTotal / allTax;
  for (const d of isdDocs) d.eligibility = "Unclassified";
  summary.method = "Pro-rate fallback";
  summary.warning =
    "No supplier subset ties exactly to eligible target. " +
    `Approximate eligible share = ${(ratio * 100).toFixed(2)}% of bill tax. ` +
    "Get HQ to flag bills manually before claiming.";
  return summary;
}

// ISD matching against Sanitized ITC — mirrors apply_isd_matches
export function applyIsdMatches(invoices, isdDocs) {
  if (!isdDocs.length) return;
  const byInv = new Map();
  const byGstin = new Map();
  for (const inv of invoices) {
    if (inv.source !== "PO" && inv.source !== "Non PO") continue;
    const keys = new Set([normDoc(inv.invoice_no), normDoc(inv.sap_invoice_no)]);
    for (const d of inv.source_docs) keys.add(normDoc(d));
    for (const k of keys) {
      if (!k) continue;
      if (!byInv.has(k)) byInv.set(k, []);
      byInv.get(k).push(inv);
    }
    const g = asText(inv.resolved_gstin || inv.source_gstin).toUpperCase();
    if (g && g !== "IMPORTGOODS" && g !== "IMPORTSERVICE") {
      if (!byGstin.has(g)) byGstin.set(g, []);
      byGstin.get(g).push(inv);
    }
  }

  for (const d of isdDocs) {
    // Invoice-level match
    const cands = byInv.get(normDoc(d.doc_no)) || [];
    if (cands.length) {
      const best = [...cands].sort((a, b) => {
        const ka = a.source_gstin === "" || a.source_gstin === d.gstin ? 0 : 1;
        const kb = b.source_gstin === "" || b.source_gstin === d.gstin ? 0 : 1;
        if (ka !== kb) return ka - kb;
        return Math.abs(a.total_tax - d.total_tax) - Math.abs(b.total_tax - d.total_tax);
      })[0];
      d.matched = true;
      d.matched_source_doc = best.sap_invoice_no || best.invoice_no;
      d.matched_books_tax = round2(best.total_tax);
      d.tax_diff = amountDiff(best.total_tax, d.total_tax);
      best.support.isd_flag = "ISD";
      best.support.isd_supplier_gstin = d.gstin;
      best.support.isd_bill_tax = round2(d.total_tax);
      best.support.isd_tax_diff = d.tax_diff;
      const existing = asText(best.support.support_reference);
      const ref = `ISD bill ${d.doc_no}`;
      best.support.support_reference = existing ? `${existing}; ${ref}` : ref;
    }
    // GSTIN-level presence
    const gstinRows = byGstin.get(d.gstin) || [];
    d.gstin_in_itc = gstinRows.length > 0;
    d.sanitized_gstin_tax = round2(gstinRows.reduce((s, i) => s + i.total_tax, 0));
  }

  // Mark every Sanitized ITC row with how it relates to the ISD return:
  //   "Invoice in ISD" (this bill's doc no is in the ISD b2b) >
  //   "GSTIN in ISD"   (same supplier GSTIN appears in the ISD return) > "No".
  const isdGstins = new Set(isdDocs.map((d) => asText(d.gstin).toUpperCase()).filter(Boolean));
  for (const inv of invoices) {
    const g = asText(inv.resolved_gstin || inv.source_gstin).toUpperCase();
    if (inv.support.isd_flag === "ISD") inv.support.appearing_in_isd = "Invoice in ISD";
    else if (g && isdGstins.has(g)) inv.support.appearing_in_isd = "GSTIN in ISD";
    else inv.support.appearing_in_isd = "No";
  }
}
