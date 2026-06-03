// Ledger tie-out (with substring fallback for the 52,125-style case) —
// mirrors _matching_gl_rows / _ledger_amounts / build_ledger_tie
import { asText, asFloat, normDoc, round2, REC_ACCOUNTS, RCM_ACCOUNTS } from "./helpers.js";

function matchingGlRows(inv, glRows) {
  const docKeys = new Set([normDoc(inv.sap_invoice_no)]);
  for (const d of inv.source_docs) docKeys.add(normDoc(d));
  docKeys.delete("");
  const refKeys = new Set(inv.source === "PO" ? [normDoc(inv.invoice_no)] : []);
  refKeys.delete("");
  const matched = [];
  for (const row of glRows) {
    const account = asText(row["Account"]);
    if (!(account in REC_ACCOUNTS) && !(account in RCM_ACCOUNTS)) continue;
    const rowDoc = new Set([normDoc(row["Document Number"]), normDoc(row["Invoice reference"])]);
    rowDoc.delete("");
    const rowRef = normDoc(row["Reference"]);
    let hit = false;
    for (const k of docKeys) {
      if (rowDoc.has(k)) { hit = true; break; }
    }
    if (hit) { matched.push(row); continue; }
    if (refKeys.size && refKeys.has(rowRef)) { matched.push(row); continue; }
    // Substring fallback: catches cases like PO ref "9090000056/26-27" vs GL ref "90000056/26-27"
    if (inv.source === "PO" && rowRef && rowRef.length >= 8) {
      for (const rk of refKeys) {
        if (rk.length >= 8 && (rowRef.includes(rk) || rk.includes(rowRef))) {
          matched.push(row);
          break;
        }
      }
    }
  }
  return matched;
}

function ledgerAmounts(inv, glRows) {
  const out = {
    rec_igst: 0, rec_cgst: 0, rec_sgst: 0,
    rcm_igst: 0, rcm_cgst: 0, rcm_sgst: 0,
    gl_docs: new Set(), gl_refs: new Set(), line_count: 0,
  };
  for (const row of matchingGlRows(inv, glRows)) {
    const account = asText(row["Account"]);
    const amount = asFloat(row["Amount in local currency"]);
    const comp = REC_ACCOUNTS[account] || RCM_ACCOUNTS[account];
    if (!comp) continue;
    if (account in REC_ACCOUNTS) {
      out[`rec_${comp.toLowerCase()}`] += amount;
    } else {
      out[`rcm_${comp.toLowerCase()}`] += -amount; // RCM payable stored as positive liability
    }
    out.line_count += 1;
    if (asText(row["Document Number"])) out.gl_docs.add(asText(row["Document Number"]));
    if (asText(row["Reference"])) out.gl_refs.add(asText(row["Reference"]));
  }
  out.rec_total = out.rec_igst + out.rec_cgst + out.rec_sgst;
  out.rcm_total = out.rcm_igst + out.rcm_cgst + out.rcm_sgst;
  out.itc_vs_rec_diff = round2(inv.total_tax - out.rec_total);
  const expected = inv.reverse_charge === "Y" ? inv.total_tax : 0;
  out.expected_rcm = expected;
  out.rcm_diff = round2(expected - out.rcm_total);
  out.status =
    Math.abs(out.itc_vs_rec_diff) <= 2 && Math.abs(out.rcm_diff) <= 2 ? "OK" : "Ledger tie mismatch";
  if (out.status !== "OK" && !inv.correction_flag) {
    inv.correction_flag = "Correct books immediately";
    inv.correction_amount = round2(out.itc_vs_rec_diff || out.rcm_diff);
    inv.correction_action =
      "ITC from source does not tie to SAP GST ledger movement for this invoice. " +
      "Correct the GST ledger posting or source classification.";
  }
  return out;
}

export function buildLedgerTie(invoices, glRows) {
  const details = [];
  const buckets = new Map();
  for (const inv of invoices) {
    const amts = ledgerAmounts(inv, glRows);
    const key = JSON.stringify([inv.source, inv.itc_type, inv.reverse_charge]);
    if (!buckets.has(key)) {
      buckets.set(key, {
        taxable: 0, igst: 0, cgst: 0, sgst: 0, tax: 0,
        rec_igst: 0, rec_cgst: 0, rec_sgst: 0,
        rcm_igst: 0, rcm_cgst: 0, rcm_sgst: 0,
        diff: 0, rcm_diff: 0, rows: 0, exceptions: 0,
      });
    }
    const b = buckets.get(key);
    b.rows += 1;
    for (const f of ["taxable", "igst", "cgst", "sgst"]) b[f] += inv[f];
    b.tax += inv.total_tax;
    for (const f of ["rec_igst", "rec_cgst", "rec_sgst", "rcm_igst", "rcm_cgst", "rcm_sgst"]) b[f] += amts[f];
    b.diff += amts.itc_vs_rec_diff;
    b.rcm_diff += amts.rcm_diff;
    if (amts.status !== "OK") b.exceptions += 1;
    details.push([
      amts.status, inv.source, inv.itc_type, inv.reverse_charge,
      inv.resolved_gstin, inv.resolved_party,
      inv.invoice_no, inv.sap_invoice_no, inv.invoice_date,
      round2(inv.taxable), round2(inv.igst), round2(inv.cgst), round2(inv.sgst), round2(inv.total_tax),
      round2(amts.rec_igst), round2(amts.rec_cgst), round2(amts.rec_sgst), round2(amts.rec_total),
      round2(amts.rcm_igst), round2(amts.rcm_cgst), round2(amts.rcm_sgst), round2(amts.rcm_total),
      round2(amts.itc_vs_rec_diff), round2(amts.expected_rcm), round2(amts.rcm_diff),
      [...amts.gl_docs].sort().join(", "), [...amts.gl_refs].sort().join(", "),
    ]);
  }
  const summary = [];
  const sortedKeys = [...buckets.keys()].sort();
  for (const key of sortedKeys) {
    const [source, category, reverse] = JSON.parse(key);
    const b = buckets.get(key);
    const recTotal = b.rec_igst + b.rec_cgst + b.rec_sgst;
    const rcmTotal = b.rcm_igst + b.rcm_cgst + b.rcm_sgst;
    const status =
      Math.abs(b.diff) <= 2 && Math.abs(b.rcm_diff) <= 2 && b.exceptions === 0 ? "OK" : "Review";
    summary.push([
      status, source, category, reverse, b.rows,
      round2(b.taxable), round2(b.igst), round2(b.cgst), round2(b.sgst), round2(b.tax),
      round2(b.rec_igst), round2(b.rec_cgst), round2(b.rec_sgst), round2(recTotal),
      round2(b.rcm_igst), round2(b.rcm_cgst), round2(b.rcm_sgst), round2(rcmTotal),
      round2(b.diff), round2(b.rcm_diff), b.exceptions,
    ]);
  }
  return [summary, details];
}
