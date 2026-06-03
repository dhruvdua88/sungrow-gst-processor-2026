// Prior-month alias matching — mirrors load_prior_aliases / apply_prior_aliases
// Browser version: takes a list of {name, workbook} for prior focused output files
import { asText, cleanName, normDoc, amountDiff } from "./helpers.js";
import { bestDoc } from "./resolve.js";
import { readSheetRows } from "./readers.js";

export function loadPriorAliases(priorFiles) {
  // priorFiles: [{name: string, workbook: XLSX workbook}]
  const result = { aliases: new Map(), conflicts: new Map() };
  if (!priorFiles || !priorFiles.length) return result;
  const raw = new Map(); // cleanName -> Map(gstin -> alias)
  for (const { name, workbook } of priorFiles) {
    if (!name.includes("Focused_Output")) continue;
    if (!workbook.SheetNames.includes("Sanitized ITC")) continue;
    let rows;
    try {
      rows = readSheetRows(workbook, "Sanitized ITC");
    } catch {
      continue;
    }
    if (!rows.length) continue;
    const headers = Object.keys(rows[0]);
    if (!headers.includes("Party GSTIN/UIN") || !headers.includes("Party name")) continue;
    for (const row of rows) {
      const gstin = asText(row["Party GSTIN/UIN"]).toUpperCase();
      const party = asText(row["Party name"]);
      if (!gstin || gstin === "IMPORTGOODS" || gstin === "IMPORTSERVICE" || !party) continue;
      const key = cleanName(party);
      if (!key) continue;
      if (!raw.has(key)) raw.set(key, new Map());
      const gmap = raw.get(key);
      if (!gmap.has(gstin)) gmap.set(gstin, { gstin, party, source_files: new Set(), count: 0 });
      const alias = gmap.get(gstin);
      alias.count += 1;
      alias.source_files.add(name);
    }
  }
  for (const [key, gstins] of raw) {
    if (gstins.size === 1) {
      result.aliases.set(key, gstins.values().next().value);
    } else {
      result.conflicts.set(key, new Set(gstins.keys()));
    }
  }
  return result;
}

export function applyPriorAliases(invoices, r2bDocs, prior) {
  if (!prior.aliases.size && !prior.conflicts.size) return;
  for (const inv of invoices) {
    if (inv.source_gstin || inv.matched_doc || inv.source === "IMPORT") continue;
    const key = cleanName(inv.supplier_name);
    if (!key) continue;
    if (prior.conflicts.has(key)) {
      inv.gstin_match = "Ambiguous prior match";
      inv.mapping_method = "Ambiguous prior focused file alias";
      inv.mapping_notes =
        `Prior focused files contain conflicting GSTINs for ${inv.supplier_name}: ` +
        `${[...prior.conflicts.get(key)].sort().join(", ")}.`;
      continue;
    }
    const alias = prior.aliases.get(key);
    if (!alias) continue;
    const cand = bestDoc(
      inv,
      r2bDocs.filter(
        (d) =>
          !d.matched &&
          d.gstin === alias.gstin &&
          d.section !== "IMPG" &&
          d.section !== "ISD" &&
          Math.abs(inv.total_tax - d.total_tax) <= Math.max(2, inv.total_tax * 0.01)
      )
    );
    inv.resolved_gstin = alias.gstin;
    inv.resolved_party = alias.party;
    inv.gstin_match = "Prior focused file alias";
    inv.mapping_method = "Prior focused file alias";
    inv.mapping_confidence = Math.max(inv.mapping_confidence, 68);
    inv.mapping_notes =
      `GSTIN resolved from prior focused file(s): ${[...alias.source_files].sort().join(", ")}.`;
    if (cand) {
      inv.matched_doc = cand;
      cand.matched = true;
      inv.resolved_party = cand.party || alias.party;
      inv.itc_available = cand.itc_available || "Y";
      inv.reverse_charge = cand.reverse_charge === "Y" || inv.ledger_rcm_evidence ? "Y" : "N";
      inv.itc_type = inv.reverse_charge === "Y" ? "RCM-B2B" : "B2B";
      inv.tax_diff_vs_2b = amountDiff(inv.total_tax, cand.total_tax);
      inv.taxable_diff_vs_2b = amountDiff(inv.taxable, cand.taxable);
      inv.invoice_match = Math.abs(inv.tax_diff_vs_2b) <= 2 ? "Match" : "Tax mismatch";
      inv.mapping_method = "Prior alias + 2B GSTIN/tax match";
      inv.mapping_confidence = inv.invoice_match === "Match" ? 86 : 78;
    }
  }
}
