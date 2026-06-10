import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { processGstr1, validateDocRow, buildJson, VALID_RATES, VALID_UQC } from "./engine.js";
import "./gstr1.css";

const inr = (n) => (n == null || n === "" ? "—" : Number(n).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
const inr0 = (n) => (n == null ? "—" : "₹" + Number(n).toLocaleString("en-IN", { maximumFractionDigits: 0 }));
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const GSTIN_RE = /([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][0-9A-Z]Z[0-9A-Z])/;

function FilePick({ label, hint, accept, file, onFile }) {
  const ref = useRef(null);
  return (
    <div className="g1-filerow">
      <label className="g1-label">{label}</label>
      <div className={`g1-drop ${file ? "filled" : ""}`} onClick={() => ref.current?.click()} title={file?.name || ""}>
        {file ? file.name : hint}
      </div>
      <button type="button" className="g1-browse" onClick={() => ref.current?.click()}>Browse</button>
      {file && <button type="button" className="g1-clear" onClick={() => { onFile(null); if (ref.current) ref.current.value = ""; }}>×</button>}
      <input ref={ref} type="file" accept={accept} style={{ display: "none" }} onChange={(e) => onFile(e.target.files?.[0] || null)} />
    </div>
  );
}

function Banner({ res }) {
  const cls = res.portalReady ? (res.allPass ? "ok" : "warn") : "bad";
  const head = res.portalReady
    ? res.allPass ? "ALL CHECKS PASSED — GSTR-1 JSON ready to file" : "PORTAL-VALID — resolve deviations with client, JSON can be generated"
    : "BLOCKING VALIDATION FAILURES — fix before filing";
  return (
    <div className={`g1-banner ${cls}`}>
      <div className="g1-banner-h">{head}</div>
      <div className="g1-banner-s">
        {res.checks.filter((c) => c.ok).length}/{res.checks.length} checks passed
        {res.fails.length ? ` · ${res.fails.length} deviation(s): ${res.fails.map((c) => c.id).join(", ")}` : ""}
        {res.blockingFails.length ? ` · BLOCKING: ${res.blockingFails.map((c) => c.id).join(", ")}` : ""}
      </div>
    </div>
  );
}

function Cards({ res }) {
  const t = res.totals, c = res.counts;
  const cards = [
    ["Table 12 taxable", inr0(t.t12Txbl)], ["IGST", inr0(t.t12Igst)], ["CGST", inr0(t.t12Cgst)], ["SGST", inr0(t.t12Sgst)],
    ["JSON B2B invoices", c.jsonB2bInvoices, true], ["Credit notes", c.cdnr, true], ["Zero-value docs (excl.)", c.zeroDocs, true], ["HSN rows", c.hsnRows, true],
  ];
  return (
    <div className="g1-cards">
      {cards.map(([l, v]) => (<div key={l} className="g1-card"><span className="g1-card-l">{l}</span><span className="g1-card-v">{v}</span></div>))}
    </div>
  );
}

function OvBucket({ title, map, note }) {
  return (
    <div className="g1-ov-block">
      <div className="g1-ov-bt">{title}</div>
      <table className="g1-table compact">
        <thead><tr><th className="left">Category</th><th className="num">Invoices</th><th className="num">Lines</th><th className="num">Taxable</th><th className="num">Tax</th></tr></thead>
        <tbody>
          {Object.entries(map).map(([k, v]) => (
            <tr key={k} className={/non-?revenue/i.test(k) ? "row-warn" : ""}>
              <td className="left">{k}</td><td className="num">{v.invoices}</td><td className="num">{v.lines}</td><td className="num">{inr(v.txbl)}</td><td className="num">{v.tax ? inr(v.tax) : "—"}</td>
            </tr>
          ))}
        </tbody>
      </table>
      {note && <p className="g1-ov-note">{note}</p>}
    </div>
  );
}

function FileOverview({ res }) {
  const f = res.fileOverview;
  const cr = (n) => "₹" + Number(n || 0).toLocaleString("en-IN", { maximumFractionDigits: 0 });
  return (
    <section className="g1-card g1-ov">
      <h3 className="g1-ov-h">📋 File Overview — what the client sent <span>({f.sales.invoices} invoices · {f.sales.lines} lines in “Sales” · {cr(f.sales.txbl)} taxable)</span></h3>
      <div className="g1-ov-grid">
        <OvBucket title="By Type of Revenue" map={f.sales.byRevenue}
          note="Yes-Revenue = booked revenue. Non-Revenue = billed but revenue deferred (still in GSTR-1 if e-invoiced)." />
        <OvBucket title="By Shipment" map={f.sales.byShipment} note={f.sales.byShipment["Export / SEZ"] ? "Export/SEZ present — handled in the exports module (planned)." : "All domestic this period — no exports/SEZ."} />
        <OvBucket title="By Customer Type" map={f.sales.byCustomer} note={f.sales.byCustomer["Unregistered (B2C)"]?.invoices ? "B2C present — handled in the B2C module (planned)." : "All registered (B2B) this period — no B2C."} />
      </div>
      <div className="g1-ov-extra">
        <div className="g1-ov-pill excl"><span>SRV FOC — free of cost (EXCLUDED, not uploaded)</span><b>{f.foc.lines} lines · qty {Math.abs(f.foc.qty).toLocaleString("en-IN")} · LC {cr(Math.abs(f.foc.amt))}</b></div>
        <div className="g1-ov-pill"><span>SRV details — service dispatches</span><b>{f.srv.lines} lines · {f.srv.dispatches} dispatches</b></div>
        <div className="g1-ov-pill"><span>Support sheets</span><b>Summary {f.support.summaryRows} · Physical stock {f.support.stockRows}</b></div>
      </div>
      <p className="g1-ov-foot">→ Of this file, only valid B2B e-invoices (incl. Non-Revenue ones carrying an IRN) flow to GSTR-1. FOC service replacements have no billing and are excluded. Sheets read: {f.sheets.join(", ")}.</p>
    </section>
  );
}

function ChecksTable({ res }) {
  return (
    <table className="g1-table checks">
      <thead><tr><th>#</th><th>Check</th><th>Expected</th><th>Actual</th><th>Result</th></tr></thead>
      <tbody>
        {res.checks.map((c) => (
          <tr key={c.id} className={c.ok ? "" : res.blockingFails.includes(c) ? "row-bad" : "row-warn"}>
            <td>{c.id}</td><td className="left">{c.desc}</td><td className="left small">{c.expected}</td><td className="left small">{c.actual}</td>
            <td className={c.ok ? "pass" : "fail"}>{c.ok ? "PASS" : (res.blockingFails.includes(c) ? "FAIL" : "DEVIATION")}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function Deviations({ res }) {
  const r = res.recon;
  return (
    <div className="g1-dev">
      <div className="g1-recon-grid">
        <div><span>Client Sales book (on valid e-invoices)</span><b>{inr(r.salesValidTxbl)}</b></div>
        <div><span>Portal e-invoices (same set)</span><b>{inr(r.portalValidTxbl)}</b></div>
        <div className={Math.abs(r.diff) > 1 ? "diffbad" : "diffok"}><span>Difference</span><b>{inr(r.diff)}</b></div>
      </div>

      {r.valMismatch.length > 0 && (
        <>
          <h4 className="g1-dev-h bad">⚠ Value mismatches — client book vs portal e-invoice (ask Naveen)</h4>
          <table className="g1-table"><thead><tr><th>Invoice</th><th>Recipient</th><th className="num">Sales book</th><th className="num">e-Invoice (portal)</th><th className="num">Difference</th></tr></thead>
            <tbody>{r.valMismatch.map((m) => (<tr key={m.inum} className="row-bad"><td>{m.inum}</td><td className="left">{m.name}</td><td className="num">{inr(m.sales)}</td><td className="num">{inr(m.portal)}</td><td className="num">{inr(m.diff)}</td></tr>))}</tbody>
          </table>
        </>
      )}

      {res.cancelled.length > 0 && (
        <p className="g1-note"><b>Portal-cancelled IRNs:</b> {res.cancelled.join(", ")} — the JSON carries the BOOK copy of any of these still in the sales sheet; confirm with the client.</p>
      )}
      {res.cdnr.length > 0 && (
        <p className="g1-note"><b>Credit/Debit notes (reduce Table 12):</b> {res.cdnr.map((c) => `${c.nt_num} (${c.ntty}, ${c.name}, ${inr(c.txbl)})`).join("; ")}</p>
      )}
      {r.missingInSales.length > 0 && (
        <p className="g1-note bad"><b>e-Invoices missing from client book:</b> {r.missingInSales.join(", ")}</p>
      )}

      {res.recon.nonRevButEinvoiced.length > 0 && (
        <>
          <h4 className="g1-dev-h warn">ℹ {res.recon.nonRevButEinvoiced.length} invoices tagged "Non-Revenue" carry a valid IRN → INCLUDED in GSTR-1</h4>
          <p className="g1-note">"Type of Revenue" is an accounting flag, not a GST flag. The portal already auto-populated these from the IRN; filtering them out would understate the return. Confirm revenue-recognition treatment with the client separately.</p>
          <table className="g1-table"><thead><tr><th>Invoice</th><th>Recipient</th><th>Client tag</th><th className="num">Portal taxable</th></tr></thead>
            <tbody>{res.recon.nonRevButEinvoiced.map((m) => (<tr key={m.inum} className="row-warn"><td>{m.inum}</td><td className="left">{m.name}</td><td className="left small">{m.rev}</td><td className="num">{inr(m.txbl)}</td></tr>))}</tbody>
          </table>
        </>
      )}
    </div>
  );
}

async function buildExcel(res) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Sungrow GST Processor 2026";
  const hdrFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2E5496" } };
  const hdrFont = { bold: true, color: { argb: "FFFFFFFF" } };
  const style = (ws, row) => row.eachCell((c) => { c.fill = hdrFill; c.font = hdrFont; c.alignment = { horizontal: "center", wrapText: true }; });

  const s1 = wb.addWorksheet("Summary");
  s1.columns = [{ width: 34 }, { width: 60 }];
  const rows1 = [
    ["GSTIN", res.meta.gstin], ["Return period (fp)", res.meta.fp], ["Status", res.portalReady ? (res.allPass ? "ALL CHECKS PASSED" : "PORTAL-VALID, deviations to resolve") : "BLOCKING FAILURES"],
    ["Checks", `${res.checks.filter((c) => c.ok).length}/${res.checks.length} passed`],
    ["Table 12 taxable", res.totals.t12Txbl], ["Table 12 IGST", res.totals.t12Igst], ["Table 12 CGST", res.totals.t12Cgst], ["Table 12 SGST", res.totals.t12Sgst],
    ["Valid e-invoices", res.counts.validInvoices], ["Cancelled IRNs", res.cancelled.join(", ")], ["Credit notes", res.counts.cdnr],
  ];
  rows1.forEach((r) => { const row = s1.addRow(r); row.getCell(1).font = { bold: true }; });

  const s2 = wb.addWorksheet("Table 12 - HSN");
  const h2 = s2.addRow(["Sr", "HSN/SAC", "Description", "UQC", "Qty", "Rate %", "Taxable", "IGST", "CGST", "SGST", "Cess", "Total"]); style(s2, h2);
  s2.columns = [{ width: 5 }, { width: 12 }, { width: 34 }, { width: 7 }, { width: 11 }, { width: 7 }, { width: 16 }, { width: 14 }, { width: 13 }, { width: 13 }, { width: 9 }, { width: 16 }];
  res.table12.forEach((h) => s2.addRow([h.sr, h.hsn, h.desc, h.uqc, h.qty, h.rt, h.txval, h.iamt, h.camt, h.samt, h.csamt, h.total]));
  s2.addRow(["", "TOTAL", "", "", "", "", res.totals.t12Txbl, res.totals.t12Igst, res.totals.t12Cgst, res.totals.t12Sgst, res.totals.t12Cess, ""]).font = { bold: true };

  const s3 = wb.addWorksheet("Table 13 - Docs");
  const h3 = s3.addRow(["Code", "Nature of Document", "From", "To", "Total", "Cancelled", "Net"]); style(s3, h3);
  s3.columns = [{ width: 7 }, { width: 34 }, { width: 14 }, { width: 14 }, { width: 10 }, { width: 11 }, { width: 10 }];
  res.table13.forEach((d) => s3.addRow([d.code, d.nature, d.from, d.to, d.total, d.cancel, d.net]));

  const s4 = wb.addWorksheet("Validation Checks");
  const h4 = s4.addRow(["#", "Check", "Expected", "Actual", "Result"]); style(s4, h4);
  s4.columns = [{ width: 6 }, { width: 54 }, { width: 26 }, { width: 40 }, { width: 12 }];
  res.checks.forEach((c) => { const row = s4.addRow([c.id, c.desc, c.expected, c.actual, c.ok ? "PASS" : (res.blockingFails.includes(c) ? "FAIL" : "DEVIATION")]); row.getCell(5).fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.ok ? "FFC6EFCE" : "FFFFC7CE" } }; });

  const s5 = wb.addWorksheet("Deviations");
  s5.columns = [{ width: 14 }, { width: 45 }, { width: 18 }, { width: 18 }, { width: 18 }];
  s5.addRow(["DEVIATION REGISTER — client report vs portal e-invoices"]).font = { bold: true, size: 13 };
  s5.addRow([]);
  const hh = s5.addRow(["Invoice", "Recipient", "Sales book", "Portal", "Difference"]); style(s5, hh);
  res.recon.valMismatch.forEach((m) => s5.addRow([m.inum, m.name, m.sales, m.portal, m.diff]));
  s5.addRow([]);
  s5.addRow(["Cancelled IRNs", res.cancelled.join(", ")]);
  s5.addRow(["Credit notes", res.cdnr.map((c) => c.nt_num).join(", ")]);
  s5.addRow(["Non-Revenue but e-invoiced (included)", res.recon.nonRevButEinvoiced.length]);
  res.recon.nonRevButEinvoiced.forEach((m) => s5.addRow([m.inum, m.name, "", "", m.txbl]));

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

// Multi-sheet, hyperlinked client workbook: Action Points cover + detailed working sheets.
async function buildClientExcel(res) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Sungrow GST Processor 2026";
  const BRAND = "FF1E4E8C", ACCENT = "FF0EA5E9", LIGHT = "FFEFF6FF", BAND = "FFF6FAFF", INK = "FF0F2C52";
  const RED = "FFFEE2E2", AMBER = "FFFEF3C7", GREEN = "FFDCFCE7";
  const MONEY = "#,##0.00";
  const thin = { style: "thin", color: { argb: "FFD9E2EC" } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const fill = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
  const t = res.totals;
  const totalTax = (t.t12Igst || 0) + (t.t12Cgst || 0) + (t.t12Sgst || 0) + (t.t12Cess || 0);
  const totalVal = (t.t12Txbl || 0) + totalTax;
  const f = (n) => Number(n || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // ---- per-sheet helpers ----
  const banner = (ws, span, text, sub2) => {
    ws.mergeCells(`A1:${span}1`);
    const c = ws.getCell("A1");
    c.value = text; c.font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    c.alignment = { vertical: "middle", indent: 1 }; c.fill = fill(BRAND); ws.getRow(1).height = 30;
    ws.mergeCells(`A2:${span}2`);
    const s = ws.getCell("A2");
    s.value = sub2; s.font = { bold: true, size: 11, color: { argb: "FFFFFFFF" } };
    s.alignment = { vertical: "middle", indent: 1 }; s.fill = fill(ACCENT); ws.getRow(2).height = 20;
    ws.addRow([]);
  };
  const sectionTitle = (ws, span, text) => {
    const r = ws.addRow([text]);
    ws.mergeCells(`A${r.number}:${span}${r.number}`);
    const c = r.getCell(1);
    c.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
    c.fill = fill(ACCENT); c.alignment = { vertical: "middle", indent: 1 }; r.height = 22;
  };
  const headerRow = (ws, cells) => {
    const r = ws.addRow(cells);
    r.eachCell((c) => { c.fill = fill(BRAND); c.font = { bold: true, size: 10, color: { argb: "FFFFFFFF" } }; c.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; c.border = border; });
    r.height = 24; return r;
  };
  const dataRow = (ws, cells, i, moneyCols = []) => {
    const r = ws.addRow(cells);
    r.eachCell((c) => { c.border = border; if (i % 2) c.fill = fill(BAND); });
    moneyCols.forEach((col) => (r.getCell(col).numFmt = MONEY));
    return r;
  };
  const link = (sheet) => ({ text: `→ ${sheet}`, hyperlink: `#'${sheet}'!A1` });

  const SH = { action: "Action Points", summary: "Summary", recon: "Reconciliation", t12: "Table 12 HSN", t13: "Table 13 Docs", canc: "Cancelled Invoices", dev: "Deviations", checks: "Checks" };

  // ===== Build the issue / action-point list =====
  const cancTot = (res.cancelledDetail || []).reduce((a, c) => a + c.txbl, 0);
  const actionFor = (id) => ({
    C1: "Fix the HSN/SAC code (numeric, 6 or 8 digits) — editable in the app's Fix & Generate tab. JSON BLOCKED until corrected.",
    C2: "Correct the GST rate to a valid slab. JSON BLOCKED until fixed.",
    C3: "Use a portal-standard UQC — editable in Fix & Generate. JSON BLOCKED until fixed.",
    C4: "Add the HSN description (portal Phase-3 mandatory) — editable in Fix & Generate. JSON BLOCKED until fixed.",
    C5: "Recheck tax = taxable × rate for the flagged HSN.",
    C6: "Make CGST = SGST on the flagged intra-state HSN.",
    C12: "B2C / export documents found in the books — these JSON sections are pending; file them via the portal/offline tool this month.",
    C13: "Add the missing e-invoice(s) to the sales sheet (possible omission).",
    C14: "Confirm the correct taxable value with the client (book vs portal differ). The JSON carries the BOOK figure.",
    C16: "A portal credit/debit note is not reflected in the sales sheet — Table 12 may not be net of CDN.",
    C17: "JSON no longer ties to the books — engine issue, re-run with fresh files.",
    C18: "JSON B2B − CDN no longer equals Table-12 hsn_b2b — engine issue, re-run with fresh files.",
    C20: "Fix the flagged document fields (GSTIN/date/POS/value/tax-split) in the Fix & Generate tab. JSON BLOCKED until fixed.",
  }[id] || "Review and resolve with the client.");
  const whereFor = (id) => (["C9"].includes(id) ? SH.canc : ["C13", "C14"].includes(id) ? SH.dev : ["C7", "C8", "C15", "C16"].includes(id) ? SH.recon : ["C10", "C11"].includes(id) ? SH.t13 : SH.t12);
  const issues = [];
  for (const c of res.checks) {
    if (c.ok) continue;
    const blocker = res.blockingFails.includes(c);
    issues.push({ sev: blocker ? "BLOCKER" : "REVIEW", area: c.id, issue: c.desc, detail: `Found: ${c.actual}`, where: whereFor(c.id), action: actionFor(c.id) });
  }
  if ((res.cancelledDetail || []).length) {
    issues.push({ sev: "CONFIRM", area: "Cancelled", issue: `${res.cancelledDetail.length} invoice(s) in the sales sheet but IRN cancelled on the portal`, detail: `₹${f(cancTot)} — ${res.cancelledDetail.slice(0, 3).map((c) => `${c.inum} ${c.name}`).join("; ")}`, where: SH.canc, action: "Confirm with client: remove from GSTR-1 (sale not done) or it was re-billed under a new number." });
  }
  const sevRank = { BLOCKER: 0, CONFIRM: 1, REVIEW: 2, INFO: 3 };
  issues.sort((a, b) => sevRank[a.sev] - sevRank[b.sev]);
  const sevFill = { BLOCKER: RED, CONFIRM: AMBER, REVIEW: AMBER, INFO: LIGHT };

  // ===== Sheet 1: ACTION POINTS (cover) =====
  const a = wb.addWorksheet(SH.action, { views: [{ showGridLines: false }] });
  a.columns = [{ width: 12 }, { width: 14 }, { width: 46 }, { width: 50 }, { width: 16 }, { width: 60 }];
  banner(a, "F", "GSTR-1 — ACTION POINTS & ISSUES", `${res.meta.supplierName ? res.meta.supplierName + "   ·   " : ""}GSTIN ${res.meta.gstin || "—"}      Return Period: ${res.meta.period || res.meta.fp}`);
  const statusRow = a.addRow([res.portalReady ? "JSON READY TO GENERATE" : "JSON BLOCKED — FIX BLOCKERS FIRST"]);
  a.mergeCells(`A${statusRow.number}:F${statusRow.number}`);
  statusRow.getCell(1).font = { bold: true, size: 13, color: { argb: res.portalReady ? "FF166534" : "FF991B1B" } };
  statusRow.getCell(1).fill = fill(res.portalReady ? GREEN : RED);
  statusRow.getCell(1).alignment = { vertical: "middle", indent: 1 }; statusRow.height = 24;
  const cnt = (s) => issues.filter((i) => i.sev === s).length;
  const sline = a.addRow([`${cnt("BLOCKER")} blocker(s) · ${cnt("CONFIRM")} to confirm · ${cnt("REVIEW")} to review · ${res.checks.filter((c) => c.ok).length}/${res.checks.length} checks passed`]);
  a.mergeCells(`A${sline.number}:F${sline.number}`); sline.getCell(1).font = { italic: true, color: { argb: "FF64748B" } };
  a.addRow([]);
  sectionTitle(a, "F", "WHAT NEEDS ACTION");
  headerRow(a, ["Severity", "Ref", "Issue", "Detail", "Where", "What to do"]);
  if (!issues.length) { const r = a.addRow(["—", "", "No issues. Return is clean.", "", "", ""]); r.getCell(3).font = { bold: true, color: { argb: "FF166534" } }; }
  issues.forEach((it, i) => {
    const r = a.addRow([it.sev, it.area, it.issue, it.detail, "", it.action]);
    r.eachCell((c) => { c.border = border; c.alignment = { vertical: "top", wrapText: true }; });
    r.getCell(1).fill = fill(sevFill[it.sev]); r.getCell(1).font = { bold: true, color: { argb: it.sev === "BLOCKER" ? "FF991B1B" : "FF92400E" } };
    const lk = r.getCell(5); lk.value = link(it.where); lk.font = { color: { argb: "FF0369A1" }, underline: true };
    if (i % 2) [2, 3, 4, 6].forEach((col) => (r.getCell(col).fill = fill(BAND)));
  });
  a.addRow([]);
  sectionTitle(a, "F", "WORKBOOK INDEX");
  [[SH.summary, "Headline figures & tax split"], [SH.recon, "Sales sheet ↔ e-invoice reconciliation bridge"], [SH.t12, "Table 12 — full HSN/SAC summary (the books)"], [SH.t13, "Table 13 — documents issued"], [SH.canc, "Invoices cancelled on portal but in the books"], [SH.dev, "Per-invoice deviations (book vs portal)"], [SH.checks, "All 16 validation checks"]].forEach(([s, d], i) => {
    const r = a.addRow(["", "", { text: s, hyperlink: `#'${s}'!A1` }, d, "", ""]);
    r.getCell(3).font = { color: { argb: "FF0369A1" }, underline: true, bold: true };
    r.getCell(4).font = { color: { argb: "FF475569" } };
    if (i % 2) r.eachCell((c) => (c.fill = fill(BAND)));
  });

  // ===== Sheet 2: SUMMARY =====
  const s = wb.addWorksheet(SH.summary, { views: [{ showGridLines: false }] });
  s.columns = [{ width: 30 }, { width: 26 }, { width: 30 }, { width: 26 }];
  banner(s, "D", "GSTR-1 — SUMMARY", `Return Period ${res.meta.period || res.meta.fp}   ·   ${link(SH.action).text.replace("→ ", "back to ")}`);
  s.getCell("C2").value = link(SH.action); s.getCell("C2").font = { color: { argb: "FFFFFFFF" }, underline: true };
  const kv = (label, value, money) => { const r = s.addRow([label, value, "", ""]); r.getCell(1).font = { bold: true, color: { argb: INK } }; if (money) r.getCell(2).numFmt = MONEY; r.getCell(1).fill = fill(LIGHT); r.eachCell((c) => (c.border = border)); };
  kv("Status", res.portalReady ? (res.allPass ? "ALL CHECKS PASSED" : "PORTAL-VALID — deviations to resolve") : "BLOCKING FAILURES — JSON blocked");
  kv("Checks passed", `${res.checks.filter((c) => c.ok).length} / ${res.checks.length}`);
  kv("Table 12 taxable (= sales sheet)", t.t12Txbl, true);
  kv("IGST", t.t12Igst, true); kv("CGST", t.t12Cgst, true); kv("SGST", t.t12Sgst, true); kv("Cess", t.t12Cess, true);
  kv("Total tax", totalTax, true); kv("Total invoice value", totalVal, true);
  kv("Valid e-invoices", res.counts.validInvoices); kv("Cancelled IRNs", res.cancelled.join(", ") || "none");
  kv("Credit/Debit notes", res.counts.cdnr);

  // ===== Sheet 3: RECONCILIATION =====
  if (res.bridge) {
    const b = res.bridge;
    const rc = wb.addWorksheet(SH.recon, { views: [{ showGridLines: false }] });
    rc.columns = [{ width: 4 }, { width: 50 }, { width: 22 }, { width: 56 }];
    banner(rc, "D", "SALES SHEET → GSTR-1 RECONCILIATION", "Panel 1: where every document in the sales sheet goes. Panel 2: cross-check against the e-invoice portal.");
    rc.getCell("C2").value = link(SH.action); rc.getCell("C2").font = { color: { argb: "FFFFFFFF" }, underline: true };
    const brRow = (label, amount, note, kind) => {
      const r = rc.addRow(["", label, amount, note || ""]);
      r.getCell(3).numFmt = MONEY;
      r.getCell(4).font = { size: 9, italic: true, color: { argb: "FF94A3B8" } };
      const strong = kind === "base" || kind === "result";
      r.getCell(2).font = { bold: strong, color: { argb: INK } };
      r.getCell(3).font = { bold: strong, color: { argb: INK } };
      if (strong) [2, 3, 4].forEach((c) => (r.getCell(c).fill = fill(LIGHT)));
      if (kind === "warn") [2, 3, 4].forEach((c) => (r.getCell(c).fill = fill(AMBER)));
      if (kind === "residual" && Math.abs(amount) >= 1) [2, 3].forEach((c) => (r.getCell(c).fill = fill(RED)));
      [2, 3, 4].forEach((c) => (r.getCell(c).border = border));
    };
    sectionTitle(rc, "D", "1 · What's in the client sales sheet, and where each piece goes");
    b.composition.forEach((row) => brRow(row.label, row.amt, row.note));
    brRow("Total = Client sales sheet taxable", b.salesSheetTxbl, Math.abs(b.residual0) < 1 ? "every document accounted for" : `${f(b.residual0)} unaccounted — INVESTIGATE`, "result");
    brRow("Table 12 (HSN summary in the return)", b.table12Txbl, Math.abs(b.table12Txbl - b.salesSheetTxbl) < 1 ? "equals the sales sheet" : "DOES NOT equal the sales sheet — investigate");
    rc.addRow([]);
    sectionTitle(rc, "D", "2 · Cross-check: do the filed B2B invoices match the e-invoice portal?");
    brRow("B2B invoices being filed (from the books)", b.jsonB2bTxbl, "Table 4 / JSON b2b", "base");
    if (!b.crossCheck.length) brRow("No differences", 0, "books and portal agree invoice-for-invoice");
    b.crossCheck.forEach((row) => brRow(row.label, row.amt, row.note, "warn"));
    brRow("= Valid e-invoices on the portal", b.portalValidTxbl, "from the EINV dump", "result");
    brRow("Unexplained difference", b.residual, Math.abs(b.residual) < 1 ? "fully explained" : "INVESTIGATE", "residual");
  }

  // ===== Sheet 4: TABLE 12 HSN =====
  const h12 = wb.addWorksheet(SH.t12, { views: [{ state: "frozen", ySplit: 4, showGridLines: false }] });
  h12.columns = [{ width: 5 }, { width: 12 }, { width: 38 }, { width: 7 }, { width: 12 }, { width: 8 }, { width: 17 }, { width: 15 }, { width: 14 }, { width: 14 }, { width: 11 }, { width: 17 }];
  banner(h12, "L", "TABLE 12 — HSN / SAC SUMMARY (from the sales sheet)", "Taxable/tax/qty from the books; UQC & description enriched from e-invoice support.");
  headerRow(h12, ["Sr", "HSN/SAC", "Description", "UQC", "Qty", "Rate %", "Taxable", "IGST", "CGST", "SGST", "Cess", "Total"]);
  res.table12.forEach((h, i) => {
    const r = dataRow(h12, [h.sr, h.hsn, h.desc, h.uqc, h.qty, h.rt, h.txval, h.iamt, h.camt, h.samt, h.csamt, h.total], i, [7, 8, 9, 10, 11, 12]);
    r.getCell(3).alignment = { wrapText: true }; r.getCell(5).numFmt = "#,##0"; r.getCell(6).numFmt = "0.00";
    if (!h.desc) r.getCell(3).fill = fill(AMBER); // flag blank description (C4)
    if (!/^\d{6}$|^\d{8}$/.test(String(h.hsn))) r.getCell(2).fill = fill(RED); // flag bad HSN code (C1)
  });
  const tr = h12.addRow(["", "TOTAL", "", "", "", "", t.t12Txbl, t.t12Igst, t.t12Cgst, t.t12Sgst, t.t12Cess, totalVal]);
  tr.eachCell((c) => { c.font = { bold: true, color: { argb: INK } }; c.fill = fill(LIGHT); c.border = border; });
  [7, 8, 9, 10, 11, 12].forEach((col) => (tr.getCell(col).numFmt = MONEY));

  // ===== Sheet 5: TABLE 13 DOCS =====
  const d13 = wb.addWorksheet(SH.t13, { views: [{ showGridLines: false }] });
  d13.columns = [{ width: 7 }, { width: 34 }, { width: 16 }, { width: 16 }, { width: 10 }, { width: 11 }, { width: 10 }];
  banner(d13, "G", "TABLE 13 — DOCUMENTS ISSUED", `Return Period ${res.meta.period || res.meta.fp}`);
  headerRow(d13, ["Code", "Nature of Document", "From", "To", "Total", "Cancelled", "Net"]);
  res.table13.forEach((d, i) => dataRow(d13, [d.code, d.nature, d.from, d.to, d.total, d.cancel, d.net], i));

  // ===== Sheet 6: CANCELLED INVOICES =====
  const cs = wb.addWorksheet(SH.canc, { views: [{ showGridLines: false }] });
  cs.columns = [{ width: 16 }, { width: 40 }, { width: 20 }, { width: 18 }, { width: 16 }, { width: 40 }];
  banner(cs, "F", "CANCELLED INVOICES — confirm with client", "In the sales sheet (so in Table 12) but IRN cancelled on the portal. Remove if the sale didn't happen / was re-billed.");
  headerRow(cs, ["Invoice", "Recipient", "GSTIN", "Book taxable", "Portal status", "Action"]);
  (res.cancelledDetail || []).forEach((c, i) => {
    const r = dataRow(cs, [c.inum, c.name, c.gstin, c.txbl, c.portalStatus || "cancelled", "Confirm: remove from GSTR-1 or keep (re-billed?)"], i, [4]);
    r.getCell(5).fill = fill(AMBER);
  });
  if ((res.cancelledDetail || []).length) {
    const r = cs.addRow(["", "TOTAL", "", cancTot, "", ""]);
    r.eachCell((c) => { c.font = { bold: true, color: { argb: INK } }; c.fill = fill(LIGHT); c.border = border; });
    r.getCell(4).numFmt = MONEY;
  } else cs.addRow(["—", "No cancelled invoices.", "", "", "", ""]);

  // ===== Sheet 7: DEVIATIONS =====
  const dv = wb.addWorksheet(SH.dev, { views: [{ showGridLines: false }] });
  dv.columns = [{ width: 16 }, { width: 40 }, { width: 18 }, { width: 18 }, { width: 18 }];
  banner(dv, "E", "DEVIATIONS — client book vs portal e-invoice", "Confirm the correct figure with the client for each.");
  sectionTitle(dv, "E", "Per-invoice value mismatches (C14)");
  headerRow(dv, ["Invoice", "Recipient", "Sales book", "Portal", "Difference"]);
  if (res.recon.valMismatch.length) res.recon.valMismatch.forEach((m, i) => { const r = dataRow(dv, [m.inum, m.name, m.sales, m.portal, m.diff], i, [3, 4, 5]); r.getCell(5).font = { bold: true, color: { argb: "FF991B1B" } }; });
  else dv.addRow(["—", "No value mismatches.", "", "", ""]);
  dv.addRow([]);
  sectionTitle(dv, "E", "e-Invoices missing from the sales book (C13)");
  if (res.recon.missingInSales.length) res.recon.missingInSales.forEach((n, i) => dataRow(dv, [n, "missing in sales sheet", "", "", ""], i));
  else dv.addRow(["—", "None missing.", "", "", ""]);
  dv.addRow([]);
  sectionTitle(dv, "E", "Tagged 'Non-Revenue' but e-invoiced — INCLUDED in GSTR-1");
  headerRow(dv, ["Invoice", "Recipient", "Portal taxable", "Client tag", ""]);
  if (res.recon.nonRevButEinvoiced.length) res.recon.nonRevButEinvoiced.forEach((m, i) => dataRow(dv, [m.inum, m.name, m.txbl, m.rev, ""], i, [3]));
  else dv.addRow(["—", "None.", "", "", ""]);

  // ===== Sheet 8: CHECKS =====
  const ck = wb.addWorksheet(SH.checks, { views: [{ state: "frozen", ySplit: 4, showGridLines: false }] });
  ck.columns = [{ width: 6 }, { width: 56 }, { width: 28 }, { width: 42 }, { width: 12 }];
  banner(ck, "E", "VALIDATION CHECKS", `${res.checks.filter((c) => c.ok).length}/${res.checks.length} passed · ${res.blockingFails.length} blocking`);
  headerRow(ck, ["#", "Check", "Expected", "Actual", "Result"]);
  res.checks.forEach((c, i) => {
    const verdict = c.ok ? "PASS" : (res.blockingFails.includes(c) ? "FAIL" : "DEVIATION");
    const r = dataRow(ck, [c.id, c.desc, c.expected, c.actual, verdict], i);
    r.getCell(5).fill = fill(c.ok ? GREEN : (res.blockingFails.includes(c) ? RED : AMBER));
    r.getCell(5).font = { bold: true };
    [2, 3, 4].forEach((col) => (r.getCell(col).alignment = { wrapText: true, vertical: "top" }));
  });

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function BridgeBlock({ b }) {
  return (
    <>
      <div className="g1-cli-sec" style={{ marginTop: 0 }}>1 · What's in the client sales sheet, and where each piece goes</div>
      <div className="g1-cli-wrap" style={{ marginTop: 0 }}>
        <table className="g1-table g1-br">
          <tbody>
            {b.composition.map((r) => (
              <tr key={r.label}>
                <td className="left">{r.label}</td>
                <td className="num">{inr(r.amt)}</td>
                <td className="left small">{r.note}</td>
              </tr>
            ))}
            <tr className={Math.abs(b.residual0) < 1 ? "row-total" : "row-bad"}>
              <td className="left">Total = Client sales sheet taxable</td>
              <td className="num">{inr(b.salesSheetTxbl)}</td>
              <td className="left small">{Math.abs(b.residual0) < 1 ? "✓ every document accounted for" : `⚠ ${inr(b.residual0)} unaccounted — investigate`}</td>
            </tr>
            <tr>
              <td className="left">Table 12 (HSN summary in the return)</td>
              <td className="num">{inr(b.table12Txbl)}</td>
              <td className="left small">{Math.abs(b.table12Txbl - b.salesSheetTxbl) < 1 ? "✓ equals the sales sheet" : "⚠ does not equal the sales sheet"}</td>
            </tr>
          </tbody>
        </table>
      </div>
      <div className="g1-cli-sec">2 · Cross-check: do the filed B2B invoices match the e-invoice portal?</div>
      <div className="g1-cli-wrap" style={{ marginTop: 0 }}>
        <table className="g1-table g1-br">
          <tbody>
            <tr className="row-total"><td className="left">B2B invoices being filed (from the books)</td><td className="num">{inr(b.jsonB2bTxbl)}</td><td className="left small">Table 4 / JSON b2b</td></tr>
            {b.crossCheck.length === 0 && (
              <tr><td className="left">No differences</td><td className="num">—</td><td className="left small">books and portal agree invoice-for-invoice</td></tr>
            )}
            {b.crossCheck.map((r) => (
              <tr key={r.label} className="row-warn">
                <td className="left">{r.label}</td>
                <td className="num">{inr(r.amt)}</td>
                <td className="left small">{r.note}</td>
              </tr>
            ))}
            <tr className="row-total"><td className="left">= Valid e-invoices on the portal</td><td className="num">{inr(b.portalValidTxbl)}</td><td className="left small">from the EINV dump</td></tr>
            <tr className={Math.abs(b.residual) < 1 ? "" : "row-bad"}>
              <td className="left">Unexplained difference</td>
              <td className="num">{inr(b.residual)}</td>
              <td className="left small">{Math.abs(b.residual) < 1 ? "✓ fully explained" : "⚠ investigate"}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </>
  );
}

function Fold({ title, badge, open, children }) {
  return (
    <details className="g1-fold" {...(open ? { open: true } : {})}>
      <summary>{title}{badge != null ? <span className="badge">{badge}</span> : null}</summary>
      <div className="g1-fold-body">{children}</div>
    </details>
  );
}

function ClientSummary({ res, onDownload, busy }) {
  const t = res.totals;
  const totalTax = (t.t12Igst || 0) + (t.t12Cgst || 0) + (t.t12Sgst || 0) + (t.t12Cess || 0);
  const totalVal = (t.t12Txbl || 0) + totalTax;
  const cls = res.portalReady ? (res.allPass ? "ok" : "warn") : "bad";
  const statusTxt = res.portalReady ? (res.allPass ? "✓ Ready to file" : "⚠ Review deviations") : "✗ Blocking issues";
  const rollup = useMemo(() => {
    const m = {};
    for (const h of res.table12) {
      const g = (m[h.rt] ||= { rt: h.rt, lines: 0, txval: 0, iamt: 0, camt: 0, samt: 0, csamt: 0 });
      g.lines++; g.txval += h.txval; g.iamt += h.iamt; g.camt += h.camt; g.samt += h.samt; g.csamt += h.csamt;
    }
    return Object.values(m).sort((a, b) => a.rt - b.rt);
  }, [res.table12]);
  const kpi = (label, value, sub, accent) => (
    <div className={`g1-cli-kpi ${accent ? "accent" : ""}`} key={label}>
      <span className="k">{label}</span><span className="v">{value}</span>{sub ? <span className="s">{sub}</span> : null}
    </div>
  );
  return (
    <div className="g1-cli">
      <div className="g1-cli-hero">
        <div>
          <h3>GSTR-1 Summary — {res.meta.period || res.meta.fp}</h3>
          <p>{res.meta.supplierName ? res.meta.supplierName + " · " : ""}GSTIN {res.meta.gstin || "—"} · {res.counts.jsonB2bInvoices} invoices · {res.counts.hsnRows} HSN lines</p>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <span className={`g1-cli-status ${cls}`}>{statusTxt}</span>
          <button className="g1-cli-dl" onClick={onDownload} disabled={busy}>{busy ? "Building…" : "⬇ Download formatted Excel"}</button>
        </div>
      </div>

      {!res.portalReady && (
        <div className="g1-cli-block bad">
          <b>⛔ JSON generation is BLOCKED.</b> Fix these before the GSTR-1 JSON can be downloaded:
          <ul>{res.blockingFails.map((c) => (<li key={c.id}><b>{c.id}</b> — {c.desc}. <span className="g1-cli-found">Found: {c.actual}</span></li>))}</ul>
        </div>
      )}
      {res.portalReady && res.fails.length > 0 && (
        <div className="g1-cli-block warn">
          <b>⚠ {res.fails.length} item(s) to confirm with the client</b> (JSON can still generate): {res.fails.map((c) => c.id).join(", ")}. See the <b>Deviations</b> tab / Excel <b>Action Points</b> sheet.
        </div>
      )}

      <div className="g1-cli-kpis">
        {kpi("Taxable Value", inr0(t.t12Txbl))}
        {kpi("Total Tax", inr0(totalTax))}
        {kpi("Total Invoice Value", inr0(totalVal), "taxable + tax", true)}
        {kpi("Invoices in JSON", res.counts.jsonB2bInvoices, `${res.counts.cdnr} CN/DN · ${res.counts.zeroDocs} zero-value excluded`)}
      </div>
      <div className="g1-cli-split">
        <span><b>IGST</b> {inr0(t.t12Igst)}</span>
        <span><b>CGST</b> {inr0(t.t12Cgst)}</span>
        <span><b>SGST</b> {inr0(t.t12Sgst)}</span>
        <span><b>Cess</b> {inr0(t.t12Cess)}</span>
      </div>

      {res.bridge && (
        <Fold open title="Sales sheet → GSTR-1 reconciliation" badge={Math.abs(res.bridge.residual) < 1 && Math.abs(res.bridge.residual0) < 1 ? "✓ tied" : "⚠ check"}>
          <BridgeBlock b={res.bridge} />
        </Fold>
      )}

      <div className="g1-cli-sec">HSN / SAC Summary <span className="badge">{rollup.length} rates · {res.table12.length} lines</span></div>
      <div className="g1-cli-wrap">
        <table className="g1-table">
          <thead><tr><th className="num">Rate %</th><th className="num">Lines</th><th className="num">Taxable</th><th className="num">IGST</th><th className="num">CGST</th><th className="num">SGST</th><th className="num">Cess</th></tr></thead>
          <tbody>
            {rollup.map((g) => (
              <tr key={g.rt}><td className="num">{g.rt}</td><td className="num">{g.lines}</td><td className="num">{inr(g.txval)}</td><td className="num">{inr(g.iamt)}</td><td className="num">{inr(g.camt)}</td><td className="num">{inr(g.samt)}</td><td className="num">{inr(g.csamt)}</td></tr>
            ))}
            <tr className="row-total"><td className="num">TOTAL</td><td className="num">{res.table12.length}</td><td className="num">{inr(t.t12Txbl)}</td><td className="num">{inr(t.t12Igst)}</td><td className="num">{inr(t.t12Cgst)}</td><td className="num">{inr(t.t12Sgst)}</td><td className="num">{inr(t.t12Cess)}</td></tr>
          </tbody>
        </table>
      </div>

      <Fold title="View all HSN/SAC lines" badge={res.table12.length}>
        <div className="g1-cli-wrap" style={{ marginTop: 0 }}>
          <table className="g1-table">
            <thead><tr><th>Sr</th><th>HSN/SAC</th><th className="left">Description</th><th>UQC</th><th className="num">Qty</th><th className="num">Rate%</th><th className="num">Taxable</th><th className="num">IGST</th><th className="num">CGST</th><th className="num">SGST</th><th className="num">Cess</th></tr></thead>
            <tbody>
              {res.table12.map((h) => (
                <tr key={h.sr + h.hsn + h.rt}><td>{h.sr}</td><td>{h.hsn}</td><td className="left small">{h.desc}</td><td>{h.uqc}</td><td className="num">{Number(h.qty).toLocaleString("en-IN")}</td><td className="num">{h.rt}</td><td className="num">{inr(h.txval)}</td><td className="num">{inr(h.iamt)}</td><td className="num">{inr(h.camt)}</td><td className="num">{inr(h.samt)}</td><td className="num">{inr(h.csamt)}</td></tr>
              ))}
            </tbody>
          </table>
        </div>
      </Fold>

      <Fold title="Documents issued (Table 13)" badge={res.table13.length}>
        <div className="g1-cli-wrap" style={{ marginTop: 0 }}>
          <table className="g1-table">
            <thead><tr><th>Code</th><th className="left">Nature of Document</th><th>From</th><th>To</th><th className="num">Total</th><th className="num">Cancelled</th><th className="num">Net</th></tr></thead>
            <tbody>{res.table13.map((d) => (<tr key={d.code}><td>{d.code}</td><td className="left">{d.nature}</td><td>{d.from}</td><td>{d.to}</td><td className="num">{d.total}</td><td className="num">{d.cancel}</td><td className="num">{d.net}</td></tr>))}</tbody>
          </table>
        </div>
      </Fold>

      <p className="g1-cli-note">Bird's-eye view — figures match the GSTR-1 JSON. HSN shown rate-wise; expand for line detail. <b>Download formatted Excel</b> for the full client workbook with action points.</p>
    </div>
  );
}

function download(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

// Portal-schema validation of a single HSN row → { field: errorMessage }.
// These are the rules that make the GST portal REJECT the JSON.
function validateHsnRow(row) {
  const errs = {};
  const hsn = String(row.hsn_sc ?? "").trim();
  const isSac = /^99/.test(hsn);
  const uqc = String(row.uqc ?? "").trim().toUpperCase();
  if (!/^\d{6}$|^\d{8}$/.test(hsn)) errs.hsn_sc = "HSN/SAC must be numeric, 6 or 8 digits";
  if (!VALID_RATES.includes(Number(row.rt))) errs.rt = "Not a valid GST rate slab";
  if (!VALID_UQC.has(uqc)) errs.uqc = "UQC not a portal-standard code";
  else if (isSac && uqc !== "NA") errs.uqc = "Service SAC (99…) must use UQC NA (RET191353)";
  else if (!isSac && uqc === "NA") errs.uqc = "Goods must use a real UQC, not NA";
  if (isSac && Number(row.qty) !== 0) errs.qty = "Service SAC qty must be 0 (RET191355)";
  if (!isSac && !(Number(row.qty) > 0)) errs.qty = "Goods qty must be greater than 0";
  if (!String(row.desc ?? "").trim()) errs.desc = "Description is required (Phase-3 mandatory)";
  return errs;
}

const HSN_COLS = [
  { f: "hsn_sc", label: "HSN/SAC", w: 96 }, { f: "desc", label: "Description", w: 240 },
  { f: "uqc", label: "UQC", w: 64 }, { f: "qty", label: "Qty", w: 84, num: true },
  { f: "rt", label: "Rate %", w: 64, num: true }, { f: "txval", label: "Taxable", w: 120, num: true },
  { f: "iamt", label: "IGST", w: 104, num: true }, { f: "camt", label: "CGST", w: 96, num: true },
  { f: "samt", label: "SGST", w: 96, num: true }, { f: "csamt", label: "Cess", w: 80, num: true },
];

const DOC_COLS = [
  { f: "kind", label: "Type", w: 52, ro: true },
  { f: "inum", label: "Doc #", w: 96, ro: true },
  { f: "name", label: "Recipient", w: 200, ro: true },
  { f: "ctin", label: "GSTIN", w: 168 },
  { f: "idt", label: "Date (dd-mm-yyyy)", w: 116 },
  { f: "pos", label: "POS", w: 52 },
  { f: "invTyp", label: "Inv type", w: 72 },
  { f: "rchrg", label: "RCM", w: 48 },
  { f: "val", label: "Doc value", w: 110, num: true },
];

// Editable grids: fix invalid document fields and HSN cells, see edits in red,
// then generate a JSON guaranteed to pass portal-schema validation.
function EditableHsn({ res }) {
  const seedHsn = () => {
    const h = res.json?.hsn || {};
    if (Array.isArray(h.data)) return h.data.map((r) => ({ ...r, _seg: "data" }));
    return [
      ...(h.hsn_b2b || []).map((r) => ({ ...r, _seg: "b2b" })),
      ...(h.hsn_b2c || []).map((r) => ({ ...r, _seg: "b2c" })),
    ];
  };
  const seedDocs = () => (res.jsonDocs || []).map((d) => ({ ...d }));
  const [rows, setRows] = useState(seedHsn);
  const [docs, setDocs] = useState(seedDocs);
  const [edited, setEdited] = useState(() => new Set());
  const [done, setDone] = useState(false);
  const [onlyBad, setOnlyBad] = useState(true);
  // Reseed happens via a `key` prop on the parent (remounts on each new run).

  const supplierState = res.meta.supplierState || (res.meta.gstin || "").slice(0, 2);
  const fp = res.meta.fp;
  const errorsByRow = useMemo(() => rows.map(validateHsnRow), [rows]);
  const docErrsByRow = useMemo(() => docs.map((d) => validateDocRow(d, { supplierState, fp })), [docs, supplierState, fp]);
  const hsnErrors = errorsByRow.reduce((a, e) => a + Object.keys(e).length, 0);
  const docErrors = docErrsByRow.reduce((a, e) => a + Object.keys(e).length, 0);
  const totalErrors = hsnErrors + docErrors;
  const badDocCount = docErrsByRow.filter((e) => Object.keys(e).length).length;

  const setCell = (i, f, v) => {
    setRows((prev) => prev.map((r, idx) => (idx === i ? { ...r, [f]: v } : r)));
    setEdited((prev) => new Set(prev).add(`h${i}:${f}`));
    setDone(false);
  };
  const setDocCell = (i, f, v) => {
    setDocs((prev) => prev.map((r, idx) => (idx === i ? { ...r, [f]: f === "val" ? v : String(v).toUpperCase() } : r)));
    setEdited((prev) => new Set(prev).add(`d${i}:${f}`));
    setDone(false);
  };

  const generate = () => {
    // Rebuild b2b/cdnr from the (possibly edited) documents — same builder as the engine,
    // so the exported JSON always corresponds 1:1 with the client's books + the fixes made here.
    const normDocs = docs.map((d) => ({ ...d, ctin: String(d.ctin).trim().toUpperCase(), val: Number(d.val) || 0 }));
    const json = buildJson({
      gstin: res.meta.gstin, fp: res.meta.fp,
      b2bDocs: normDocs.filter((d) => d.kind !== "CN"),
      cnDocs: normDocs.filter((d) => d.kind === "CN"),
      t12: [], t12B2b: [], t12B2c: [], t13: res.table13,
      version: res.json?.version || "GST3.1.2",
    });
    const mk = (arr) => arr.map((r, i) => ({
      num: i + 1, hsn_sc: String(r.hsn_sc).trim(), desc: String(r.desc || "").slice(0, 30),
      uqc: String(r.uqc || "OTH").toUpperCase(), qty: Number(r.qty) || 0, txval: Number(r.txval) || 0,
      rt: Number(r.rt) || 0, iamt: Number(r.iamt) || 0, camt: Number(r.camt) || 0,
      samt: Number(r.samt) || 0, csamt: Number(r.csamt) || 0,
    }));
    const seg = {};
    for (const r of rows) (seg[r._seg] ||= []).push(r);
    json.hsn = seg.data ? { data: mk(seg.data) } : { hsn_b2b: mk(seg.b2b || []), hsn_b2c: mk(seg.b2c || []) };
    download(new Blob([JSON.stringify(json, null, 2)], { type: "application/json" }),
      `GSTR1_${res.meta.gstin || "return"}_${res.meta.fp}_EDITED.json`);
    setDone(true);
  };

  const visibleDocs = docs.map((d, i) => [d, i]).filter(([, i]) => !onlyBad || Object.keys(docErrsByRow[i]).length);

  return (
    <div className="g1-edit-view">
      <p className="g1-cli-note">
        Everything below comes from the <b>client Sales sheet</b> — fix whatever the portal would reject
        (invalid GSTIN/date/POS on a document, bad HSN code, blank description, wrong UQC, …).
        <b> Cells you change turn red.</b> Cells that are still invalid get a red outline — hover to see why.
        The <b>Generate JSON</b> button enables once every document and HSN row passes portal-schema validation.
      </p>
      <div className="g1-edit-bar">
        {totalErrors === 0
          ? <span className="g1-edit-ok">✓ All {docs.length} documents and {rows.length} HSN rows valid — ready to generate</span>
          : <span className="g1-edit-err">⛔ {totalErrors} invalid cell(s): {docErrors} on {badDocCount} document(s) · {hsnErrors} on {errorsByRow.filter((e) => Object.keys(e).length).length} HSN row(s)</span>}
        <button className="g1-edit-gen" onClick={generate} disabled={totalErrors > 0}>⬇ Generate GSTR-1 JSON</button>
        <button className="g1-dl" onClick={() => { setRows(seedHsn()); setDocs(seedDocs()); setEdited(new Set()); setDone(false); }}>↺ Reset to original</button>
        {done && <span className="g1-edit-ok">JSON downloaded ✓</span>}
      </div>

      <div className="g1-cli-sec">Documents — B2B invoices &amp; credit notes <span className="badge">{docs.length} docs · {badDocCount} with errors</span>
        <label style={{ marginLeft: 12, fontWeight: 400, fontSize: 12 }}>
          <input type="checkbox" checked={onlyBad} onChange={(e) => setOnlyBad(e.target.checked)} /> show only documents with errors
        </label>
      </div>
      {visibleDocs.length === 0 ? (
        <p className="g1-cli-note">✓ {onlyBad ? "No document-level errors." : "No documents."} {onlyBad && docs.length ? "Untick the filter to browse all documents." : ""}</p>
      ) : (
        <div className="g1-edit-wrap">
          <table className="g1-edit">
            <thead>
              <tr>
                <th>#</th>
                {DOC_COLS.map((c) => <th key={c.f} style={{ minWidth: c.w }}>{c.label}</th>)}
                <th>Rate-wise items</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {visibleDocs.map(([row, i]) => {
                const errs = docErrsByRow[i];
                const firstErr = Object.values(errs)[0];
                return (
                  <tr key={row.kind + row.inum}>
                    <td className="g1-edit-rn">{i + 1}</td>
                    {DOC_COLS.map((c) => {
                      const isEdited = edited.has(`d${i}:${c.f}`);
                      const isBad = !!errs[c.f];
                      return (
                        <td key={c.f} className={`${c.num ? "num" : ""} ${isEdited ? "edited" : ""} ${isBad ? "bad" : ""}`} title={errs[c.f] || ""}>
                          {c.ro
                            ? <span className="g1-edit-ro">{row[c.f]}</span>
                            : <input value={row[c.f] ?? ""} onChange={(e) => setDocCell(i, c.f, e.target.value)} inputMode={c.num ? "decimal" : "text"} />}
                        </td>
                      );
                    })}
                    <td className={`g1-edit-ro small ${errs.items ? "bad" : ""}`} title={errs.items || ""}>
                      {(row.items || []).map((it) => `${it.rt}%: ${inr(it.txval)}`).join(" · ")}
                    </td>
                    <td className="g1-edit-status">{firstErr ? <span className="g1-edit-rowerr" title={Object.values(errs).join("; ")}>⚠ {firstErr}</span> : <span className="g1-edit-ok">✓</span>}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      <div className="g1-cli-sec" style={{ marginTop: 16 }}>Table 12 — HSN / SAC rows <span className="badge">{rows.length} rows</span></div>
      <div className="g1-edit-wrap">
        <table className="g1-edit">
          <thead>
            <tr>
              <th>#</th>
              {HSN_COLS.map((c) => <th key={c.f} style={{ minWidth: c.w }}>{c.label}</th>)}
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, i) => {
              const errs = errorsByRow[i];
              const firstErr = Object.values(errs)[0];
              return (
                <tr key={i}>
                  <td className="g1-edit-rn">{i + 1}</td>
                  {HSN_COLS.map((c) => {
                    const isEdited = edited.has(`h${i}:${c.f}`);
                    const isBad = !!errs[c.f];
                    return (
                      <td key={c.f} className={`${c.num ? "num" : ""} ${isEdited ? "edited" : ""} ${isBad ? "bad" : ""}`} title={errs[c.f] || ""}>
                        <input value={row[c.f] ?? ""} onChange={(e) => setCell(i, c.f, e.target.value)} inputMode={c.num ? "decimal" : "text"} />
                      </td>
                    );
                  })}
                  <td className="g1-edit-status">{firstErr ? <span className="g1-edit-rowerr" title={Object.values(errs).join("; ")}>⚠ {firstErr}</span> : <span className="g1-edit-ok">✓</span>}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

export default function Gstr1Processor() {
  const [einvFile, setEinvFile] = useState(null);
  const [salesFile, setSalesFile] = useState(null);
  const [month, setMonth] = useState(4); // May (0-based)
  const [year, setYear] = useState(2026);
  const [version, setVersion] = useState("GST3.1.2"); // JSON schema version; match your Offline Tool if portal rejects
  const [running, setRunning] = useState(false);
  const [res, setRes] = useState(null);
  const [err, setErr] = useState("");
  const [tab, setTab] = useState("client");
  const [xlBusy, setXlBusy] = useState(false);

  const fp = useMemo(() => String(month + 1).padStart(2, "0") + year, [month, year]);

  const run = async () => {
    setErr("");
    if (!einvFile || !salesFile) { setErr("Select both files: the e-invoice dump and the client GSTR-1 working file."); return; }
    setRunning(true); setRes(null);
    try {
      const [eBuf, sBuf] = await Promise.all([einvFile.arrayBuffer(), salesFile.arrayBuffer()]);
      const einvWb = XLSX.read(eBuf, { type: "array", cellDates: true });
      const salesWb = XLSX.read(sBuf, { type: "array", cellDates: true });
      const gstin = (einvFile.name.match(GSTIN_RE) || [])[1] || "";
      const out = processGstr1(einvWb, salesWb, { gstin, fp, version: version.trim() || "GST3.1.2", periodLabel: `${MONTHS[month]} ${year}` });
      if (out.errors.length) setErr(out.errors.join(" "));
      setRes(out);
    } catch (e) {
      console.error(e); setErr(e.message || String(e));
    } finally { setRunning(false); }
  };

  const dlJson = () => download(new Blob([JSON.stringify(res.json, null, 2)], { type: "application/json" }), `GSTR1_${res.meta.gstin || "return"}_${res.meta.fp}.json`);
  const dlExcel = async () => download(await buildExcel(res), `GSTR1_validation_${res.meta.gstin || "SIPL"}_${res.meta.fp}.xlsx`);
  const dlClientExcel = async () => { setXlBusy(true); try { download(await buildClientExcel(res), `GSTR1_Client_Summary_${res.meta.gstin || "SIPL"}_${res.meta.fp}.xlsx`); } finally { setXlBusy(false); } };

  return (
    <div className="g1">
      <section className="g1-intro">
        <h2>GSTR-1 Processor</h2>
        <p>
          Builds the <b>entire GSTR-1 JSON — B2B invoices, credit/debit notes, Table 12 (HSN) and Table 13
          (Documents issued) — from the client Sales sheet. The books are the base; every number in the JSON
          corresponds 1:1 with the client's data.</b> The GST-portal e-invoice dump is <b>support only</b>: it enriches
          UQC/description/invoice-type and cross-validates every deviation, but never feeds a number into the JSON.
          Validation errors are fixable in the <b>Fix &amp; Generate</b> tab before export. All processing is local —
          nothing leaves your browser.
        </p>
        <ol className="g1-steps">
          <li>Upload the client's <b>GSTR-1 working file</b> — the base the JSON is built from.</li>
          <li>Upload the GST-portal <b>e-invoice dump</b> (EINV_&lt;gstin&gt;_&lt;fy&gt;.xlsx) — the cross-check.</li>
          <li>Pick the return month. Run. Fix any validation errors in-app, resolve deviations with the client, download the JSON.</li>
        </ol>
      </section>

      <section className="g1-card">
        <FilePick label="① Client GSTR-1 working (.xlsx)" hint="GSTR-1 Data <month> <yy>.xlsx shared by the client — the base" accept=".xlsx" file={salesFile} onFile={setSalesFile} />
        <FilePick label="② E-invoice dump (.xlsx)" hint="EINV_<gstin>_<fy>.xlsx from the GST portal — the cross-check" accept=".xlsx" file={einvFile} onFile={setEinvFile} />
        <div className="g1-period">
          <label>Return period</label>
          <select value={month} onChange={(e) => setMonth(+e.target.value)}>{MONTHS.map((m, i) => <option key={m} value={i}>{m}</option>)}</select>
          <input type="number" value={year} onChange={(e) => setYear(+e.target.value)} style={{ width: 90 }} />
          <span className="g1-fp">fp = {fp}</span>
          <label style={{ marginLeft: 16 }}>JSON version</label>
          <input value={version} onChange={(e) => setVersion(e.target.value)} style={{ width: 110 }} title="Schema version written into the JSON. If the portal rejects, match the value your current Offline Tool emits." />
        </div>
        <p className="g1-ov-note" style={{ margin: "2px 0 0 240px" }}>
          HSN section auto-bifurcates into <code>hsn_b2b</code>/<code>hsn_b2c</code> for periods from May-2025 (portal rule). If the portal still rejects, open a JSON the current Offline Tool generates and copy its <code>version</code> string here.
        </p>
        <button className="g1-run" onClick={run} disabled={running}>{running ? "Processing…" : "Process GSTR-1"}</button>
        {err && <p className="g1-err">{err}</p>}
      </section>

      {res && (
        <>
          <Banner res={res} />
          <FileOverview res={res} />
          <Cards res={res} />
          <section className="g1-card">
            <div className="g1-actions">
              <button className="g1-dl primary" onClick={dlJson} disabled={!res.portalReady} title={res.portalReady ? "" : "Resolve blocking failures first"}>⬇ Download GSTR-1 JSON</button>
              <button className="g1-dl" onClick={dlExcel}>⬇ Download validation workbook (.xlsx)</button>
            </div>
            <div className="g1-tabs">
              {[["client", "🦅 Client Summary"], ["edit", "✏️ Fix & Generate JSON"], ["dev", "Deviations & Reconciliation"], ["t12", "Table 12 (HSN)"], ["t13", "Table 13 (Docs)"], ["checks", "Validation Checks"], ["json", "JSON Preview"]].map(([id, l]) => (
                <button key={id} className={tab === id ? "active" : ""} onClick={() => setTab(id)}>{l}</button>
              ))}
            </div>

            {tab === "client" && <ClientSummary res={res} onDownload={dlClientExcel} busy={xlBusy} />}
            {tab === "edit" && <EditableHsn key={`${res.meta.fp}:${res.counts.hsnRows}:${res.counts.salesRows}`} res={res} />}
            {tab === "dev" && <Deviations res={res} />}
            {tab === "checks" && <ChecksTable res={res} />}
            {tab === "t12" && (
              <table className="g1-table"><thead><tr><th>Sr</th><th>HSN/SAC</th><th className="left">Description</th><th>UQC</th><th className="num">Qty</th><th className="num">Rate%</th><th className="num">Taxable</th><th className="num">IGST</th><th className="num">CGST</th><th className="num">SGST</th><th className="num">Cess</th></tr></thead>
                <tbody>{res.table12.map((h) => (<tr key={h.sr + h.hsn + h.rt}><td>{h.sr}</td><td>{h.hsn}</td><td className="left small">{h.desc}</td><td>{h.uqc}</td><td className="num">{Number(h.qty).toLocaleString("en-IN")}</td><td className="num">{h.rt}</td><td className="num">{inr(h.txval)}</td><td className="num">{inr(h.iamt)}</td><td className="num">{inr(h.camt)}</td><td className="num">{inr(h.samt)}</td><td className="num">{inr(h.csamt)}</td></tr>))}
                  <tr className="row-total"><td></td><td>TOTAL</td><td></td><td></td><td></td><td></td><td className="num">{inr(res.totals.t12Txbl)}</td><td className="num">{inr(res.totals.t12Igst)}</td><td className="num">{inr(res.totals.t12Cgst)}</td><td className="num">{inr(res.totals.t12Sgst)}</td><td className="num">{inr(res.totals.t12Cess)}</td></tr>
                </tbody></table>
            )}
            {tab === "t13" && (
              <table className="g1-table"><thead><tr><th>Code</th><th className="left">Nature of Document</th><th>From</th><th>To</th><th className="num">Total</th><th className="num">Cancelled</th><th className="num">Net</th></tr></thead>
                <tbody>{res.table13.map((d) => (<tr key={d.code}><td>{d.code}</td><td className="left">{d.nature}</td><td>{d.from}</td><td>{d.to}</td><td className="num">{d.total}</td><td className="num">{d.cancel}</td><td className="num">{d.net}</td></tr>))}</tbody></table>
            )}
            {tab === "json" && <pre className="g1-json">{JSON.stringify(res.json, null, 2)}</pre>}
          </section>
        </>
      )}
    </div>
  );
}
