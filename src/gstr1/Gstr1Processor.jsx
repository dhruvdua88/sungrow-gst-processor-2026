import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { processGstr1 } from "./engine.js";
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
    ["Valid e-invoices", c.validInvoices, true], ["Cancelled IRNs", c.cancelled, true], ["Credit notes", c.cdnr, true], ["HSN rows", c.hsnRows, true],
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
        <p className="g1-note"><b>Cancelled IRNs (excluded from GSTR-1):</b> {res.cancelled.join(", ")}</p>
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

// Beautifully formatted, client-facing workbook (distinct from the internal validation one).
async function buildClientExcel(res) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Sungrow GST Processor 2026";
  const BRAND = "FF1E4E8C", ACCENT = "FF0EA5E9", LIGHT = "FFEFF6FF", BAND = "FFF6FAFF", INK = "FF0F2C52";
  const ws = wb.addWorksheet("GSTR-1 Summary", { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 5 }, { width: 12 }, { width: 40 }, { width: 7 }, { width: 12 }, { width: 8 },
    { width: 17 }, { width: 15 }, { width: 14 }, { width: 14 }, { width: 11 }, { width: 17 },
  ];
  const MONEY = "#,##0.00";
  const thin = { style: "thin", color: { argb: "FFD9E2EC" } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const fill = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
  const COL = (n) => ws.getColumn(n).letter;

  const t = res.totals;
  const totalTax = (t.t12Igst || 0) + (t.t12Cgst || 0) + (t.t12Sgst || 0) + (t.t12Cess || 0);
  const totalVal = (t.t12Txbl || 0) + totalTax;

  // Title + subtitle banner
  ws.mergeCells("A1:L1");
  const title = ws.getCell("A1");
  title.value = "GSTR-1  ·  OUTWARD SUPPLIES SUMMARY";
  title.font = { bold: true, size: 18, color: { argb: "FFFFFFFF" } };
  title.alignment = { vertical: "middle", indent: 1 };
  title.fill = fill(BRAND);
  ws.getRow(1).height = 34;
  ws.mergeCells("A2:L2");
  const sub = ws.getCell("A2");
  sub.value = `${res.meta.supplierName ? res.meta.supplierName + "   ·   " : ""}GSTIN ${res.meta.gstin || "—"}      Return Period: ${res.meta.period || res.meta.fp}`;
  sub.font = { bold: true, size: 11, color: { argb: "FFFFFFFF" } };
  sub.alignment = { vertical: "middle", indent: 1 };
  sub.fill = fill(ACCENT);
  ws.getRow(2).height = 22;

  // KPI bands
  const band = (labelRow, items) => {
    const per = 12 / items.length;
    let col = 1;
    ws.getRow(labelRow).height = 16;
    ws.getRow(labelRow + 1).height = 24;
    for (const it of items) {
      const c1 = COL(col), c2 = COL(col + per - 1);
      ws.mergeCells(`${c1}${labelRow}:${c2}${labelRow}`);
      ws.mergeCells(`${c1}${labelRow + 1}:${c2}${labelRow + 1}`);
      const L = ws.getCell(`${c1}${labelRow}`);
      L.value = it.label; L.alignment = { horizontal: "center", vertical: "middle" };
      L.font = { bold: true, size: 9, color: { argb: "FF64748B" } };
      L.fill = fill(LIGHT); L.border = border;
      const V = ws.getCell(`${c1}${labelRow + 1}`);
      V.value = it.value; V.alignment = { horizontal: "center", vertical: "middle" };
      V.font = { bold: true, size: 13, color: { argb: INK } };
      V.fill = fill(it.accent ? "FFE0F2FE" : "FFFFFFFF"); V.border = border;
      if (it.fmt) V.numFmt = it.fmt;
      col += per;
    }
  };
  band(4, [
    { label: "TAXABLE VALUE", value: t.t12Txbl, fmt: MONEY }, { label: "IGST", value: t.t12Igst, fmt: MONEY },
    { label: "CGST", value: t.t12Cgst, fmt: MONEY }, { label: "SGST", value: t.t12Sgst, fmt: MONEY },
    { label: "CESS", value: t.t12Cess, fmt: MONEY }, { label: "TOTAL TAX", value: totalTax, fmt: MONEY },
  ]);
  band(7, [
    { label: "TOTAL INVOICE VALUE", value: totalVal, fmt: MONEY, accent: true },
    { label: "VALID E-INVOICES", value: res.counts.validInvoices },
    { label: "HSN / SAC LINES", value: res.counts.hsnRows },
    { label: "CREDIT NOTES", value: res.counts.cdnr },
  ]);

  const sectionTitle = (text) => {
    const r = ws.addRow([text]);
    ws.mergeCells(`A${r.number}:L${r.number}`);
    const c = r.getCell(1);
    c.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
    c.fill = fill(ACCENT); c.alignment = { vertical: "middle", indent: 1 };
    r.height = 22;
  };
  const headerRow = (cells) => {
    const r = ws.addRow(cells);
    r.eachCell((c) => { c.fill = fill(BRAND); c.font = { bold: true, size: 10, color: { argb: "FFFFFFFF" } }; c.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; c.border = border; });
    r.height = 24;
    return r;
  };

  if (res.bridge) {
    const b = res.bridge;
    ws.addRow([]);
    sectionTitle("TABLE 12 (= SALES SHEET)  →  E-INVOICE SUPPORT RECONCILIATION");
    const brRow = (label, amount, note, kind) => {
      const r = ws.addRow([]);
      ws.mergeCells(`B${r.number}:H${r.number}`);
      ws.mergeCells(`I${r.number}:K${r.number}`);
      r.getCell(2).value = label;
      r.getCell(9).value = note || "";
      r.getCell(12).value = amount; r.getCell(12).numFmt = MONEY;
      r.getCell(9).font = { size: 9, italic: true, color: { argb: "FF94A3B8" } };
      const strong = kind === "base" || kind === "result";
      r.getCell(2).font = { bold: strong, color: { argb: INK } };
      r.getCell(12).font = { bold: strong, color: { argb: INK } };
      if (strong) [2, 9, 12].forEach((c) => (r.getCell(c).fill = fill(LIGHT)));
      if (kind === "residual" && Math.abs(amount) >= 1) [2, 12].forEach((c) => (r.getCell(c).fill = fill("FFFEE2E2")));
      [2, 9, 12].forEach((c) => (r.getCell(c).border = border));
    };
    brRow("Table 12 = Client Sales Sheet — taxable (BASE)", b.table12Txbl, "books drive Table 12", "base");
    brRow("less: B2C supplies", -b.less.b2c, "in Table 12 (hsn_b2c); not in B2B e-invoices");
    brRow("less: Exports / SEZ", -b.less.export, "in Table 12; reported via 6A separately");
    brRow("less: Cancelled invoices", -b.less.cancelled, "still in sales sheet; no IRN");
    brRow("less: Registered B2B not e-invoiced", -b.less.notEinvoiced, "book-only — investigate omission");
    brRow("less: Book vs portal value differences", -b.less.bookVsPortalAdj, "per-invoice mismatches (C14)");
    brRow("less: Credit / Debit notes", -b.less.cdn, "support nets these; Table 12 does not");
    brRow("= Portal valid B2B e-invoices (SUPPORT)", b.portalSupportTxbl, "cross-check figure", "result");
    brRow("Unreconciled residual", b.residual, Math.abs(b.residual) < 1 ? "fully reconciled" : "INVESTIGATE", "residual");
  }

  ws.addRow([]);
  sectionTitle("HSN / SAC SUMMARY  —  Table 12");
  headerRow(["Sr", "HSN/SAC", "Description", "UQC", "Qty", "Rate %", "Taxable", "IGST", "CGST", "SGST", "Cess", "Total"]);
  res.table12.forEach((h, i) => {
    const r = ws.addRow([h.sr, h.hsn, h.desc, h.uqc, h.qty, h.rt, h.txval, h.iamt, h.camt, h.samt, h.csamt, h.total]);
    r.eachCell((c) => { c.border = border; if (i % 2) c.fill = fill(BAND); });
    r.getCell(3).alignment = { wrapText: true };
    r.getCell(5).numFmt = "#,##0";
    r.getCell(6).numFmt = "0.00";
    [7, 8, 9, 10, 11, 12].forEach((col) => (r.getCell(col).numFmt = MONEY));
  });
  const tr = ws.addRow(["", "TOTAL", "", "", "", "", t.t12Txbl, t.t12Igst, t.t12Cgst, t.t12Sgst, t.t12Cess, totalVal]);
  tr.eachCell((c) => { c.font = { bold: true, color: { argb: INK } }; c.fill = fill(LIGHT); c.border = border; });
  [7, 8, 9, 10, 11, 12].forEach((col) => (tr.getCell(col).numFmt = MONEY));

  ws.addRow([]);
  sectionTitle("DOCUMENTS ISSUED  —  Table 13");
  headerRow(["Code", "Nature of Document", "From", "To", "Total", "Cancelled", "Net"]);
  res.table13.forEach((d, i) => {
    const r = ws.addRow([d.code, d.nature, d.from, d.to, d.total, d.cancel, d.net]);
    r.eachCell((c) => { c.border = border; if (i % 2) c.fill = fill(BAND); });
  });

  ws.addRow([]);
  const fr = ws.addRow(["Prepared with Sungrow GST Processor 2026 · figures auto-derived from GST-portal e-invoice data · for client review."]);
  ws.mergeCells(`A${fr.number}:L${fr.number}`);
  fr.getCell(1).font = { italic: true, size: 9, color: { argb: "FF94A3B8" } };

  ws.views = [{ state: "frozen", ySplit: 2, showGridLines: false }];
  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function BridgeBlock({ b }) {
  const row = (label, v, note) => (
    <tr key={label} className={Math.abs(v) < 1 ? "g1-br-zero" : ""}>
      <td className="left">less: {label}</td>
      <td className="num">{v < 0 ? "+ " : "− "}{inr(Math.abs(v))}</td>
      <td className="left small">{note}</td>
    </tr>
  );
  return (
    <>
      <div className="g1-cli-sec">Table 12 → E-invoice Support Reconciliation <span className="badge">base = sales sheet</span></div>
      <div className="g1-cli-wrap">
        <table className="g1-table g1-br">
          <tbody>
            <tr className="row-total"><td className="left">Table 12 = Client Sales Sheet — taxable (BASE)</td><td className="num">{inr(b.table12Txbl)}</td><td className="left small">books drive Table 12</td></tr>
            {row("B2C supplies", b.less.b2c, "in Table 12 (hsn_b2c); not in B2B e-invoices")}
            {row("Exports / SEZ", b.less.export, "in Table 12; reported via 6A separately")}
            {row("Cancelled invoices", b.less.cancelled, "still in sales sheet; no IRN")}
            {row("Registered B2B not e-invoiced", b.less.notEinvoiced, "book-only — investigate omission")}
            {row("Book vs portal value differences", b.less.bookVsPortalAdj, "per-invoice mismatches (C14)")}
            {row("Credit / Debit notes", b.less.cdn, "support nets these; Table 12 does not")}
            <tr className="row-warn"><td className="left">= Portal valid B2B e-invoices (SUPPORT)</td><td className="num">{inr(b.portalSupportTxbl)}</td><td className="left small">cross-check figure</td></tr>
            <tr className={Math.abs(b.residual) < 1 ? "" : "row-bad"}><td className="left">Unreconciled residual</td><td className="num">{inr(b.residual)}</td><td className="left small">{Math.abs(b.residual) < 1 ? "✓ fully reconciled" : "⚠ investigate — rows not classified"}</td></tr>
          </tbody>
        </table>
      </div>
    </>
  );
}

function ClientSummary({ res, onDownload, busy }) {
  const t = res.totals;
  const totalTax = (t.t12Igst || 0) + (t.t12Cgst || 0) + (t.t12Sgst || 0) + (t.t12Cess || 0);
  const totalVal = (t.t12Txbl || 0) + totalTax;
  const cls = res.portalReady ? (res.allPass ? "ok" : "warn") : "bad";
  const statusTxt = res.portalReady ? (res.allPass ? "✓ Ready to file" : "⚠ Review deviations") : "✗ Blocking issues";
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
          <p>{res.meta.supplierName ? res.meta.supplierName + " · " : ""}GSTIN {res.meta.gstin || "—"} · {res.counts.validInvoices} invoices · {res.counts.hsnRows} HSN lines</p>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <span className={`g1-cli-status ${cls}`}>{statusTxt}</span>
          <button className="g1-cli-dl" onClick={onDownload} disabled={busy}>{busy ? "Building…" : "⬇ Download formatted Excel"}</button>
        </div>
      </div>

      <div className="g1-cli-kpis">
        {kpi("Taxable Value", inr0(t.t12Txbl))}
        {kpi("IGST", inr0(t.t12Igst))}
        {kpi("CGST", inr0(t.t12Cgst))}
        {kpi("SGST", inr0(t.t12Sgst))}
      </div>
      <div className="g1-cli-kpis">
        {kpi("Cess", inr0(t.t12Cess))}
        {kpi("Total Tax", inr0(totalTax))}
        {kpi("Total Invoice Value", inr0(totalVal), "taxable + tax", true)}
        {kpi("Credit Notes", res.counts.cdnr, res.counts.cancelled ? `${res.counts.cancelled} cancelled IRNs` : null)}
      </div>

      {res.bridge && <BridgeBlock b={res.bridge} />}

      <div className="g1-cli-sec">HSN / SAC Summary <span className="badge">Table 12 · {res.table12.length} lines</span></div>
      <div className="g1-cli-wrap">
        <table className="g1-table">
          <thead><tr><th>Sr</th><th>HSN/SAC</th><th className="left">Description</th><th>UQC</th><th className="num">Qty</th><th className="num">Rate%</th><th className="num">Taxable</th><th className="num">IGST</th><th className="num">CGST</th><th className="num">SGST</th><th className="num">Cess</th><th className="num">Total</th></tr></thead>
          <tbody>
            {res.table12.map((h) => (
              <tr key={h.sr + h.hsn + h.rt}><td>{h.sr}</td><td>{h.hsn}</td><td className="left small">{h.desc}</td><td>{h.uqc}</td><td className="num">{Number(h.qty).toLocaleString("en-IN")}</td><td className="num">{h.rt}</td><td className="num">{inr(h.txval)}</td><td className="num">{inr(h.iamt)}</td><td className="num">{inr(h.camt)}</td><td className="num">{inr(h.samt)}</td><td className="num">{inr(h.csamt)}</td><td className="num">{inr(h.total)}</td></tr>
            ))}
            <tr className="row-total"><td></td><td>TOTAL</td><td></td><td></td><td></td><td></td><td className="num">{inr(t.t12Txbl)}</td><td className="num">{inr(t.t12Igst)}</td><td className="num">{inr(t.t12Cgst)}</td><td className="num">{inr(t.t12Sgst)}</td><td className="num">{inr(t.t12Cess)}</td><td className="num">{inr(totalVal)}</td></tr>
          </tbody>
        </table>
      </div>

      <div className="g1-cli-sec">Documents Issued <span className="badge">Table 13</span></div>
      <div className="g1-cli-wrap">
        <table className="g1-table">
          <thead><tr><th>Code</th><th className="left">Nature of Document</th><th>From</th><th>To</th><th className="num">Total</th><th className="num">Cancelled</th><th className="num">Net</th></tr></thead>
          <tbody>{res.table13.map((d) => (<tr key={d.code}><td>{d.code}</td><td className="left">{d.nature}</td><td>{d.from}</td><td>{d.to}</td><td className="num">{d.total}</td><td className="num">{d.cancel}</td><td className="num">{d.net}</td></tr>))}</tbody>
        </table>
      </div>

      <p className="g1-cli-note">Bird's-eye view for the client — Table 12 (HSN/SAC) and Table 13 (documents issued), with the same figures that flow into the GSTR-1 JSON. Use <b>Download formatted Excel</b> for a presentation-grade workbook to share with the client.</p>
    </div>
  );
}

function download(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
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
          Builds Table 12 (HSN) &amp; Table 13 (Documents issued) <b>from the client Sales sheet — the books are the base</b>, and uses
          the GST-portal e-invoice dump as <b>support</b> to enrich UQC/description and cross-validate every deviation. Exports a
          portal-schema GSTR-1 JSON. <b>Table 12 covers all sales (B2B + B2C + exports);</b> invoice-level JSON for B2C/exports is
          handled later. All processing is local — nothing leaves your browser.
        </p>
        <ol className="g1-steps">
          <li>Upload the GST-portal <b>e-invoice dump</b> (EINV_&lt;gstin&gt;_&lt;fy&gt;.xlsx) — the authoritative source.</li>
          <li>Upload the client's <b>GSTR-1 working file</b> (the report under validation).</li>
          <li>Pick the return month. Run. Review deviations, resolve with client, then download the JSON.</li>
        </ol>
      </section>

      <section className="g1-card">
        <FilePick label="① E-invoice dump (.xlsx)" hint="EINV_<gstin>_<fy>.xlsx from the GST portal" accept=".xlsx" file={einvFile} onFile={setEinvFile} />
        <FilePick label="② Client GSTR-1 working (.xlsx)" hint="GSTR-1 Data <month> <yy>.xlsx shared by the client" accept=".xlsx" file={salesFile} onFile={setSalesFile} />
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
              {[["client", "🦅 Client Summary"], ["dev", "Deviations & Reconciliation"], ["t12", "Table 12 (HSN)"], ["t13", "Table 13 (Docs)"], ["checks", "Validation Checks"], ["json", "JSON Preview"]].map(([id, l]) => (
                <button key={id} className={tab === id ? "active" : ""} onClick={() => setTab(id)}>{l}</button>
              ))}
            </div>

            {tab === "client" && <ClientSummary res={res} onDownload={dlClientExcel} busy={xlBusy} />}
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
