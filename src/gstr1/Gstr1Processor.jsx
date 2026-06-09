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
  const [running, setRunning] = useState(false);
  const [res, setRes] = useState(null);
  const [err, setErr] = useState("");
  const [tab, setTab] = useState("dev");

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
      const out = processGstr1(einvWb, salesWb, { gstin, fp, periodLabel: `${MONTHS[month]} ${year}` });
      if (out.errors.length) setErr(out.errors.join(" "));
      setRes(out);
    } catch (e) {
      console.error(e); setErr(e.message || String(e));
    } finally { setRunning(false); }
  };

  const dlJson = () => download(new Blob([JSON.stringify(res.json, null, 2)], { type: "application/json" }), `GSTR1_${res.meta.gstin || "return"}_${res.meta.fp}.json`);
  const dlExcel = async () => download(await buildExcel(res), `GSTR1_validation_${res.meta.gstin || "SIPL"}_${res.meta.fp}.xlsx`);

  return (
    <div className="g1">
      <section className="g1-intro">
        <h2>GSTR-1 Processor</h2>
        <p>
          Validates the client's monthly GSTR-1 working against the GST-portal e-invoice auto-population, surfaces every deviation
          for client confirmation, builds Table 12 (HSN) &amp; Table 13 (Documents issued), and exports a portal-schema GSTR-1 JSON.
          <b> Scope: B2B + Credit/Debit Notes. </b> Exports / B2C handled later. All processing is local — nothing leaves your browser.
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
        </div>
        <button className="g1-run" onClick={run} disabled={running}>{running ? "Processing…" : "Process GSTR-1"}</button>
        {err && <p className="g1-err">{err}</p>}
      </section>

      {res && (
        <>
          <Banner res={res} />
          <Cards res={res} />
          <section className="g1-card">
            <div className="g1-actions">
              <button className="g1-dl primary" onClick={dlJson} disabled={!res.portalReady} title={res.portalReady ? "" : "Resolve blocking failures first"}>⬇ Download GSTR-1 JSON</button>
              <button className="g1-dl" onClick={dlExcel}>⬇ Download validation workbook (.xlsx)</button>
            </div>
            <div className="g1-tabs">
              {[["dev", "Deviations & Reconciliation"], ["t12", "Table 12 (HSN)"], ["t13", "Table 13 (Docs)"], ["checks", "Validation Checks"], ["json", "JSON Preview"]].map(([id, l]) => (
                <button key={id} className={tab === id ? "active" : ""} onClick={() => setTab(id)}>{l}</button>
              ))}
            </div>

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
