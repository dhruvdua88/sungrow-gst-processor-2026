import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { parseDaybook, analyzeDaybook } from "./engine.js";
import "./daybook.css";

const inr = (n) =>
  n == null || n === "" ? "—" : "₹" + Number(n).toLocaleString("en-IN", { maximumFractionDigits: 0 });
const pctStr = (p) => (p == null ? "—" : Number(p).toFixed(1) + "%");

// ---------- generic sortable + filterable table ----------
function SortTable({ columns, rows, initialSort, initialDir = "desc", maxRows = 1000, placeholder = "Filter…", chips }) {
  const [sortKey, setSortKey] = useState(initialSort);
  const [dir, setDir] = useState(initialDir);
  const [q, setQ] = useState("");
  const valOf = (row, c) => (c.value ? c.value(row) : "");
  const filtered = useMemo(() => {
    const n = q.trim().toLowerCase();
    if (!n) return rows;
    return rows.filter((r) => columns.some((c) => String(valOf(r, c)).toLowerCase().includes(n)));
  }, [rows, q, columns]);
  const sorted = useMemo(() => {
    if (!sortKey) return filtered;
    const col = columns.find((c) => c.key === sortKey);
    if (!col) return filtered;
    const a = [...filtered];
    a.sort((x, y) => {
      const vx = valOf(x, col), vy = valOf(y, col);
      const cmp = col.numeric ? Number(vx) - Number(vy) : String(vx).localeCompare(String(vy));
      return dir === "asc" ? cmp : -cmp;
    });
    return a;
  }, [filtered, sortKey, dir, columns]);
  const shown = sorted.slice(0, maxRows);
  const click = (k) => { if (sortKey === k) setDir((d) => (d === "asc" ? "desc" : "asc")); else { setSortKey(k); setDir("desc"); } };
  return (
    <div>
      <div className="dbk-toolbar">
        <input className="dbk-search" placeholder={placeholder} value={q} onChange={(e) => setQ(e.target.value)} />
        {chips}
        <span className="dbk-count">{sorted.length.toLocaleString("en-IN")} rows{sorted.length > maxRows ? ` · top ${maxRows}` : ""}</span>
      </div>
      <div className="dbk-tablewrap">
        <table className="dbk-table">
          <thead><tr>{columns.map((c) => (
            <th key={c.key} className={c.numeric ? "num" : ""} onClick={() => click(c.key)}>
              {c.label}{sortKey === c.key ? (dir === "asc" ? " ▲" : " ▼") : ""}
            </th>
          ))}</tr></thead>
          <tbody>
            {shown.map((row, i) => (
              <tr key={i}>{columns.map((c) => (
                <td key={c.key} className={c.numeric ? "num" : ""}>{c.render ? c.render(row) : String(valOf(row, c))}</td>
              ))}</tr>
            ))}
            {shown.length === 0 && <tr><td className="dbk-empty" colSpan={columns.length}>No matching rows.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------- expandable group table ----------
function GroupTable({ groups, parentCols, childCols, getChildren, rowKey, initialSort, placeholder = "Filter…", maxRows = 400 }) {
  const [sortKey, setSortKey] = useState(initialSort);
  const [dir, setDir] = useState("desc");
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(() => new Set());
  const valOf = (row, c) => (c.value ? c.value(row) : "");
  const filtered = useMemo(() => {
    const n = q.trim().toLowerCase();
    if (!n) return groups;
    return groups.filter((g) => parentCols.some((c) => String(valOf(g, c)).toLowerCase().includes(n)));
  }, [groups, q, parentCols]);
  const sorted = useMemo(() => {
    if (!sortKey) return filtered;
    const col = parentCols.find((c) => c.key === sortKey);
    if (!col) return filtered;
    const a = [...filtered];
    a.sort((x, y) => { const vx = valOf(x, col), vy = valOf(y, col); const cmp = col.numeric ? Number(vx) - Number(vy) : String(vx).localeCompare(String(vy)); return dir === "asc" ? cmp : -cmp; });
    return a;
  }, [filtered, sortKey, dir, parentCols]);
  const shown = sorted.slice(0, maxRows);
  const click = (k) => { if (sortKey === k) setDir((d) => (d === "asc" ? "desc" : "asc")); else { setSortKey(k); setDir("desc"); } };
  const toggle = (k) => setOpen((s) => { const n = new Set(s); n.has(k) ? n.delete(k) : n.add(k); return n; });
  return (
    <div>
      <div className="dbk-toolbar">
        <input className="dbk-search" placeholder={placeholder} value={q} onChange={(e) => setQ(e.target.value)} />
        <span className="dbk-count">{sorted.length.toLocaleString("en-IN")} rows · click to expand</span>
      </div>
      <div className="dbk-tablewrap">
        <table className="dbk-table">
          <thead><tr><th className="dbk-expcol"></th>{parentCols.map((c) => (
            <th key={c.key} className={c.numeric ? "num" : ""} onClick={() => click(c.key)}>{c.label}{sortKey === c.key ? (dir === "asc" ? " ▲" : " ▼") : ""}</th>
          ))}</tr></thead>
          <tbody>
            {shown.map((g) => {
              const k = rowKey(g); const isOpen = open.has(k); const kids = isOpen ? getChildren(g) : [];
              return (
                <>
                  <tr key={k} className={"dbk-grouprow" + (isOpen ? " open" : "")} onClick={() => toggle(k)}>
                    <td className="dbk-expcol">{isOpen ? "▼" : "▶"}</td>
                    {parentCols.map((c) => (<td key={c.key} className={c.numeric ? "num" : ""}>{c.render ? c.render(g) : String(valOf(g, c))}</td>))}
                  </tr>
                  {isOpen && kids.map((ch, ci) => (
                    <tr key={k + "-" + ci} className="dbk-childrow">
                      <td className="dbk-expcol"></td>
                      {childCols.map((c, idx) => (<td key={c.key} className={(c.numeric ? "num" : "") + (idx === 0 ? " dbk-childfirst" : "")}>{c.render ? c.render(ch) : String(valOf(ch, c))}</td>))}
                    </tr>
                  ))}
                </>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

const Kpi = ({ label, value, tone }) => (
  <div className="dbk-kpi">
    <div className="dbk-kpi-label">{label}</div>
    <div className={"dbk-kpi-value" + (tone ? " " + tone : "")}>{value}</div>
  </div>
);
const Card = ({ title, sub, children }) => (
  <section className="dbk-card">
    <div className="dbk-card-head"><h3>{title}</h3>{sub ? <p>{sub}</p> : null}</div>
    <div className="dbk-card-body">{children}</div>
  </section>
);

// ---------- P&L ----------
function PnlView({ pnl }) {
  if (!pnl.available) return <Card title="Profit & Loss"><p className="dbk-muted">No P&L accounts (GL 6xxx/7xxx) found.</p></Card>;
  const rev = Math.max(pnl.revenue, 1);
  const Row = ({ label, amt, p, bold, neg, bar }) => (
    <div className={"dbk-pnl-row" + (bold ? " bold" : "")}>
      <span className="dbk-pnl-label">{label}</span>
      <span className="dbk-pnl-bar"><span style={{ width: Math.min(1, Math.abs(bar || 0)) * 100 + "%" }} /></span>
      <span className={"dbk-pnl-amt num" + (neg ? " neg" : "")}>{inr(amt)}</span>
      <span className="dbk-pnl-pct num">{p != null ? pctStr(p) : ""}</span>
    </div>
  );
  return (
    <Card title="Profit & Loss" sub="Built from the daybook — GL 6xxx (revenue / COGS / finance) and 7xxx (operating expense).">
      <div className="dbk-pnl-grid">
        <div>
          <Row label="Revenue" amt={pnl.revenue} p={100} bar={1} />
          <Row label="Less: Cost of sales" amt={-pnl.cogs} p={-(pnl.cogs / rev) * 100} bar={pnl.cogs / rev} neg />
          <Row label="Gross Profit" amt={pnl.grossProfit} p={pnl.grossMarginPct} bold neg={pnl.grossProfit < 0} bar={Math.abs(pnl.grossProfit) / rev} />
          <Row label="Less: Operating expenses" amt={-pnl.operatingExpenses} p={-(pnl.operatingExpenses / rev) * 100} bar={pnl.operatingExpenses / rev} neg />
          <Row label="Less: Finance (net)" amt={-pnl.finance} p={-(pnl.finance / rev) * 100} bar={Math.abs(pnl.finance) / rev} neg />
          <Row label={pnl.pbt < 0 ? "Loss before tax" : "Profit before tax"} amt={pnl.pbt} p={pnl.netMarginPct} bold neg={pnl.pbt < 0} bar={Math.abs(pnl.pbt) / rev} />
        </div>
        <div className="dbk-topexp">
          <h4>Top operating expense heads</h4>
          {pnl.opexLines.slice(0, 8).map((l) => {
            const max = Math.abs((pnl.opexLines[0] || {}).amount || 1);
            return (
              <div key={l.code} className="dbk-topexp-row">
                <span className="dbk-topexp-name" title={l.name}>{l.name}</span>
                <span className="dbk-topexp-bar"><span style={{ width: (Math.abs(l.amount) / max) * 100 + "%" }} /></span>
                <span className="num dbk-topexp-amt">{inr(l.amount)}</span>
              </div>
            );
          })}
        </div>
      </div>
    </Card>
  );
}

const bucketTag = (b) => <span className={"dbk-tag b" + String(b).replace("%", "").replace("other", "x")}>{b}</span>;

// ---------- Sales & GST ----------
function SalesView({ sales }) {
  if (!sales.available) return <Card title="Sales & GST"><p className="dbk-muted">No sale vouchers (revenue GL 6xxx) found.</p></Card>;
  const custParent = [
    { key: "customer", label: "Customer", value: (r) => r.customer },
    { key: "taxable", label: "Taxable Sales", numeric: true, value: (r) => r.taxable, render: (r) => inr(r.taxable) },
    { key: "gst", label: "Output GST", numeric: true, value: (r) => r.gst, render: (r) => inr(r.gst) },
    { key: "rate", label: "Blended %", numeric: true, value: (r) => r.rate || 0, render: (r) => (r.rate != null ? pctStr(r.rate) : "—") },
    { key: "vouchers", label: "Vch", numeric: true, value: (r) => r.vouchers },
  ];
  const custChild = [
    { key: "rate", label: "Rate", value: (r) => r.rate, render: (r) => bucketTag(r.rate) },
    { key: "taxable", label: "Taxable", numeric: true, value: (r) => r.taxable, render: (r) => inr(r.taxable) },
    { key: "gst", label: "GST", numeric: true, value: (r) => r.gst, render: (r) => inr(r.gst) },
    { key: "v", label: "Vch", numeric: true, value: (r) => r.vouchers },
    { key: "sp", label: "", value: () => "" },
  ];
  const vCols = [
    { key: "date", label: "Date", value: (r) => r.date },
    { key: "docType", label: "Type", value: (r) => r.docType },
    { key: "docNo", label: "Doc No", value: (r) => r.docNo },
    { key: "customer", label: "Customer", value: (r) => r.customer },
    { key: "taxable", label: "Taxable", numeric: true, value: (r) => r.taxable, render: (r) => inr(r.taxable) },
    { key: "sgst", label: "SGST", numeric: true, value: (r) => r.sgst, render: (r) => (r.sgst ? inr(r.sgst) : "—") },
    { key: "cgst", label: "CGST", numeric: true, value: (r) => r.cgst, render: (r) => (r.cgst ? inr(r.cgst) : "—") },
    { key: "igst", label: "IGST", numeric: true, value: (r) => r.igst, render: (r) => (r.igst ? inr(r.igst) : "—") },
    { key: "rate", label: "GST Rate", numeric: true, value: (r) => r.rate, render: (r) => bucketTag(r.rateBucket) },
    { key: "flag", label: "GST charged?", value: (r) => (r.hasGst ? "Yes" : "No GST"), render: (r) => (r.hasGst ? <span className="dbk-ok">✓ Yes</span> : <span className="dbk-bad">✗ No GST</span>) },
  ];
  const salesPct = (n) => (sales.totalTaxable ? ((n / sales.totalTaxable) * 100).toFixed(1) + "%" : "—");
  return (
    <Card title="Sales & GST" sub="Every sale voucher (revenue GL 6xxx) with output SGST/CGST/IGST, effective rate, and a GST-charged check.">
      <div className="dbk-kpis">
        <Kpi label="Taxable sales" value={inr(sales.totalTaxable)} />
        <Kpi label="Output GST" value={inr(sales.totalGst)} tone="brand" />
        <Kpi label="Blended rate" value={sales.blendedRate != null ? sales.blendedRate.toFixed(2) + "%" : "—"} tone="ok" />
        <Kpi label="Sale vouchers" value={String(sales.vouchers.length)} />
        <Kpi label="Sales without GST" value={sales.noGstCount + " · " + inr(sales.noGstValue)} tone={sales.noGstCount ? "bad" : "ok"} />
      </div>
      <h4 className="dbk-subhead">GST rate distribution</h4>
      <div className="dbk-ratecards">
        {sales.rateBuckets.map((b) => (
          <div key={b.rate} className="dbk-ratecard">
            <div className="dbk-ratecard-top"><span className="dbk-ratecard-rate">{b.rate}</span><span className="num">{b.vouchers} vch</span></div>
            <div className="dbk-ratecard-amt num">{inr(b.taxable)}</div>
            <div className="dbk-ratecard-sub num">GST {inr(b.gst)} · {salesPct(b.taxable)}</div>
          </div>
        ))}
      </div>
      {sales.noGstCount > 0 && (
        <div className="dbk-warn">⚠ <strong>{sales.noGstCount}</strong> sale voucher(s) totalling <strong>{inr(sales.noGstValue)}</strong> carry <strong>no output GST</strong>. Verify zero-rated (export/SEZ), stock-transfer, or a missed charge (filter “No GST” below).</div>
      )}
      <h4 className="dbk-subhead">Sales by customer — click to see the rate split</h4>
      <GroupTable groups={sales.customers} parentCols={custParent} childCols={custChild} getChildren={(c) => c.byRate} rowKey={(c) => c.customer} initialSort="taxable" placeholder="Filter customer…" />
      <h4 className="dbk-subhead">Sales voucher explorer</h4>
      <SortTable columns={vCols} rows={sales.vouchers} initialSort="taxable" placeholder="Filter customer / doc / rate / No GST…" />
    </Card>
  );
}

const pctCell = (p) => p == null ? <span className="dbk-muted">—</span> : <span className={"num " + (p >= 0.1 ? "dbk-ok" : p > 0 ? "dbk-warnp" : "dbk-muted")}>{p.toFixed(1)}%</span>;
const eff = (tds, amt) => (amt ? (tds / amt) * 100 : null);

// ---------- Expense head -> party ----------
function ExpenseView({ rows }) {
  const parent = [
    { key: "ledger", label: "Expense Head (Ledger)", value: (r) => r.ledger },
    { key: "glCode", label: "G/L", value: (r) => r.glCode },
    { key: "section", label: "Sec", value: (r) => r.section || "—" },
    { key: "expense", label: "Total Amount", numeric: true, value: (r) => r.expense, render: (r) => inr(r.expense) },
    { key: "tds", label: "TDS", numeric: true, value: (r) => r.tds, render: (r) => inr(r.tds) },
    { key: "eff", label: "TDS %", numeric: true, value: (r) => r.effRate ?? -1, render: (r) => pctCell(r.effRate) },
    { key: "rcm", label: "RCM", numeric: true, value: (r) => r.rcm, render: (r) => (r.rcm ? inr(r.rcm) : "—") },
    { key: "vendors", label: "Parties", numeric: true, value: (r) => r.vendors },
  ];
  const child = [
    { key: "party", label: "Party", value: (r) => r.party },
    { key: "g", label: "", value: () => "" }, { key: "s", label: "", value: () => "" },
    { key: "amount", label: "Amount", numeric: true, value: (r) => r.amount, render: (r) => inr(r.amount) },
    { key: "tds", label: "TDS", numeric: true, value: (r) => r.tds, render: (r) => inr(r.tds) },
    { key: "eff", label: "TDS %", numeric: true, value: (r) => eff(r.tds, r.amount) ?? -1, render: (r) => pctCell(eff(r.tds, r.amount)) },
    { key: "rcm", label: "RCM", numeric: true, value: (r) => r.rcm, render: (r) => (r.rcm ? inr(r.rcm) : "—") },
    { key: "docs", label: "Vch", numeric: true, value: (r) => r.docs },
  ];
  const tExp = rows.reduce((s, r) => s + r.expense, 0), tTds = rows.reduce((s, r) => s + r.tds, 0), tRcm = rows.reduce((s, r) => s + r.rcm, 0);
  return (
    <Card title="Expense Head → Party" sub="Each expense ledger with TDS, effective rate and RCM. Click a head for the party-wise split.">
      <div className="dbk-kpis">
        <Kpi label="Total expense (AP)" value={inr(tExp)} />
        <Kpi label="TDS deducted" value={inr(tTds)} tone="brand" />
        <Kpi label="Effective TDS %" value={tExp ? ((tTds / tExp) * 100).toFixed(2) + "%" : "—"} tone="ok" />
        <Kpi label="RCM (REVE CH)" value={inr(tRcm)} tone="manual" />
      </div>
      <GroupTable groups={rows} parentCols={parent} childCols={child} getChildren={(r) => r.parties} rowKey={(r) => r.glCode + r.ledger} initialSort="expense" placeholder="Filter expense head / GL / section…" />
    </Card>
  );
}

// ---------- Party / AP ----------
function PartyView({ rows }) {
  const parent = [
    { key: "party", label: "Party (AP / Accounting pro.)", value: (r) => r.party },
    { key: "expense", label: "Total Amount", numeric: true, value: (r) => r.expense, render: (r) => inr(r.expense) },
    { key: "tds", label: "TDS", numeric: true, value: (r) => r.tds, render: (r) => inr(r.tds) },
    { key: "eff", label: "TDS %", numeric: true, value: (r) => eff(r.tds, r.expense) ?? -1, render: (r) => pctCell(eff(r.tds, r.expense)) },
    { key: "rcm", label: "RCM", numeric: true, value: (r) => r.rcm, render: (r) => (r.rcm ? inr(r.rcm) : "—") },
    { key: "apCredit", label: "AP Invoiced", numeric: true, value: (r) => r.apCredit, render: (r) => inr(r.apCredit) },
    { key: "apDebit", label: "AP Paid", numeric: true, value: (r) => r.apDebit, render: (r) => inr(r.apDebit) },
    { key: "out", label: "Outstanding", numeric: true, value: (r) => r.apCredit - r.apDebit, render: (r) => inr(r.apCredit - r.apDebit) },
  ];
  const child = [
    { key: "ledger", label: "Expense head", value: (r) => r.ledger },
    { key: "amount", label: "Amount", numeric: true, value: (r) => r.amount, render: (r) => inr(r.amount) },
    { key: "tds", label: "TDS", numeric: true, value: (r) => r.tds, render: (r) => inr(r.tds) },
    { key: "rcm", label: "RCM", numeric: true, value: (r) => r.rcm, render: (r) => (r.rcm ? inr(r.rcm) : "—") },
    { key: "docs", label: "Vch", numeric: true, value: (r) => r.docs },
    { key: "s1", label: "", value: () => "" }, { key: "s2", label: "", value: () => "" }, { key: "s3", label: "", value: () => "" },
  ];
  return (
    <Card title="Party / AP — Bird's-eye" sub="Every AP party: spend, TDS, RCM, AP invoiced vs paid (outstanding). Click a party to see where the money goes.">
      <GroupTable groups={rows} parentCols={parent} childCols={child} getChildren={(r) => r.heads} rowKey={(r) => r.party} initialSort="expense" placeholder="Filter party…" />
    </Card>
  );
}

const LINE_TYPES = ["Expense", "TDS", "GST", "Vendor/AP", "Bank", "Inventory/FA", "Rev/COS", "Receivable", "Other"];
function TxnView({ txns }) {
  const [type, setType] = useState("All");
  const counts = useMemo(() => { const c = {}; for (const t of txns) c[t.lineType] = (c[t.lineType] || 0) + 1; return c; }, [txns]);
  const rows = useMemo(() => (type === "All" ? txns : txns.filter((t) => t.lineType === type)), [txns, type]);
  const cols = [
    { key: "date", label: "Date", value: (r) => r.dateValue ?? r.date, render: (r) => r.date },
    { key: "voucher_type", label: "Doc Type", value: (r) => r.voucher_type },
    { key: "voucher_no", label: "Doc No", value: (r) => r.voucher_no },
    { key: "gl_code", label: "G/L Account", value: (r) => r.gl_code },
    { key: "ledger", label: "G/L Account Text (Ledger)", value: (r) => r.ledger },
    { key: "vendor", label: "Accounting pro. (Party)", value: (r) => r.vendor },
    { key: "reference", label: "Reference", value: (r) => r.reference },
    { key: "narration", label: "Text", value: (r) => r.narration },
    { key: "lineType", label: "Type", value: (r) => r.lineType, render: (r) => <span className="dbk-ltag">{r.lineType}{r.tdsLedgerSection ? " " + r.tdsLedgerSection : ""}</span> },
    { key: "debit", label: "Debit", numeric: true, value: (r) => r.debit, render: (r) => (r.debit ? inr(r.debit) : "") },
    { key: "credit", label: "Credit", numeric: true, value: (r) => r.credit, render: (r) => (r.credit ? inr(r.credit) : "") },
  ];
  const chip = (label, key, n) => (
    <button key={key} className={"dbk-chip" + (type === key ? " active" : "")} onClick={() => setType(key)}>{label} <span className="num">{n.toLocaleString("en-IN")}</span></button>
  );
  const chips = (<>{chip("All", "All", txns.length)}{LINE_TYPES.filter((t) => counts[t]).map((t) => chip(t, t, counts[t]))}</>);
  return (
    <Card title="Transactions Explorer" sub="Every line — GL #, Ledger, Party, Reference side by side. Sort any column, filter by text or line type.">
      <SortTable columns={cols} rows={rows} initialSort="debit" placeholder="Filter GL / ledger / party / reference / text…" chips={<div className="dbk-chips">{chips}</div>} maxRows={1000} />
    </Card>
  );
}

// ---------- Excel export ----------
function downloadExcel(month, res, txns) {
  const wb = XLSX.utils.book_new();
  const add = (name, rows) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{ Note: "none" }]), name.slice(0, 31));
  const p = res.mis.pnl, s = res.sales, m = res.mis;
  add("Summary", [
    { Metric: "Month", Value: month }, { Metric: "Transactions", Value: txns.length },
    { Metric: "Revenue", Value: Math.round(p.revenue) }, { Metric: "Gross Profit", Value: Math.round(p.grossProfit) },
    { Metric: "Profit/Loss before tax", Value: Math.round(p.pbt) },
    { Metric: "Taxable sales", Value: Math.round(s.totalTaxable) }, { Metric: "Output GST", Value: Math.round(s.totalGst) },
    { Metric: "Sales without GST (count/value)", Value: s.noGstCount + " / " + Math.round(s.noGstValue) },
    { Metric: "Expense thru AP", Value: Math.round(m.totalExpense) }, { Metric: "TDS deducted", Value: Math.round(m.totalTds) },
    { Metric: "RCM (REVE CH)", Value: Math.round(m.totalRcm) },
  ]);
  if (p.available) add("P&L", [
    { Line: "Revenue", Amount: Math.round(p.revenue) }, { Line: "Cost of sales", Amount: Math.round(p.cogs) },
    { Line: "Gross Profit", Amount: Math.round(p.grossProfit) }, { Line: "Operating expenses", Amount: Math.round(p.operatingExpenses) },
    { Line: "Finance (net)", Amount: Math.round(p.finance) }, { Line: p.pbt < 0 ? "Loss before tax" : "Profit before tax", Amount: Math.round(p.pbt) },
    ...p.opexLines.map((l) => ({ Line: "  " + l.name + " (" + l.code + ")", Amount: Math.round(l.amount) })),
  ]);
  if (s.available) {
    add("GST Rate Summary", s.rateBuckets.map((b) => ({ "GST Rate": b.rate, Vouchers: b.vouchers, "Taxable Sales": Math.round(b.taxable), "Output GST": Math.round(b.gst) })));
    add("Sales Vouchers", s.vouchers.map((v) => ({ Date: v.date, Type: v.docType, "Doc No": v.docNo, Customer: v.customer, Taxable: Math.round(v.taxable), SGST: Math.round(v.sgst), CGST: Math.round(v.cgst), IGST: Math.round(v.igst), "GST Rate": v.rateBucket, "GST charged?": v.hasGst ? "Yes" : "NO GST" })));
  }
  const ehp = [];
  for (const r of m.expenseTds) {
    ehp.push({ "Expense Head": r.ledger, "G/L": r.glCode, Party: "— TOTAL —", Amount: Math.round(r.expense), TDS: Math.round(r.tds), RCM: Math.round(r.rcm), Vch: r.docs });
    for (const pr of r.parties) ehp.push({ "Expense Head": r.ledger, "G/L": r.glCode, Party: pr.party, Amount: Math.round(pr.amount), TDS: Math.round(pr.tds), RCM: Math.round(pr.rcm), Vch: pr.docs });
  }
  add("Expense x Party", ehp);
  const pah = [];
  for (const v of m.vendors) {
    pah.push({ Party: v.party, "Expense Head": "— TOTAL —", Amount: Math.round(v.expense), TDS: Math.round(v.tds), RCM: Math.round(v.rcm), "AP Invoiced": Math.round(v.apCredit), "AP Paid": Math.round(v.apDebit) });
    for (const h of v.heads) pah.push({ Party: v.party, "Expense Head": h.ledger, Amount: Math.round(h.amount), TDS: Math.round(h.tds), RCM: Math.round(h.rcm), "AP Invoiced": "", "AP Paid": "" });
  }
  add("Party AP x Head", pah);
  add("Transactions", txns.map((t) => ({ Date: t.date, "Doc Type": t.voucher_type, "Doc No": t.voucher_no, "G/L Account": t.gl_code, Ledger: t.ledger, Party: t.vendor, Reference: t.reference, Text: t.narration, "Line Type": t.lineType, "TDS Sec": t.tdsLedgerSection, Debit: t.debit, Credit: t.credit })));
  const outArr = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  const blob = new Blob([outArr], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "Daybook_MIS_" + (month || "report") + ".xlsx";
  a.click();
  URL.revokeObjectURL(a.href);
}

// ---------- main ----------
export default function DaybookAnalysis() {
  const [month, setMonth] = useState("Jun-26");
  const [fileName, setFileName] = useState(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState(null);
  const [data, setData] = useState(null); // { txns, res }
  const inputRef = useRef(null);

  const onFile = async (file) => {
    if (!file) return;
    setBusy(true); setError(null);
    try {
      const parsed = await parseDaybook(file);
      if (!parsed.txns.length) throw new Error("No transactions parsed — file may be empty or an unexpected format.");
      const res = analyzeDaybook(parsed.txns);
      setData({ txns: parsed.txns, res });
      setFileName(file.name);
    } catch (e) {
      setError(e.message || String(e));
      setData(null);
    } finally { setBusy(false); }
  };

  return (
    <div className="dbk-root">
      <section className="dbk-card dbk-upload-card">
        <div className="dbk-upload-head">
          <div>
            <h3>Daybook Analysis</h3>
            <p>Upload one SAP/Tally daybook (.xlsx). Builds the P&amp;L, Sales &amp; GST, Expense → Party and Party / AP views — all in your browser.</p>
          </div>
          <div className="dbk-month">
            <span>Month</span>
            <input value={month} onChange={(e) => setMonth(e.target.value)} />
          </div>
        </div>
        <div className={"dbk-drop" + (fileName ? " done" : "")} onClick={() => inputRef.current && inputRef.current.click()}
          onDragOver={(e) => e.preventDefault()} onDrop={(e) => { e.preventDefault(); onFile(e.dataTransfer.files[0]); }}>
          <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={(e) => onFile(e.target.files[0])} />
          {fileName ? <span className="dbk-ok">✓ {fileName}</span> : <span>Click or drop the daybook .xlsx</span>}
        </div>
        {busy && <div className="dbk-busy">Parsing & analysing…</div>}
        {error && <div className="dbk-warn">⚠ {error}</div>}
        {data && (
          <div className="dbk-kpis dbk-kpis-top">
            <Kpi label="Revenue" value={inr(data.res.mis.pnl.revenue)} tone="brand" />
            <Kpi label="Gross margin" value={pctStr(data.res.mis.pnl.grossMarginPct)} />
            <Kpi label={data.res.mis.pnl.pbt < 0 ? "Loss before tax" : "Profit before tax"} value={inr(data.res.mis.pnl.pbt)} tone={data.res.mis.pnl.pbt < 0 ? "bad" : "ok"} />
            <Kpi label="Output GST" value={inr(data.res.sales.totalGst)} />
            <Kpi label="TDS deducted" value={inr(data.res.mis.totalTds)} />
            <Kpi label="RCM" value={inr(data.res.mis.totalRcm)} tone="manual" />
          </div>
        )}
        {data && <button className="dbk-dl" onClick={() => downloadExcel(month, data.res, data.txns)}>⬇ Download Excel workbook</button>}
      </section>

      {data && (
        <>
          <PnlView pnl={data.res.mis.pnl} />
          <SalesView sales={data.res.sales} />
          <ExpenseView rows={data.res.mis.expenseTds} />
          <PartyView rows={data.res.mis.vendors} />
          <TxnView txns={data.txns} />
        </>
      )}
    </div>
  );
}
