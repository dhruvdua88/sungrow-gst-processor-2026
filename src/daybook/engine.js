// ---------------------------------------------------------------------------
// Daybook Analysis engine (browser, plain JS). SAP/Tally daybook -> P&L,
// Sales & GST, Expense -> Party, Party / AP. Self-contained; does not touch
// the existing GST-processor engine.
//
// Incorporates the practices built for the SDIPL daybook MIS:
//  - keyword column detection (GL code / GL text=ledger / Accounting pro.=party
//    / reference / DR / CR), no fixed column names.
//  - line classification by GL prefix + ledger text.
//  - P&L from GL 6xxx (revenue/COGS/finance) + 7xxx (operating expense).
//  - Expense -> AP -> TDS/RCM, pro-rata across a voucher's expense lines.
//  - RCM = full REVE CH credit (GL 2221013010/30/40).
//  - Sales output GST per voucher + effective rate + "GST charged?" check.
// ---------------------------------------------------------------------------
import * as XLSX from "xlsx";

// ---- parsing ----
export async function parseDaybook(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: null, raw: true });
  if (!matrix.length) return { headers: [], rows: [], txns: [] };
  // header row = best text-filled row in first 25
  let best = 0, bs = -1;
  for (let i = 0; i < Math.min(matrix.length, 25); i++) {
    const r = matrix[i] || [];
    const tc = r.filter((c) => typeof c === "string" && String(c).trim()).length;
    const f = r.filter((c) => c !== null && c !== undefined && String(c).trim() !== "").length;
    const sc = tc * 2 + f;
    if (f >= 2 && sc > bs) { bs = sc; best = i; }
  }
  const width = matrix.reduce((m, r) => Math.max(m, r ? r.length : 0), 0);
  const headers = [];
  const hr = matrix[best] || [];
  for (let c = 0; c < width; c++) {
    let l = hr[c] == null ? "" : String(hr[c]).trim().replace(/\s+/g, " ");
    if (!l) l = "Column " + (c + 1);
    headers.push(l);
  }
  const rows = [];
  for (let i = best + 1; i < matrix.length; i++) {
    const r = matrix[i] || [];
    if (r.filter((c) => c !== null && c !== undefined && String(c).trim() !== "").length === 0) continue;
    const a = [];
    for (let c = 0; c < width; c++) a.push(r[c] == null ? null : r[c]);
    rows.push(a);
  }
  const map = detectColumns(headers);
  const txns = normalize(rows, map);
  return { headers, rows, map, txns, sheetName: wb.SheetNames[0] };
}

// ---- column detection ----
const FIELD_KEYWORDS = {
  date: ["voucher date", "posting date", "document date", "trans date", "date"],
  voucher_no: ["documentno", "document no", "voucher no", "voucher number", "vch no", "doc no", "voucher"],
  voucher_type: ["document type", "voucher type", "vch type", "transaction type", "type"],
  ledger: ["g/l account text", "gl account text", "account text", "account name", "ledger", "particulars", "head", "nominal"],
  gl_code: ["g/l account", "gl account", "account code", "account no", "gl code", "gl"],
  vendor: ["accounting pro", "vendor", "supplier", "party", "customer", "payee", "name"],
  narration: ["narration", "description", "remarks", "remark", "text", "note", "particular"],
  reference: ["reference", "ref.key", "ref key", "ref no", "ref"],
  invoice_no: ["invoice no", "invoice number", "bill no", "invoice", "bill"],
  debit: ["debit", "dr amount", "dr"],
  credit: ["credit", "cr amount", "cr"],
  amount: ["net amount", "amount", "value", "gross amount"],
};
const norm = (s) => String(s ?? "").toLowerCase().replace(/[^a-z0-9 /]/g, " ").replace(/\s+/g, " ").trim();
function scoreHeader(header, kws) {
  const h = norm(header);
  if (!h) return 0;
  for (let i = 0; i < kws.length; i++) {
    if (h === kws[i]) return 100 - i;
    if (h.includes(kws[i])) return 60 - i;
  }
  return 0;
}
export function detectColumns(headers) {
  const map = {};
  const used = new Set();
  const cands = [];
  for (const field of Object.keys(FIELD_KEYWORDS)) {
    headers.forEach((h, col) => {
      const sc = scoreHeader(h, FIELD_KEYWORDS[field]);
      if (sc > 0) cands.push({ field, col, score: sc });
    });
  }
  cands.sort((a, b) => b.score - a.score);
  for (const c of cands) {
    if (map[c.field] !== undefined || used.has(c.col)) continue;
    map[c.field] = c.col;
    used.add(c.col);
  }
  return map;
}

// ---- normalization ----
const PAN_RE = /[A-Z]{5}[0-9]{4}[A-Z]/;
const GSTIN_RE = /[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][0-9A-Z]{3}/;
export function toNumber(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return isFinite(v) ? v : 0;
  let s = String(v).trim();
  if (!s) return 0;
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  s = s.replace(/[₹$,\s]/g, "").replace(/[^0-9.\-]/g, "");
  if (s === "" || s === "-" || s === ".") return 0;
  const n = parseFloat(s);
  if (!isFinite(n)) return 0;
  return neg ? -n : n;
}
const pad = (n) => (n < 10 ? "0" + n : "" + n);
function fmtDate(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
function excelSerialToDate(serial) {
  if (!isFinite(serial) || serial <= 0 || serial > 80000) return null;
  const ms = Math.round((serial - 25569) * 86400 * 1000);
  const d = new Date(ms);
  return isNaN(d.getTime()) ? null : d;
}
function parseDate(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return { display: fmtDate(v), value: v.getTime() };
  if (typeof v === "number") {
    const d = excelSerialToDate(v);
    if (d) return { display: fmtDate(d), value: d.getTime() };
    return { display: String(v), value: null };
  }
  const s = String(v ?? "").trim();
  if (!s) return { display: "", value: null };
  const p = Date.parse(s);
  return isNaN(p) ? { display: s, value: null } : { display: s, value: p };
}
const toStr = (v) => (v == null ? "" : v instanceof Date ? fmtDate(v) : String(v).trim());
const cell = (row, idx) => (idx === undefined ? null : row[idx] ?? null);

export function parseTdsSection(ledger) {
  const u = ledger.toUpperCase();
  const m = u.match(/\b(19[0-9][A-Z]?|192B|192)\b/);
  if (!m) return "";
  if (!/TDS|T\.?D\.?S|WITHHOLD|TAX DED/.test(u)) return "";
  return m[1];
}
export function classifyLine(glCode, ledger, tdsSection) {
  const g = String(glCode).replace(/\s/g, "");
  const l = ledger.toUpperCase();
  if (tdsSection || /\bTDS\b|TCS|WITHHOLD/.test(l)) return "TDS";
  if (/GST|IGST|CGST|SGST|UTGST|ITC/.test(l)) return "GST";
  if (g.startsWith("2202") || g.startsWith("2241")) return "Vendor/AP";
  if (g.startsWith("1002")) return "Bank";
  if (g.startsWith("7")) return "Expense";
  if (g.startsWith("6")) return "Rev/COS";
  if (/^1(4|5|6|7)/.test(g) || g.startsWith("5")) return "Inventory/FA";
  if (g.startsWith("112") || /RECEIVABLE|TRADE REC/.test(l)) return "Receivable";
  if (/PAYABLE|VENDOR|CREDITOR/.test(l)) return "Vendor/AP";
  if (/BANK|CASH/.test(l)) return "Bank";
  if (/EXPENSE|RENT|SALAR|FREIGHT|REPAIR|PROFE|COMMISSION/.test(l)) return "Expense";
  return "Other";
}
export function normalize(rows, map) {
  const out = [];
  rows.forEach((row, i) => {
    const di = parseDate(cell(row, map.date));
    const debit = toNumber(cell(row, map.debit));
    const credit = toNumber(cell(row, map.credit));
    let amount;
    if (map.amount !== undefined) {
      amount = toNumber(cell(row, map.amount));
      if (amount === 0 && (debit !== 0 || credit !== 0)) amount = debit - credit;
    } else amount = debit - credit;
    const ledger = toStr(cell(row, map.ledger));
    const vendor = toStr(cell(row, map.vendor));
    const narration = toStr(cell(row, map.narration));
    const voucher_type = toStr(cell(row, map.voucher_type));
    const voucher_no = toStr(cell(row, map.voucher_no));
    const gl_code = toStr(cell(row, map.gl_code));
    const reference = toStr(cell(row, map.reference));
    if (!ledger && !vendor && amount === 0 && debit === 0 && credit === 0) return;
    const rowText = row.map(toStr).join(" ").toUpperCase();
    const panM = rowText.match(PAN_RE);
    const gstM = rowText.match(GSTIN_RE);
    const tdsLedgerSection = parseTdsSection(ledger);
    out.push({
      rowIndex: i,
      date: di.display,
      dateValue: di.value,
      voucher_no, voucher_type, gl_code, ledger, vendor, narration, reference,
      debit, credit, amount, absolute_amount: Math.abs(amount),
      pan: panM ? panM[0] : "",
      gstin: gstM ? gstM[0] : "",
      combined_text: [ledger, vendor, narration, voucher_type].filter(Boolean).join(" | "),
      lineType: classifyLine(gl_code, ledger, tdsLedgerSection),
      tdsLedgerSection,
    });
  });
  return out;
}

// ---- MIS (P&L + expense->party + party->heads) ----
const isExpenseGl = (g) => g.length > 0 && (g[0] === "6" || g[0] === "7");
const isApExpenseGl = (g) => g.startsWith("7") || (g[0] === "6" && !g.startsWith("64") && !g.startsWith("66"));
const isRevenueGl = (g) => g.startsWith("6001") || g.startsWith("6051");
const isCogsGl = (g) => g.startsWith("64");
const isFinanceGl = (g) => g.startsWith("66");
const isApGl = (g) => g.startsWith("2202") || g.startsWith("2241");
const isApLine = (t) => isApGl(t.gl_code) || (!t.gl_code && t.lineType === "Vendor/AP");
const RCM_GL = new Set(["2221013010", "2221013030", "2221013040"]);
const isRcm = (t) => RCM_GL.has(t.gl_code) || /REVE\s*CH|REVERSE\s*CHARGE/i.test(t.ledger);
const isTds = (t) => !!t.tdsLedgerSection || (t.lineType === "TDS" && !isRcm(t));
const pct = (n, d) => (d ? (n / d) * 100 : null);

function inferSection(name, eff) {
  const n = name.toUpperCase();
  if (/RENT|LEASE/.test(n)) return "194I";
  if (/PROFE|LEGAL|CONSULT|AUDIT|TECHNICAL/.test(n)) return "194J";
  if (/COMMISSION|BROKERAGE/.test(n)) return "194H";
  if (/SALAR|WAGE|PAYROLL|BASIC|OVERTIME|MEAL/.test(n)) return "192B";
  if (/TRANSPORT|FREIGHT|REPAIR|HOUSEKEEP|SECURITY|CONTRACT|CARTAGE/.test(n)) return "194C";
  if (eff != null) { if (eff >= 9) return "194I/J"; if (eff >= 4) return "194C/H"; if (eff >= 1) return "194C"; }
  return "";
}

export function runMis(txns) {
  // P&L
  const glAgg = new Map();
  for (const t of txns) {
    const g = t.gl_code;
    if (!g || !isExpenseGl(g)) continue;
    let a = glAgg.get(g);
    if (!a) { a = { name: t.ledger, dr: 0, cr: 0 }; glAgg.set(g, a); }
    a.dr += t.debit; a.cr += t.credit;
    if (!a.name && t.ledger) a.name = t.ledger;
  }
  const revLines = [], cogsLines = [], opexLines = [], finLines = [];
  let revenue = 0, cogs = 0, operatingExpenses = 0, finance = 0;
  for (const [code, a] of glAgg) {
    if (isRevenueGl(code)) { const v = a.cr - a.dr; revenue += v; revLines.push({ code, name: a.name, amount: v }); }
    else if (isCogsGl(code)) { const v = a.dr - a.cr; cogs += v; cogsLines.push({ code, name: a.name, amount: v }); }
    else if (isFinanceGl(code)) { const v = a.dr - a.cr; finance += v; finLines.push({ code, name: a.name, amount: v }); }
    else { const v = a.dr - a.cr; operatingExpenses += v; opexLines.push({ code, name: a.name, amount: v }); }
  }
  const grossProfit = revenue - cogs;
  const pbt = grossProfit - operatingExpenses - finance;
  const sd = (a, b) => Math.abs(b.amount) - Math.abs(a.amount);
  [revLines, cogsLines, opexLines, finLines].forEach((x) => x.sort(sd));
  const pnl = {
    available: glAgg.size > 0, revenue, cogs, grossProfit, grossMarginPct: pct(grossProfit, revenue),
    operatingExpenses, finance, pbt, netMarginPct: pct(pbt, revenue),
    revenueLines: revLines, cogsLines, opexLines, financeLines: finLines,
  };

  // Expense -> AP -> TDS/RCM
  const byDoc = new Map();
  for (const t of txns) {
    const k = t.voucher_no || "row-" + t.rowIndex;
    const arr = byDoc.get(k); if (arr) arr.push(t); else byDoc.set(k, [t]);
  }
  const cellAcc = () => ({ amount: 0, tds: 0, rcm: 0, docs: new Set() });
  const exMap = new Map(), venMap = new Map();
  const UNALLOC = "__RCM_UNALLOCATED__";
  let totalRcmAll = 0;

  for (const [doc, lines] of byDoc) {
    const apLine = lines.find((l) => isApLine(l) && l.vendor);
    const party = (apLine && apLine.vendor) || (lines.find((l) => l.vendor) || {}).vendor || "(no party)";
    const hasAp = lines.some(isApLine);
    const tdsTot = lines.filter(isTds).reduce((s, l) => s + l.credit, 0);
    const rcmTot = lines.filter(isRcm).reduce((s, l) => s + l.credit, 0);
    totalRcmAll += rcmTot;
    const expLines = lines.filter((l) => isApExpenseGl(l.gl_code) && l.debit > 0);
    const expTot = expLines.reduce((s, l) => s + l.debit, 0);

    if ((hasAp || rcmTot > 0) && party) {
      let v = venMap.get(party);
      if (!v) { v = { expense: 0, tds: 0, rcm: 0, apCredit: 0, apDebit: 0, docs: new Set(), byHead: new Map() }; venMap.set(party, v); }
      v.expense += expTot; v.tds += tdsTot; v.rcm += rcmTot; v.docs.add(doc);
      for (const l of lines) if (isApLine(l)) { v.apCredit += l.credit; v.apDebit += l.debit; }
      if (expTot > 0) {
        for (const l of expLines) {
          const k = l.gl_code; let h = v.byHead.get(k);
          if (!h) { h = Object.assign(cellAcc(), { name: l.ledger }); v.byHead.set(k, h); }
          const sh = l.debit / expTot;
          h.amount += l.debit; h.tds += tdsTot * sh; h.rcm += rcmTot * sh; h.docs.add(doc);
        }
      } else if (rcmTot > 0) {
        let h = v.byHead.get(UNALLOC);
        if (!h) { h = Object.assign(cellAcc(), { name: "RCM (no expense line)" }); v.byHead.set(UNALLOC, h); }
        h.rcm += rcmTot; h.docs.add(doc);
      }
    }

    if (expTot > 0 && hasAp) {
      for (const l of expLines) {
        const code = l.gl_code; let ex = exMap.get(code);
        if (!ex) { ex = { name: l.ledger, expense: 0, tds: 0, rcm: 0, docs: new Set(), byParty: new Map() }; exMap.set(code, ex); }
        const share = l.debit / expTot;
        ex.expense += l.debit; ex.tds += tdsTot * share; ex.rcm += rcmTot * share; ex.docs.add(doc);
        let pc = ex.byParty.get(party);
        if (!pc) { pc = cellAcc(); ex.byParty.set(party, pc); }
        pc.amount += l.debit; pc.tds += tdsTot * share; pc.rcm += rcmTot * share; pc.docs.add(doc);
      }
    } else if (rcmTot > 0 && !hasAp) {
      let ex = exMap.get(UNALLOC);
      if (!ex) { ex = { name: "RCM (unallocated JV)", expense: 0, tds: 0, rcm: 0, docs: new Set(), byParty: new Map() }; exMap.set(UNALLOC, ex); }
      ex.rcm += rcmTot; ex.docs.add(doc);
      let pc = ex.byParty.get(party);
      if (!pc) { pc = cellAcc(); ex.byParty.set(party, pc); }
      pc.rcm += rcmTot; pc.docs.add(doc);
    }
  }

  const expenseTds = [...exMap.entries()].map(([glCode, ex]) => {
    const effRate = pct(ex.tds, ex.expense);
    return {
      glCode: glCode === UNALLOC ? "—" : glCode, ledger: ex.name,
      section: glCode === UNALLOC ? "" : inferSection(ex.name, effRate),
      expense: ex.expense, tds: ex.tds, effRate, rcm: ex.rcm, rcmRate: pct(ex.rcm, ex.expense),
      vendors: ex.byParty.size, docs: ex.docs.size,
      parties: [...ex.byParty.entries()].map(([party, c]) => ({ party, amount: c.amount, tds: c.tds, rcm: c.rcm, docs: c.docs.size }))
        .sort((a, b) => b.amount + b.rcm - (a.amount + a.rcm)),
    };
  }).filter((r) => r.expense > 0 || r.rcm > 0).sort((a, b) => b.expense - a.expense);

  const vendors = [...venMap.entries()].map(([party, v]) => {
    const heads = [...v.byHead.entries()].map(([code, h]) => ({ glCode: code === UNALLOC ? "—" : code, ledger: h.name, amount: h.amount, tds: h.tds, rcm: h.rcm, docs: h.docs.size }))
      .sort((a, b) => b.amount + b.rcm - (a.amount + a.rcm));
    return { party, expense: v.expense, tds: v.tds, rcm: v.rcm, apCredit: v.apCredit, apDebit: v.apDebit, docs: v.docs.size, topLedger: (heads[0] || {}).ledger || "", heads };
  }).sort((a, b) => b.expense + b.rcm - (a.expense + a.rcm));

  const totalExpense = expenseTds.reduce((s, r) => s + r.expense, 0);
  const totalTds = expenseTds.reduce((s, r) => s + r.tds, 0);
  return { pnl, expenseTds, vendors, totalExpense, totalTds, totalRcm: totalRcmAll, effTdsRate: pct(totalTds, totalExpense) };
}

// ---- Sales & GST ----
const isRevenueLine = (t) => {
  const g = t.gl_code;
  if (g.startsWith("6001") || g.startsWith("6051")) return true;
  const n = t.ledger.toUpperCase();
  return g.startsWith("6") && /(SALES|REVENUE|TURNOVER)/.test(n) && !/COST/.test(n);
};
const isSgstOut = (t) => t.gl_code === "2221013110" || /\bSGST\b.*PAYABLE|OUTPUT.*SGST/i.test(t.ledger);
const isCgstOut = (t) => t.gl_code === "2221013120" || /\bCGST\b.*PAYABLE|OUTPUT.*CGST/i.test(t.ledger);
const isIgstOut = (t) => /\bIGST\b.*PAYABLE|OUTPUT.*IGST/i.test(t.ledger);
const isAr = (t) => t.gl_code.startsWith("1122") || /TRADE REC|RECEIVABLE|DEBTOR|CUSTOMER/i.test(t.ledger);
function bucketOf(rate, hasGst) {
  if (!hasGst) return "0%";
  for (const b of [0, 5, 12, 18, 28]) if (Math.abs(rate - b) < 0.6) return b + "%";
  return "other";
}
export function runSales(txns) {
  const byDoc = new Map();
  for (const t of txns) { const k = t.voucher_no || "row-" + t.rowIndex; const arr = byDoc.get(k); if (arr) arr.push(t); else byDoc.set(k, [t]); }
  const vouchers = [];
  for (const [doc, lines] of byDoc) {
    const taxable = lines.filter(isRevenueLine).reduce((s, l) => s + (l.credit - l.debit), 0);
    if (Math.abs(taxable) < 1) continue;
    const sgst = lines.filter(isSgstOut).reduce((s, l) => s + (l.credit - l.debit), 0);
    const cgst = lines.filter(isCgstOut).reduce((s, l) => s + (l.credit - l.debit), 0);
    const igst = lines.filter(isIgstOut).reduce((s, l) => s + (l.credit - l.debit), 0);
    const gst = sgst + cgst + igst;
    const rate = taxable ? (gst / taxable) * 100 : 0;
    const hasGst = Math.abs(gst) >= 1;
    const customer = (lines.find((l) => isAr(l) && l.vendor) || {}).vendor || (lines.find((l) => l.vendor) || {}).vendor || "(no customer)";
    vouchers.push({ docNo: doc, docType: lines[0].voucher_type, date: lines[0].date, customer, taxable, sgst, cgst, igst, gst, rate: Math.round(rate * 10) / 10, rateBucket: bucketOf(rate, hasGst), hasGst });
  }
  if (!vouchers.length) return { available: false, vouchers: [], rateBuckets: [], customers: [], totalTaxable: 0, totalGst: 0, blendedRate: null, noGstCount: 0, noGstValue: 0 };
  vouchers.sort((a, b) => b.taxable - a.taxable);
  const order = ["0%", "5%", "12%", "18%", "28%", "other"];
  const rbMap = new Map();
  for (const v of vouchers) { let r = rbMap.get(v.rateBucket); if (!r) { r = { rate: v.rateBucket, vouchers: 0, taxable: 0, gst: 0 }; rbMap.set(v.rateBucket, r); } r.vouchers++; r.taxable += v.taxable; r.gst += v.gst; }
  const rateBuckets = [...rbMap.values()].sort((a, b) => order.indexOf(a.rate) - order.indexOf(b.rate));
  const cMap = new Map();
  for (const v of vouchers) {
    let c = cMap.get(v.customer); if (!c) { c = { taxable: 0, gst: 0, vouchers: 0, byRate: new Map() }; cMap.set(v.customer, c); }
    c.taxable += v.taxable; c.gst += v.gst; c.vouchers++;
    let br = c.byRate.get(v.rateBucket); if (!br) { br = { taxable: 0, gst: 0, vouchers: 0 }; c.byRate.set(v.rateBucket, br); }
    br.taxable += v.taxable; br.gst += v.gst; br.vouchers++;
  }
  const customers = [...cMap.entries()].map(([customer, c]) => ({
    customer, taxable: c.taxable, gst: c.gst, rate: c.taxable ? (c.gst / c.taxable) * 100 : null, vouchers: c.vouchers,
    byRate: [...c.byRate.entries()].map(([rate, b]) => ({ rate, taxable: b.taxable, gst: b.gst, vouchers: b.vouchers })).sort((a, b) => order.indexOf(a.rate) - order.indexOf(b.rate)),
  })).sort((a, b) => b.taxable - a.taxable);
  const totalTaxable = vouchers.reduce((s, v) => s + v.taxable, 0);
  const totalGst = vouchers.reduce((s, v) => s + v.gst, 0);
  const noGst = vouchers.filter((v) => v.taxable > 1 && !v.hasGst);
  return { available: true, vouchers, rateBuckets, customers, totalTaxable, totalGst, blendedRate: totalTaxable ? (totalGst / totalTaxable) * 100 : null, noGstCount: noGst.length, noGstValue: noGst.reduce((s, v) => s + v.taxable, 0) };
}

export function analyzeDaybook(txns) {
  return { mis: runMis(txns), sales: runSales(txns) };
}
