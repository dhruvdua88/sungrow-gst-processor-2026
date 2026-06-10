// GSTR-1 Processor engine — client-side, no upload.
// Inputs:  (1) Client GSTR-1 working file  "GSTR-1 Data <month> <yy>.xlsx" — the BOOKS, the base.
//          (2) GST-portal e-invoice dump  EINV_<gstin>_<fy>.xlsx — SUPPORT only (cross-check +
//              enrichment of UQC / description / invoice-type / reverse-charge).
// Output:  invoice summary, Table 12 (HSN), Table 13 (Docs issued), deviation/reconciliation
//          register, validation checks, and a portal-schema GSTR-1 JSON.
//
// The JSON (b2b, cdnr, hsn, doc_issue) is built ENTIRELY from the client Sales sheet, so it
// corresponds 1:1 with the client's data. The e-invoice dump never feeds a number into the JSON.
//
// Scope (current): B2B + Credit/Debit Notes (Registered) only. B2C / exports deferred by design.
import * as XLSX from "xlsx";

// ---------------- portal reference data (from GSTN advisories, 2025-26) ----------------
// Valid GST rate slabs the portal accepts in Table 12 (incl. GST 2.0 slab 40).
export const VALID_RATES = [0, 0.1, 0.25, 1, 1.5, 3, 5, 6, 7.5, 12, 18, 28, 40];
// Portal-standard UQC codes (subset sufficient for goods + OTH fallback for services).
export const VALID_UQC = new Set([
  "PCS", "NOS", "KGS", "MTR", "BOX", "SET", "UNT", "BAG", "OTH", "PRS", "SQM", "BTL", "CTN",
  "DOZ", "KME", "TON", "GMS", "LTR", "KLR", "NA", "PAC", "ROL", "SQF", "SQY", "TBS", "TGM",
  "THD", "TUB", "UGS", "YDS", "BDL", "BKL", "BUN", "CAN", "CBM", "CCM", "CMS", "DRM", "GGK",
  "GRS", "GYD", "MLT", "MTS", "NUM", "PCK", "QTL",
]);
// Table 13 document-nature codes (doc_num in JSON).
export const DOC_NUM = { INVOICE: 1, DEBIT_NOTE: 4, CREDIT_NOTE: 5 };
// 15-char GSTIN pattern (state code + PAN + entity + Z + checksum).
export const GSTIN_RE = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][0-9A-Z]Z[0-9A-Z]$/;
// Valid state codes for POS (01-38 states/UTs + 96 foreign + 97 other territory).
export const VALID_POS = new Set([
  ...Array.from({ length: 38 }, (_, i) => String(i + 1).padStart(2, "0")), "96", "97",
]);

// ---------------- low-level helpers ----------------
const num = (v) => {
  if (v == null || v === "") return 0;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
  return Number.isNaN(n) ? 0 : n;
};
const txt = (v) => (v == null ? "" : String(v).trim());
const round2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
const nz = (n) => Math.abs(n) > 0.005; // non-zero beyond float noise

function gridOf(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
}
// case/space tolerant sheet lookup
function findSheet(wb, ...candidates) {
  const want = candidates.map((c) => c.toLowerCase().replace(/\s+/g, ""));
  for (const name of wb.SheetNames) {
    const norm = name.toLowerCase().replace(/\s+/g, "");
    if (want.includes(norm)) return name;
  }
  return null;
}
// locate the header row index by a marker that must appear in column `col`
function headerRow(grid, col, marker) {
  const m = marker.toLowerCase();
  for (let i = 0; i < Math.min(grid.length, 15); i++) {
    if (txt(grid[i]?.[col]).toLowerCase() === m) return i;
  }
  return -1;
}

const MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
// parse "07-May-2026" or a JS Date -> {dd,mm,yyyy} and "dd-mm-yyyy"
function parseDate(v) {
  if (v instanceof Date) {
    // SheetJS serial→Date conversion can land a few hours before local midnight
    // (timezone fudge) — snap to the nearest day via UTC + 12h to avoid off-by-one.
    const u = new Date(v.getTime() + 12 * 3600 * 1000);
    return { d: u.getUTCDate(), m: u.getUTCMonth() + 1, y: u.getUTCFullYear() };
  }
  const s = txt(v);
  let mm = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (mm) return { d: +mm[1], m: MONTHS[mm[2].toLowerCase()] + 1, y: +mm[3] };
  mm = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if (mm) return { d: +mm[1], m: +mm[2], y: +mm[3] };
  return null;
}
const ddmmyyyy = (v) => {
  const p = parseDate(v);
  if (!p) return txt(v);
  return `${String(p.d).padStart(2, "0")}-${String(p.m).padStart(2, "0")}-${p.y}`;
};
const posCode = (v) => {
  const s = txt(v);
  const m = s.match(/^(\d{2})/);
  return m ? m[1] : s.slice(0, 2);
};
const cleanGstin = (v) => txt(v).replace(/\s/g, "").toUpperCase();

// ---------------- parse the e-invoice dump (SUPPORT / cross-check) ----------------
function parseEinv(wb) {
  const out = { b2b: {}, hsn: [], cdnr: [], scope: { b2c: 0, exp: 0, cdnur: 0 }, errors: [] };

  // b2b, sez, de — per-rate invoice lines
  const b2bName = findSheet(wb, "b2b, sez, de", "b2b,sez,de", "b2b sez de");
  if (!b2bName) { out.errors.push('Sheet "b2b, sez, de" not found in e-invoice dump.'); return out; }
  let g = gridOf(wb, b2bName);
  let h = headerRow(g, 0, "GSTIN/UIN of Recipient");
  if (h < 0) h = 3;
  for (let i = h + 1; i < g.length; i++) {
    const r = g[i];
    if (!r || txt(r[2]) === "") continue;
    const inv = txt(r[2]);
    const inum = inv;
    const d = (out.b2b[inum] ||= {
      inum, gstin: txt(r[0]), name: txt(r[1]), date: r[3], val: 0, pos: txt(r[5]),
      rchrg: txt(r[6]) || "N", invType: txt(r[8]) || "Regular B2B", status: txt(r[18]) || "Valid",
      items: [], txbl: 0, igst: 0, cgst: 0, sgst: 0, cess: 0,
    });
    d.val = num(r[4]); // invoice value (repeated per line — keep last)
    d.status = txt(r[18]) || d.status;
    const it = { rt: num(r[10]), txval: num(r[11]), iamt: num(r[12]), camt: num(r[13]), samt: num(r[14]), csamt: num(r[15]) };
    d.items.push(it);
    d.txbl += it.txval; d.igst += it.iamt; d.cgst += it.camt; d.sgst += it.samt; d.cess += it.csamt;
  }

  // hsn(b2b) — the portal's auto-populated Table 12 (used only to enrich UQC/description)
  const hsnName = findSheet(wb, "hsn(b2b)", "hsn (b2b)");
  if (hsnName) {
    g = gridOf(wb, hsnName);
    h = headerRow(g, 0, "HSN");
    if (h < 0) h = 3;
    for (let i = h + 1; i < g.length; i++) {
      const r = g[i];
      if (!r || txt(r[0]) === "") continue;
      out.hsn.push({
        hsn: txt(r[0]), desc: txt(r[1]), uqc: txt(r[2]), qty: num(r[3]), txval: num(r[4]),
        rt: num(r[5]), iamt: num(r[6]), camt: num(r[7]), samt: num(r[8]), csamt: num(r[9]),
      });
    }
  } else out.errors.push('Sheet "hsn(b2b)" not found — UQC/description enrichment unavailable.');

  // cdnr — credit/debit notes (registered)
  const cdnrName = findSheet(wb, "cdnr");
  if (cdnrName) {
    g = gridOf(wb, cdnrName);
    h = headerRow(g, 0, "GSTIN/UIN of Recipient");
    if (h < 0) h = 3;
    for (let i = h + 1; i < g.length; i++) {
      const r = g[i];
      if (!r || txt(r[2]) === "") continue;
      const ntRaw = txt(r[4]).toUpperCase();
      const ntty = ntRaw.startsWith("D") ? "D" : "C";
      const note = {
        gstin: txt(r[0]), name: txt(r[1]), nt_num: txt(r[2]), nt_dt: r[3], ntty,
        pos: txt(r[5]), rchrg: txt(r[6]) || "N", invType: txt(r[7]) || "Regular B2B", val: num(r[8]),
        status: txt(r[18]) || "Valid",
        items: [{ rt: num(r[10]), txval: num(r[11]), iamt: num(r[12]), camt: num(r[13]), samt: num(r[14]), csamt: num(r[15]) }],
        txbl: num(r[11]), igst: num(r[12]), cgst: num(r[13]), sgst: num(r[14]), cess: num(r[15]),
      };
      out.cdnr.push(note);
    }
  }

  // scope sheets — must be empty in current scope
  const countData = (nm, col, marker) => {
    const s = findSheet(wb, nm);
    if (!s) return 0;
    const gg = gridOf(wb, s);
    const hh = headerRow(gg, col, marker);
    const start = hh < 0 ? 4 : hh + 1;
    let c = 0;
    for (let i = start; i < gg.length; i++) if (gg[i] && txt(gg[i][col]) !== "") c++;
    return c;
  };
  out.scope.b2c = countData("hsn(b2c)", 0, "HSN");
  out.scope.exp = countData("exp", 0, "GSTIN/UIN of Recipient") || countData("exp", 1, "Invoice number");
  out.scope.cdnur = countData("cdnur", 0, "UR Type") || countData("cdnur", 1, "Note Number");
  return out;
}

// ---------------- parse the client Sales working file (THE BASE for the JSON) ----------------
function parseSales(wb) {
  const out = { inv: {}, hsn: {}, hsnB2b: {}, hsnB2c: {}, rows: 0, errors: [] };
  const name = findSheet(wb, "Sales");
  if (!name) { out.errors.push('Sheet "Sales" not found in client GSTR-1 working file.'); return out; }
  const g = gridOf(wb, name);
  let h = headerRow(g, 4, "Invoice no");
  if (h < 0) h = 1;
  for (let i = h + 1; i < g.length; i++) {
    const r = g[i];
    if (!r || txt(r[4]) === "") continue;
    out.rows++;
    const inv = txt(r[4]);
    const d = (out.inv[inv] ||= {
      inum: inv, name: txt(r[6]), gstin: new Set(), rev: new Set(), ship: new Set(),
      txbl: 0, cgst: 0, sgst: 0, igst: 0, tax: 0, val: 0, export: false,
      date: null, byRate: {},
    });
    if (!d.date) d.date = r[5]; // invoice date (first line wins)
    const txbl = num(r[17]); // R Total (taxable)
    d.txbl += txbl; d.cgst += num(r[19]); d.sgst += num(r[20]); d.igst += num(r[21]); d.tax += num(r[28]);
    d.val += num(r[23]); // invoice total (taxable + tax + TCS)
    d.gstin.add(txt(r[24])); d.rev.add(txt(r[26])); d.ship.add(txt(r[25]));
    if (/export|sez/i.test(txt(r[25])) || /export|sez/i.test(txt(r[26]))) d.export = true; // shipment / type flags export
    // Rate normalisation: component (CGST+SGST) rows show half rate; IGST rows show full rate.
    const s = num(r[18]), cg = num(r[19]), sg = num(r[20]), ig = num(r[21]);
    const totalRate = round2(cg || sg ? s * 2 : s);
    // Per-invoice rate-wise items — these become the JSON b2b/cdnr itm_det lines.
    const ri = (d.byRate[totalRate] ||= { rt: totalRate, txval: 0, iamt: 0, camt: 0, samt: 0, csamt: 0 });
    ri.txval += txbl; ri.iamt += ig; ri.camt += cg; ri.samt += sg;
    // HSN rebuild — aggregate into the combined map AND a B2B/B2C bucket (registered => B2B).
    // Zero-value lines (no taxable, no tax) are dispatch records, not supplies — skip in Table 12.
    if (!nz(txbl) && !nz(ig + cg + sg)) continue;
    const code = txt(r[32]) || "(blank)";
    const hk = `${code}|${totalRate}`;
    const registered = cleanGstin(r[24]).length >= 15;
    const bucket = registered ? out.hsnB2b : out.hsnB2c;
    for (const store of [out.hsn, bucket]) {
      const hh = (store[hk] ||= { hsn: code, rt: totalRate, qty: 0, txval: 0, iamt: 0, camt: 0, samt: 0 });
      hh.qty += num(r[12]); hh.txval += txbl; hh.iamt += ig; hh.camt += cg; hh.samt += sg;
    }
  }
  return out;
}

// ---------------- bird's-eye summary of the client's working file ----------------
function fileSummary(wb) {
  const out = {
    sales: { invoices: 0, lines: 0, txbl: 0, tax: 0,
      byRevenue: {}, byShipment: {}, byCustomer: { "Registered (B2B)": blk(), "Unregistered (B2C)": blk() } },
    foc: { lines: 0, qty: 0, amt: 0 },
    srv: { lines: 0, dispatches: 0, qty: 0 },
    support: { summaryRows: 0, stockRows: 0 },
    sheets: wb.SheetNames.slice(),
  };
  function blk() { return { invoices: new Set(), lines: 0, txbl: 0, tax: 0 }; }

  // Sales
  const sName = findSheet(wb, "Sales");
  if (sName) {
    const g = gridOf(wb, sName);
    let h = headerRow(g, 4, "Invoice no"); if (h < 0) h = 1;
    const seen = new Set();
    for (let i = h + 1; i < g.length; i++) {
      const r = g[i]; if (!r || txt(r[4]) === "") continue;
      const inv = txt(r[4]), txbl = num(r[17]), tax = num(r[28]);
      out.sales.lines++; out.sales.txbl += txbl; out.sales.tax += tax; seen.add(inv);
      const rv = txt(r[26]) || "(blank)";
      const sh = txt(r[25]) || "(blank)";
      const isExport = /export|sez/i.test(sh) || /export|sez/i.test(txt(r[26]));
      const shipKey = isExport ? "Export / SEZ" : (sh || "Domestic");
      const reg = txt(r[24]).length >= 15 ? "Registered (B2B)" : "Unregistered (B2C)";
      for (const [bucket, key] of [[out.sales.byRevenue, rv], [out.sales.byShipment, shipKey]]) {
        const b = (bucket[key] ||= blk());
        b.lines++; b.txbl += txbl; b.tax += tax; b.invoices.add(inv);
      }
      const c = out.sales.byCustomer[reg]; c.lines++; c.txbl += txbl; c.tax += tax; c.invoices.add(inv);
    }
    out.sales.invoices = seen.size;
  }
  // SRV FOC (free of cost)
  const focName = findSheet(wb, "SRV FOC", "srvfoc");
  if (focName) {
    const g = gridOf(wb, focName);
    let h = headerRow(g, 0, "Pstng Date"); if (h < 0) h = 2;
    for (let i = h + 1; i < g.length; i++) {
      const r = g[i]; if (!r || txt(r[1]) === "") continue;
      out.foc.lines++; out.foc.qty += num(r[3]); out.foc.amt += num(r[15]);
    }
  }
  // SRV details
  const srvName = findSheet(wb, "SRV details", "srvdetails");
  if (srvName) {
    const g = gridOf(wb, srvName);
    let h = headerRow(g, 0, "SL No"); if (h < 0) h = 0;
    const disp = new Set();
    for (let i = h + 1; i < g.length; i++) {
      const r = g[i]; if (!r || txt(r[0]) === "") continue;
      out.srv.lines++; out.srv.qty += num(r[11]); if (txt(r[3])) disp.add(txt(r[3]));
    }
    out.srv.dispatches = disp.size;
  }
  const countRows = (nm) => { const s = findSheet(wb, nm); if (!s) return 0; const g = gridOf(wb, s); return g.filter((r) => r && r.some((c) => c != null && c !== "")).length; };
  out.support.summaryRows = countRows("Summary");
  out.support.stockRows = countRows("Physical stock at WH") || countRows("Physical stock");

  // serialise sets -> counts, round
  const fin = (b) => ({ invoices: b.invoices.size, lines: b.lines, txbl: round2(b.txbl), tax: round2(b.tax) });
  out.sales.byRevenue = Object.fromEntries(Object.entries(out.sales.byRevenue).map(([k, v]) => [k, fin(v)]));
  out.sales.byShipment = Object.fromEntries(Object.entries(out.sales.byShipment).map(([k, v]) => [k, fin(v)]));
  out.sales.byCustomer = Object.fromEntries(Object.entries(out.sales.byCustomer).map(([k, v]) => [k, fin(v)]));
  out.sales.txbl = round2(out.sales.txbl); out.sales.tax = round2(out.sales.tax);
  out.foc.amt = round2(out.foc.amt);
  return out;
}

// ---------------- portal-schema validation of a JSON document (b2b invoice / credit note) ----------------
// Returns { field: errorMessage }. These rules are what makes the GST portal REJECT the JSON;
// every one of them is fixable in the app's "Fix & Generate" editor.
export function validateDocRow(doc, { supplierState = "", fp = "" } = {}) {
  const errs = {};
  const ctin = cleanGstin(doc.ctin);
  if (!GSTIN_RE.test(ctin)) errs.ctin = "Recipient GSTIN invalid (15-char GSTIN pattern)";
  if (!/^\d{2}-\d{2}-\d{4}$/.test(txt(doc.idt))) errs.idt = "Date must be dd-mm-yyyy";
  else if (/^(\d{2})(\d{4})$/.test(fp)) {
    const [, fm, fy] = fp.match(/^(\d{2})(\d{4})$/);
    const [, , dm, dy] = txt(doc.idt).match(/^(\d{2})-(\d{2})-(\d{4})$/);
    if (+dy > +fy || (+dy === +fy && +dm > +fm)) errs.idt = "Document date is after the return period";
  }
  if (!VALID_POS.has(txt(doc.pos))) errs.pos = "POS must be a valid 2-digit state code";
  if (!(num(doc.val) > 0)) errs.val = "Document value must be > 0";
  if (!["R", "SEWP", "SEWOP", "DE"].includes(txt(doc.invTyp))) errs.invTyp = "Invoice type must be R / SEWP / SEWOP / DE";
  const intra = txt(doc.pos) === supplierState && doc.invTyp === "R";
  const itemErrs = [];
  for (const it of doc.items || []) {
    if (!VALID_RATES.includes(Number(it.rt))) itemErrs.push(`rate ${it.rt} not a valid GST slab`);
    if (it.txval < -0.005) itemErrs.push(`negative taxable at rate ${it.rt}`);
    if (intra && nz(it.iamt)) itemErrs.push(`POS ${doc.pos} = supplier state but IGST charged at rate ${it.rt}`);
    if (!intra && nz(it.camt + it.samt)) itemErrs.push(`POS ${doc.pos} is inter-state/SEZ but CGST/SGST charged at rate ${it.rt}`);
    const expect = (Math.abs(it.txval) * Number(it.rt)) / 100;
    const got = Math.abs(it.iamt) + Math.abs(it.camt) + Math.abs(it.samt);
    if (Math.abs(expect - got) > Math.max(1, expect * 0.02)) itemErrs.push(`tax ≠ taxable × ${it.rt}%`);
  }
  if (itemErrs.length) errs.items = itemErrs.join("; ");
  return errs;
}

// ---------------- main ----------------
export function processGstr1(einvWb, salesWb, opts = {}) {
  const einv = parseEinv(einvWb);
  const sales = parseSales(salesWb);
  const fileOverview = fileSummary(salesWb);
  const errors = [...einv.errors, ...sales.errors];

  const gstin = opts.gstin || ""; // supplier GSTIN is taken from the EINV filename by the caller
  const fp = opts.fp || ""; // "MMYYYY"
  const supplierState = gstin.slice(0, 2);

  // ----- e-invoice side (SUPPORT) -----
  const b2bAll = Object.values(einv.b2b);
  const cancelled = b2bAll.filter((d) => d.status.toLowerCase() !== "valid").map((d) => d.inum).sort();
  const validInv = b2bAll.filter((d) => d.status.toLowerCase() === "valid");
  const validB2bTxbl = sum(validInv, "txbl");
  const cdnTxbl = einv.cdnr.reduce((a, c) => a + c.txbl, 0);

  // ----- classify the client's Sales documents (THE BASE for the JSON) -----
  const salesDocs = Object.values(sales.inv);
  const isReg = (d) => [...d.gstin].some((g) => cleanGstin(g).length >= 15);
  const isZero = (d) => !nz(d.txbl) && !nz(d.cgst + d.sgst + d.igst);
  const salesCn = salesDocs.filter((d) => d.txbl < -0.005);               // negative book docs → credit notes
  const zeroDocs = salesDocs.filter(isZero).map((d) => d.inum).sort();    // zero-value dispatch docs → Table 13 only
  const salesB2b = salesDocs.filter((d) => d.txbl > 0.005 && isReg(d) && !d.export);
  const salesB2c = salesDocs.filter((d) => d.txbl > 0.005 && !isReg(d) && !d.export);
  const salesExp = salesDocs.filter((d) => d.txbl > 0.005 && d.export);

  // Build the JSON-source document objects from the books, enriched (never valued) from EINV.
  const toJsonDoc = (d, kind) => {
    const e = kind === "CN"
      ? einv.cdnr.find((c) => c.nt_num === d.inum)
      : einv.b2b[d.inum];
    const ctin = cleanGstin([...d.gstin].map(cleanGstin).find((g) => g.length >= 15) || "");
    const abs = kind === "CN" ? -1 : 1; // notes carry positive numbers in the JSON
    const items = Object.values(d.byRate)
      .filter((it) => nz(it.txval) || nz(it.iamt + it.camt + it.samt))
      .map((it) => ({
        rt: it.rt, txval: round2(abs * it.txval), iamt: round2(abs * it.iamt),
        camt: round2(abs * it.camt), samt: round2(abs * it.samt), csamt: 0,
      }))
      .sort((a, b) => a.rt - b.rt);
    const val = round2(abs * (nz(d.val) ? d.val : d.txbl + d.tax));
    return {
      kind, inum: d.inum, name: d.name, ctin,
      idt: ddmmyyyy(d.date || (e && (e.date || e.nt_dt)) || ""),
      val, pos: ctin.slice(0, 2) || posCode(e?.pos || ""),
      rchrg: e?.rchrg === "Y" ? "Y" : "N",
      invTyp: e ? mapInvType(e.invType) : ([...d.ship].some((s) => /sez/i.test(s)) ? "SEWP" : "R"),
      items,
      txbl: round2(abs * d.txbl),
      einvStatus: e ? (e.status || "Valid") : "not e-invoiced",
    };
  };
  const jsonB2b = salesB2b.map((d) => toJsonDoc(d, "INV")).sort((a, b) => a.inum.localeCompare(b.inum));
  const jsonCdnr = salesCn.map((d) => toJsonDoc(d, "CN")).sort((a, b) => a.inum.localeCompare(b.inum));
  const jsonDocs = [...jsonB2b, ...jsonCdnr];

  // Per-document portal validation — every error here is editable in the app before export.
  const docErrors = jsonDocs
    .map((doc) => ({ inum: doc.inum, kind: doc.kind, errs: validateDocRow(doc, { supplierState, fp }) }))
    .filter((x) => Object.keys(x.errs).length);

  const jsonB2bTxbl = sum(jsonB2b, "txbl");
  const jsonCdnTxbl = jsonCdnr.reduce((a, c) => a + c.txbl, 0); // negative (books sign)

  // ----- Table 12 (BASE = client Sales sheet; e-invoice dump is SUPPORT only) -----
  // SAC (service) codes start with "99" → UQC must be "NA" (portal rejects PCS *and* OTH →
  // RET191353) and qty must be 0 (RET191355).
  const isSac = (hsn) => /^99/.test(String(hsn));
  const normUqc = (hsn, uqc) => (isSac(hsn) ? "NA" : (uqc || "OTH").toUpperCase());
  const normQty = (hsn, qty) => (isSac(hsn) ? 0 : qty);
  // Per-HSN UQC + description lookup from the e-invoice support (first occurrence wins).
  const support = {};
  for (const h of einv.hsn) if (!support[h.hsn]) support[h.hsn] = { uqc: h.uqc, desc: h.desc };
  const toT12 = (hsnMap) => Object.values(hsnMap)
    .map((h) => {
      const sup = support[h.hsn] || {};
      return {
        hsn: h.hsn, desc: sup.desc || "", uqc: normUqc(h.hsn, sup.uqc || ""), qty: normQty(h.hsn, h.qty),
        rt: h.rt, txval: round2(h.txval), iamt: round2(h.iamt), camt: round2(h.camt), samt: round2(h.samt), csamt: 0,
      };
    })
    .sort((a, b) => String(a.hsn).localeCompare(String(b.hsn)) || a.rt - b.rt)
    .map((h, i) => ({ sr: i + 1, ...h, total: round2(h.txval + h.iamt + h.camt + h.samt + h.csamt) }));
  const t12 = toT12(sales.hsn);          // combined Table 12 (all sales rows) — what the client sees
  const t12B2b = toT12(sales.hsnB2b);    // for JSON hsn_b2b
  const t12B2c = toT12(sales.hsnB2c);    // for JSON hsn_b2c
  const t12Txbl = sum(t12, "txval");
  const t12Tax = t12.reduce((a, h) => a + h.iamt + h.camt + h.samt, 0);

  // ----- Table 13 (documents issued — from the books) -----
  const cnNums = new Set([...salesCn.map((d) => d.inum), ...einv.cdnr.map((c) => c.nt_num)]);
  const allNums = new Set(Object.keys(sales.inv).filter((k) => !cnNums.has(k)));
  const invNums = [...allNums].filter((x) => /^\d+$/.test(x)).map(Number).sort((a, b) => a - b);
  const invFrom = invNums[0], invTo = invNums[invNums.length - 1];
  const invTotal = invNums.length;
  // Cancelled docs = portal-cancelled IRNs that the books DON'T carry (book copies stay issued).
  const cancelledDocs = cancelled.filter((c) => !sales.inv[c]);
  const invCancelled = cancelledDocs.filter((c) => /^\d+$/.test(c)).length;
  const invNet = invTotal; // every doc in the books is an issued doc; cancelled = portal-only IRNs
  const presentSet = new Set(invNums);
  const cdnNumeric = new Set([...cnNums].filter((x) => /^\d+$/.test(x)).map(Number));
  const cancNumeric = new Set(cancelledDocs.filter((x) => /^\d+$/.test(x)).map(Number));
  const gaps = [];
  for (let n = invFrom; n <= invTo; n++) if (!presentSet.has(n) && !cdnNumeric.has(n) && !cancNumeric.has(n)) gaps.push(n);

  const t13 = [{
    code: DOC_NUM.INVOICE, nature: "Invoices for outward supply",
    from: String(invFrom), to: String(invTo),
    total: invTotal + invCancelled, cancel: invCancelled, net: invNet,
  }];
  if (jsonCdnr.length || einv.cdnr.length) {
    const cnAll = [...new Set([...salesCn.map((d) => d.inum), ...einv.cdnr.map((c) => c.nt_num)])].sort();
    const dNums = new Set(einv.cdnr.filter((c) => c.ntty === "D").map((c) => c.nt_num));
    const cCount = cnAll.filter((n) => !dNums.has(n)).length;
    const dCount = cnAll.length - cCount;
    if (cCount) t13.push({ code: DOC_NUM.CREDIT_NOTE, nature: "Credit Note", from: cnAll[0], to: cnAll[cnAll.length - 1], total: cCount, cancel: 0, net: cCount });
    if (dCount) t13.push({ code: DOC_NUM.DEBIT_NOTE, nature: "Debit Note", from: cnAll[0], to: cnAll[cnAll.length - 1], total: dCount, cancel: 0, net: dCount });
  }

  // ----- reconciliation / deviations: books vs portal e-invoices -----
  const missingInSales = validInv.filter((d) => !sales.inv[d.inum]).map((d) => d.inum);
  const valMismatch = [];
  for (const d of validInv) {
    const s = sales.inv[d.inum];
    if (s && Math.abs(s.txbl - d.txbl) > 1) {
      valMismatch.push({ inum: d.inum, name: d.name, sales: round2(s.txbl), portal: round2(d.txbl), diff: round2(s.txbl - d.txbl) });
    }
  }
  // registered B2B in the books with NO valid IRN (and not zero-value) — e-invoicing omission
  const validSet = new Set(validInv.map((d) => d.inum));
  const notEinvoiced = salesB2b.filter((d) => !validSet.has(d.inum)).map((d) => ({
    inum: d.inum, name: d.name, txbl: round2(d.txbl),
    portalStatus: einv.b2b[d.inum]?.status || "no IRN",
  }));
  // invoices the client tagged Non-Revenue but which carry a valid IRN -> belong in GSTR-1
  const nonRevButEinvoiced = validInv
    .filter((d) => sales.inv[d.inum] && !sales.inv[d.inum].rev.has("Yes-Revenue"))
    .map((d) => ({ inum: d.inum, name: d.name, txbl: round2(d.txbl), rev: [...sales.inv[d.inum].rev].join(", ") }))
    .sort((a, b) => b.txbl - a.txbl);
  const salesValidTxbl = validInv.reduce((a, d) => a + (sales.inv[d.inum]?.txbl || 0), 0);

  // ----- Sales-sheet → JSON reconciliation bridge -----
  // The client Sales sheet is the BASE. Every rupee must be accounted for down to the JSON.
  const cancelledSet = new Set(cancelled);
  let b2cBook = 0, exportBook = 0, zeroBook = 0, cnBook = 0;
  for (const d of salesB2c) b2cBook += d.txbl;
  for (const d of salesExp) exportBook += d.txbl;
  for (const d of salesCn) cnBook += d.txbl;
  const salesSheetTxbl = fileOverview.sales.txbl;
  const portalSupportTxbl = validB2bTxbl - cdnTxbl; // the portal hsn(b2b) support figure
  const cancelledBook = salesB2b.filter((d) => cancelledSet.has(d.inum)).reduce((a, d) => a + d.txbl, 0);
  const notEinvBook = notEinvoiced.reduce((a, d) => a + d.txbl, 0) - cancelledBook;
  const bookVsPortalAdj = salesValidTxbl - validB2bTxbl; // C14 net (book minus portal on the valid set)
  // base − b2c − export − JSON-CN = JSON b2b (what we file). Separately the support bridge:
  // JSON b2b − cancelled − notEinv − adj = portal valid e-invoices.
  const bridge = {
    salesSheetTxbl: round2(salesSheetTxbl),
    table12Txbl: round2(t12Txbl),
    jsonB2bTxbl: round2(jsonB2bTxbl),
    jsonCdnTxbl: round2(jsonCdnTxbl),
    less: {
      b2c: round2(b2cBook), export: round2(exportBook), cancelled: round2(cancelledBook),
      notEinvoiced: round2(notEinvBook), cdnInBook: round2(cnBook), bookVsPortalAdj: round2(bookVsPortalAdj), cdn: round2(cdnTxbl),
    },
    portalSupportTxbl: round2(portalSupportTxbl),
    // Books base − B2C − exports − CN-in-books = JSON B2B (identity, residual0)
    residual0: round2(salesSheetTxbl - b2cBook - exportBook - cnBook - zeroBook - jsonB2bTxbl),
    // JSON B2B − cancelled-IRN docs − not-e-invoiced − book-vs-portal adj − portal CDN = portal support
    residual: round2(jsonB2bTxbl - cancelledBook - notEinvBook - bookVsPortalAdj - cdnTxbl - portalSupportTxbl),
  };

  // ----- validation checks -----
  const checks = [];
  const add = (id, desc, expected, actual, ok) => checks.push({ id, desc, expected: String(expected), actual: String(actual), ok: !!ok });
  const badLen = t12.filter((h) => !(/^\d+$/.test(h.hsn) && (h.hsn.length === 6 || h.hsn.length === 8))).map((h) => h.hsn);
  add("C1", "Every Table-12 HSN/SAC numeric & 6 or 8 digits (AATO>5cr ⇒ 6 min)", "0 invalid", `${badLen.length} ${badLen.slice(0, 6)}`, !badLen.length);
  const badRate = t12.filter((h) => !VALID_RATES.includes(h.rt)).map((h) => `${h.hsn}@${h.rt}`);
  add("C2", "Every Table-12 rate is a valid GST slab", "0 invalid", `${badRate.length} ${badRate.slice(0, 6)}`, !badRate.length);
  const badUqc = t12.filter((h) => h.uqc && !VALID_UQC.has(h.uqc.toUpperCase())).map((h) => `${h.hsn}:${h.uqc}`);
  add("C3", "Every UQC is a portal-standard code", "0 invalid", `${badUqc.length} ${badUqc.slice(0, 6)}`, !badUqc.length);
  const badSacUqc = t12.filter((h) => /^99/.test(h.hsn) && String(h.uqc).toUpperCase() !== "NA").map((h) => `${h.hsn}:${h.uqc}`);
  add("C3b", "Service SAC (99…) rows use UQC NA — OTH is rejected too (RET191353)", "0 invalid", `${badSacUqc.length} ${badSacUqc.slice(0, 6)}`, !badSacUqc.length);
  const badSacQty = t12.filter((h) => /^99/.test(h.hsn) && round2(h.qty) !== 0).map((h) => `${h.hsn}:${h.qty}`);
  add("C3c", "Service SAC (99…) rows have qty 0 (RET191355)", "0 invalid", `${badSacQty.length} ${badSacQty.slice(0, 6)}`, !badSacQty.length);
  const badGoodsUqc = t12.filter((h) => !/^99/.test(h.hsn) && String(h.uqc).toUpperCase() === "NA").map((h) => h.hsn);
  add("C3d", "Goods (non-99) rows use a real UQC, not NA", "0 invalid", `${badGoodsUqc.length} ${badGoodsUqc.slice(0, 6)}`, !badGoodsUqc.length);
  const badGoodsQty = t12.filter((h) => !/^99/.test(h.hsn) && round2(h.qty) <= 0).map((h) => h.hsn);
  add("C3e", "Goods (non-99) rows have qty > 0", "0 invalid", `${badGoodsQty.length} ${badGoodsQty.slice(0, 6)}`, !badGoodsQty.length);
  const negT12 = t12.filter((h) => h.txval < -0.005).map((h) => `${h.hsn}@${h.rt}`);
  add("C3f", "No Table-12 row nets negative (CDN larger than sales for an HSN/rate)", "0 negative", `${negT12.length} ${negT12.slice(0, 6)}`, !negT12.length);
  const noDesc = t12.filter((h) => !h.desc).map((h) => h.hsn);
  add("C4", "Every HSN row has a description (mandatory Phase-3)", "0 blank", `${noDesc.length} ${noDesc.slice(0, 6)}`, !noDesc.length);
  const arith = t12.filter((h) => { const e = (Math.abs(h.txval) * h.rt) / 100; return Math.abs(e - Math.abs(h.iamt + h.camt + h.samt)) > Math.max(1, e * 0.02); }).map((h) => `${h.hsn}@${h.rt}`);
  add("C5", "Per-row tax == taxable × rate (±2% rounding)", "0 breaks", `${arith.length} ${arith.slice(0, 6)}`, !arith.length);
  const csBreak = t12.filter((h) => Math.abs(h.camt - h.samt) > 1).map((h) => h.hsn);
  add("C6", "CGST == SGST on every intra-state row", "0 breaks", `${csBreak.length}`, !csBreak.length);
  add("C7", "Table-12 taxable == client Sales sheet taxable (books drive Table 12)", fmt(fileOverview.sales.txbl), fmt(t12Txbl), Math.abs(t12Txbl - fileOverview.sales.txbl) < 1);
  add("C8", "Table-12 tax == client Sales sheet tax (cess shown separately)", fmt(fileOverview.sales.tax), fmt(t12Tax), Math.abs(t12Tax - fileOverview.sales.tax) < Math.max(2, fileOverview.sales.tax * 0.005));
  add("C9", "Portal-cancelled IRNs that the books still carry — confirm with client", cancelled.length ? `confirm ${cancelled.filter((c) => sales.inv[c]).length}` : "none", `${cancelled.join(", ") || "none"}`, true);
  add("C10", "Table-13 net == total − cancelled", `${invTotal + invCancelled}-${invCancelled}=${invNet}`, `${invNet}`, invNet === invTotal);
  add("C11", "Invoice serial series continuous (gaps = CDN/cancelled only)", "0 unexplained gaps", `${gaps.length} ${gaps.slice(0, 8)}`, !gaps.length);
  add("C12", "B2C / exports in books — invoice-level JSON sections for these still pending", "0 (module pending)", `B2C ${salesB2c.length} docs ${fmt(b2cBook)} · Exports ${salesExp.length} docs ${fmt(exportBook)}`, !salesB2c.length && !salesExp.length);
  add("C13", "Every valid e-invoice present in client Sales book", "0 missing", `${missingInSales.length} ${missingInSales.slice(0, 6)}`, !missingInSales.length);
  add("C14", "Per-invoice taxable: client book == portal e-invoice", "0 mismatch", `${valMismatch.length} ${valMismatch.slice(0, 4).map((m) => m.inum)}`, !valMismatch.length);
  add("C15", "Sales sheet fully reconciles: books → JSON → portal support (both residuals 0)", "₹0 · ₹0", `${fmt(bridge.residual0)} · ${fmt(bridge.residual)}`, Math.abs(bridge.residual0) < 1 && Math.abs(bridge.residual) < 1);
  add("C16", "All portal credit/debit notes are in the sales sheet (Table 12 net of CDN)", "0 unmatched", `${einv.cdnr.filter((c) => !sales.inv[c.nt_num]).length}`, !einv.cdnr.filter((c) => !sales.inv[c.nt_num]).length);
  add("C17", "JSON B2B taxable == client Sales B2B taxable (JSON built from the books)", fmt(salesB2b.reduce((a, d) => a + d.txbl, 0)), fmt(jsonB2bTxbl), Math.abs(jsonB2bTxbl - salesB2b.reduce((a, d) => a + d.txbl, 0)) < 1);
  add("C18", "JSON B2B − credit notes taxable == Table-12 hsn_b2b taxable", fmt(sum(t12B2b, "txval")), fmt(jsonB2bTxbl - jsonCdnTxbl), Math.abs(jsonB2bTxbl - jsonCdnTxbl - sum(t12B2b, "txval")) < 1);
  add("C19", "Zero-value dispatch documents excluded from B2B JSON (kept in Table 13)", "info", `${zeroDocs.length} docs: ${zeroDocs.slice(0, 8)}${zeroDocs.length > 8 ? "…" : ""}`, true);
  const docErrCells = docErrors.reduce((a, x) => a + Object.keys(x.errs).length, 0);
  add("C20", "Every JSON document passes portal field validation (GSTIN/date/POS/value/tax-split)", "0 errors", `${docErrCells} error(s) across ${docErrors.length} doc(s): ${docErrors.slice(0, 4).map((x) => x.inum)}`, !docErrors.length);

  const fails = checks.filter((c) => !c.ok);
  const blocking = new Set(["C1", "C2", "C3", "C3f", "C4", "C7", "C8", "C10", "C12", "C17", "C18", "C20"]); // portal-blocking vs deviation-only
  const blockingFails = fails.filter((c) => blocking.has(c.id));
  const allPass = fails.length === 0;
  const portalReady = blockingFails.length === 0;

  // ----- GSTR-1 JSON (portal schema) — built ENTIRELY from the client's books -----
  const json = buildJson({ gstin, fp, b2bDocs: jsonB2b, cnDocs: jsonCdnr, t12, t12B2b, t12B2c, t13, version: opts.version || "GST3.1.2" });

  return {
    meta: { gstin, fp, period: opts.periodLabel || "", entity: opts.entity || "", supplierName: opts.entity || "", supplierState },
    scope: einv.scope,
    counts: {
      einvB2bInvoices: b2bAll.length, validInvoices: validInv.length, cancelled: cancelled.length,
      cdnr: jsonCdnr.length, salesInvoices: Object.keys(sales.inv).length, salesRows: sales.rows, hsnRows: t12.length,
      jsonB2bInvoices: jsonB2b.length, zeroDocs: zeroDocs.length,
    },
    totals: {
      t12Txbl: round2(t12Txbl), t12Igst: round2(sum(t12, "iamt")), t12Cgst: round2(sum(t12, "camt")), t12Sgst: round2(sum(t12, "samt")), t12Cess: round2(sum(t12, "csamt")),
      validB2bTxbl: round2(validB2bTxbl), cdnTxbl: round2(cdnTxbl), salesValidTxbl: round2(salesValidTxbl),
      jsonB2bTxbl: round2(jsonB2bTxbl), jsonCdnTxbl: round2(jsonCdnTxbl),
    },
    table12: t12, table13: t13, bridge,
    cancelled, cdnr: einv.cdnr,
    jsonDocs, docErrors, zeroDocs,
    cancelledDetail: cancelled.map((inum) => ({
      inum, name: sales.inv[inum]?.name || einv.b2b[inum]?.name || "",
      gstin: sales.inv[inum] ? [...sales.inv[inum].gstin][0] || "" : (einv.b2b[inum]?.gstin || ""),
      txbl: round2(sales.inv[inum]?.txbl || 0), portalStatus: einv.b2b[inum]?.status || "",
    })).sort((a, b) => b.txbl - a.txbl),
    recon: { missingInSales, valMismatch, nonRevButEinvoiced, notEinvoiced, salesValidTxbl: round2(salesValidTxbl), portalValidTxbl: round2(validB2bTxbl), diff: round2(salesValidTxbl - validB2bTxbl) },
    fileOverview,
    checks, fails, blockingFails, allPass, portalReady, errors,
    json,
  };
}

function sum(arr, key) { return arr.reduce((a, x) => a + (x[key] || 0), 0); }
function fmt(n) { return Number(n).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

// ---------------- JSON builder — every number comes from the client's books ----------------
export function buildJson({ gstin, fp, b2bDocs, cnDocs, t12, t12B2b, t12B2c, t13, version }) {
  // b2b grouped by recipient GSTIN
  const byCtin = {};
  for (const d of b2bDocs) (byCtin[d.ctin] ||= []).push(d);
  const b2b = Object.entries(byCtin).map(([ctin, docs]) => ({
    ctin,
    inv: docs.map((d) => ({
      inum: d.inum, idt: d.idt, val: round2(d.val), pos: d.pos,
      rchrg: d.rchrg === "Y" ? "Y" : "N", inv_typ: d.invTyp,
      itms: d.items.map((a, i) => ({ num: i + 1, itm_det: { rt: a.rt, txval: round2(a.txval), iamt: round2(a.iamt), camt: round2(a.camt), samt: round2(a.samt), csamt: round2(a.csamt) } })),
    })),
  }));

  // cdnr grouped by recipient GSTIN (positive numbers; ntty determines sign on the portal)
  const cByCtin = {};
  for (const c of cnDocs) (cByCtin[c.ctin] ||= []).push(c);
  const cdnrJson = Object.entries(cByCtin).map(([ctin, notes]) => ({
    ctin,
    nt: notes.map((c) => ({
      ntty: c.ntty || "C", nt_num: c.inum, nt_dt: c.idt, val: round2(Math.abs(c.val)), pos: c.pos,
      rchrg: c.rchrg === "Y" ? "Y" : "N", inv_typ: c.invTyp,
      itms: c.items.map((it, i) => ({ num: i + 1, itm_det: { rt: it.rt, txval: round2(Math.abs(it.txval)), iamt: round2(Math.abs(it.iamt)), camt: round2(Math.abs(it.camt)), samt: round2(Math.abs(it.samt)), csamt: 0 } })),
    })),
  }));

  // hsn (Table 12) — sourced from the client Sales sheet (books drive Table 12).
  // Schema changed on 2025-05-01: periods >= May-2025 use { hsn_b2b, hsn_b2c };
  // earlier periods use { data }. Description is capped at 30 chars by the portal.
  const toHsnRows = (rows) => rows.map((h, i) => ({
    num: i + 1, hsn_sc: h.hsn, desc: String(h.desc || "").slice(0, 30), uqc: (h.uqc || "OTH").toUpperCase(),
    qty: round2(h.qty), txval: round2(h.txval), rt: h.rt,
    iamt: round2(h.iamt), camt: round2(h.camt), samt: round2(h.samt), csamt: round2(h.csamt),
  }));
  const hsn = isBifurcated(fp)
    ? { hsn_b2b: toHsnRows(t12B2b), hsn_b2c: toHsnRows(t12B2c) }
    : { data: toHsnRows(t12) };

  // doc_issue (Table 13)
  const doc_issue = {
    doc_det: t13.map((d) => ({
      doc_num: d.code,
      docs: [{ num: 1, from: d.from, to: d.to, totnum: d.total, cancel: d.cancel, net_issue: d.net }],
    })),
  };

  const obj = { gstin, fp, version, hash: "hash" };
  if (b2b.length) obj.b2b = b2b;
  if (cdnrJson.length) obj.cdnr = cdnrJson;
  obj.hsn = hsn;
  obj.doc_issue = doc_issue;
  return obj;
}
// Table-12 HSN B2B/B2C bifurcation applies to filing periods on/after May 2025.
function isBifurcated(fp) {
  const m = /^(\d{2})(\d{4})$/.exec(String(fp || ""));
  if (!m) return true; // default to current (bifurcated) shape
  const mm = +m[1], yyyy = +m[2];
  return yyyy > 2025 || (yyyy === 2025 && mm >= 5);
}
function mapInvType(t) {
  const s = (t || "").toLowerCase();
  if (s.includes("sez") && s.includes("without")) return "SEWOP";
  if (s.includes("sez")) return "SEWP";
  if (s.includes("deemed")) return "DE";
  return "R";
}
