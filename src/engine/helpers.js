// Generic helpers ported 1:1 from build_sipl_gst_itc_tracker.py

export const TARGET_COMPANY_CODE = "2081";
export const RECIPIENT_GSTIN = "29AAXCS2197N1Z4"; // SIPL Karnataka - the entity whose books these are

// Reverse-charge ledger evidence
export const RCM_ACCOUNTS_LEDGER = new Set(["2221013010", "2221013030", "2221013040"]);
// REC accounts (ITC receivable side)
export const REC_ACCOUNTS = { "2221013210": "SGST", "2221013220": "CGST", "2221013230": "IGST" };
// RCM payable accounts
export const RCM_ACCOUNTS = { "2221013010": "IGST", "2221013030": "SGST", "2221013040": "CGST" };

export const STATE_CODE = {
  "01": "Jammu and Kashmir", "02": "Himachal Pradesh", "03": "Punjab", "04": "Chandigarh",
  "05": "Uttarakhand", "06": "Haryana", "07": "Delhi", "08": "Rajasthan", "09": "Uttar Pradesh",
  "10": "Bihar", "11": "Sikkim", "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur",
  "15": "Mizoram", "16": "Tripura", "17": "Meghalaya", "18": "Assam", "19": "West Bengal",
  "20": "Jharkhand", "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh", "24": "Gujarat",
  "25": "Daman and Diu", "26": "Dadra and Nagar Haveli and Daman and Diu", "27": "Maharashtra",
  "28": "Andhra Pradesh", "29": "Karnataka", "30": "Goa", "31": "Lakshadweep", "32": "Kerala",
  "33": "Tamil Nadu", "34": "Puducherry", "35": "Andaman and Nicobar Islands", "36": "Telangana",
  "37": "Andhra Pradesh", "38": "Ladakh", "97": "Other Territory",
};

export const LEGAL_WORDS = new Set([
  "PRIVATE", "LIMITED", "PVT", "LTD", "LLP", "M/S", "MS", "INDIA", "SERVICES",
  "SERVICE", "COMPANY", "CO", "THE", "CORPORATION", "CORP", "N", "A",
]);

export function round2(v) {
  // Match Python round-half-even closely enough for accounting sums.
  return Math.round((v + Number.EPSILON) * 100) / 100;
}

export function asText(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return fmtDate(value);
  return String(value).trim();
}

export function asFloat(value) {
  if (value === null || value === undefined || value === "") return 0.0;
  if (typeof value === "string" && value.trim().endsWith("%")) return 0.0;
  const result = typeof value === "number" ? value : parseFloat(String(value).replace(/,/g, ""));
  if (Number.isNaN(result) || !Number.isFinite(result)) return 0.0;
  return result;
}

export function normDoc(value) {
  return asText(value).toUpperCase().replace(/[^A-Z0-9]/g, "");
}

export function cleanName(value) {
  const text = asText(value).toUpperCase().replace(/[^A-Z0-9 ]/g, " ");
  return text
    .split(/\s+/)
    .filter((p) => p && !LEGAL_WORDS.has(p))
    .join(" ");
}

// difflib.SequenceMatcher.ratio() port (Ratcliff-Obershelp, no junk heuristic)
function sequenceRatio(a, b) {
  if (!a.length && !b.length) return 1.0;
  if (!a.length || !b.length) return 0.0;
  // b2j map
  const b2j = new Map();
  for (let i = 0; i < b.length; i++) {
    const ch = b[i];
    if (!b2j.has(ch)) b2j.set(ch, []);
    b2j.get(ch).push(i);
  }
  function findLongestMatch(alo, ahi, blo, bhi) {
    let besti = alo, bestj = blo, bestsize = 0;
    let j2len = new Map();
    for (let i = alo; i < ahi; i++) {
      const newj2len = new Map();
      const indices = b2j.get(a[i]) || [];
      for (const j of indices) {
        if (j < blo) continue;
        if (j >= bhi) break;
        const k = (j2len.get(j - 1) || 0) + 1;
        newj2len.set(j, k);
        if (k > bestsize) {
          besti = i - k + 1;
          bestj = j - k + 1;
          bestsize = k;
        }
      }
      j2len = newj2len;
    }
    return [besti, bestj, bestsize];
  }
  let matches = 0;
  const queue = [[0, a.length, 0, b.length]];
  while (queue.length) {
    const [alo, ahi, blo, bhi] = queue.pop();
    const [i, j, k] = findLongestMatch(alo, ahi, blo, bhi);
    if (k) {
      matches += k;
      queue.push([alo, i, blo, j]);
      queue.push([i + k, ahi, j + k, bhi]);
    }
  }
  return (2.0 * matches) / (a.length + b.length);
}

export function nameScore(left, right) {
  const lval = cleanName(left);
  const rval = cleanName(right);
  if (!lval || !rval) return 0.0;
  if (lval === rval) return 1.0;
  const cl = lval.replace(/ /g, "");
  const cr = rval.replace(/ /g, "");
  if (cl && cr && (cr.includes(cl) || cl.includes(cr))) return 0.97;
  if (
    (lval.includes("HSBC") && rval.includes("HONGKONG SHANGHAI BANKING")) ||
    (lval.includes("HONGKONG SHANGHAI BANKING") && rval.includes("HSBC"))
  ) {
    return 0.97;
  }
  if (rval.includes(lval) || lval.includes(rval)) return 0.96;
  return sequenceRatio(lval, rval);
}

const MONTHS = { JAN: 0, FEB: 1, MAR: 2, APR: 3, MAY: 4, JUN: 5, JUL: 6, AUG: 7, SEP: 8, OCT: 9, NOV: 10, DEC: 11 };

export function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const text = asText(value);
  if (!text) return null;
  let m;
  // %d-%m-%Y and %d/%m/%Y
  if ((m = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/))) {
    const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    if (d.getDate() === Number(m[1]) && d.getMonth() === Number(m[2]) - 1) return d;
    return null;
  }
  // %Y-%m-%d
  if ((m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    if (d.getDate() === Number(m[3])) return d;
    return null;
  }
  // %d-%b-%y and %d-%b-%Y
  if ((m = text.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2}|\d{4})$/))) {
    const mon = MONTHS[m[2].toUpperCase()];
    if (mon === undefined) return null;
    let year = Number(m[3]);
    if (m[3].length === 2) year = year < 69 ? 2000 + year : 1900 + year; // strptime %y pivot
    const d = new Date(year, mon, Number(m[1]));
    if (d.getDate() === Number(m[1])) return d;
    return null;
  }
  return null;
}

export function fmtDate(value) {
  if (!(value instanceof Date) || Number.isNaN(value.getTime())) return "";
  const dd = String(value.getDate()).padStart(2, "0");
  const mm = String(value.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}/${value.getFullYear()}`;
}

export function taxTotal(...components) {
  return round2(components.reduce((s, c) => s + asFloat(c), 0));
}

export function amountDiff(left, right) {
  return round2(asFloat(left) - asFloat(right));
}
