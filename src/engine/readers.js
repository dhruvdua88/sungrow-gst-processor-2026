// Excel reading via SheetJS — mirrors read_sheet_rows / read_gl_rows
import * as XLSX from "xlsx";
import { asText, TARGET_COMPANY_ALIASES } from "./helpers.js";

export function workbookFromArrayBuffer(buffer) {
  return XLSX.read(buffer, { type: "array", cellDates: true });
}

export function readSheetRows(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Worksheet "${sheetName}" not found in workbook.`);
  // header:1 keeps raw 2-D rows so we can zip headers like openpyxl does
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  if (!grid.length) return [];
  const headers = grid[0].map((c) => asText(c));
  const rows = [];
  for (let i = 1; i < grid.length; i++) {
    const row = grid[i];
    if (!row.some((cell) => cell !== null && cell !== "")) continue;
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = idx < row.length ? row[idx] : null;
    });
    rows.push(obj);
  }
  return rows;
}

export function readGlRows(gstWorkbook) {
  return readSheetRows(gstWorkbook, "GL Summary").filter((row) => {
    const code = asText(row["Company Code"]);
    return !code || TARGET_COMPANY_ALIASES.has(code.toUpperCase());
  });
}
