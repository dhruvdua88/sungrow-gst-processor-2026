/** Load both workbooks into engine models. Shared by the Web Worker and Node checks. */
import { COPY_BS_SHEETS, COPY_ENTRY_SHEETS, COPY_PL_SHEETS, COPY_TB_SHEETS, parseCopy } from '../engine/parseCopy'
import { parseSignedTB, SIGNED_SHEET_CANDIDATES } from '../engine/parseSigned'
import type { CopyModel, SignedTB } from '../engine/types'
import { pickSheet, readGrids, sheetNames } from './readWorkbook'

export function loadSigned(buf: ArrayBuffer): SignedTB {
  const names = sheetNames(buf)
  const sheet = pickSheet(names, SIGNED_SHEET_CANDIDATES)
  if (!sheet) throw new Error(`The signed workbook has no "T_B" sheet. Sheets found: ${names.slice(0, 12).join(', ')}…`)
  const grids = readGrids(buf, [sheet])
  return parseSignedTB(grids[sheet])
}

export function loadCopy(buf: ArrayBuffer): CopyModel {
  const names = sheetNames(buf)
  const tb = pickSheet(names, COPY_TB_SHEETS)
  if (!tb) throw new Error(`The Ind AS workbook has no "TB" sheet. Sheets found: ${names.slice(0, 12).join(', ')}…`)
  const bs = pickSheet(names, COPY_BS_SHEETS)
  const pl = pickSheet(names, COPY_PL_SHEETS)
  const en = pickSheet(names, COPY_ENTRY_SHEETS)
  const grids = readGrids(buf, [tb, bs, pl, en].filter((x): x is string => !!x))
  return parseCopy({ tb: grids[tb], bs: bs ? grids[bs] : undefined, pl: pl ? grids[pl] : undefined, entries: en ? grids[en] : undefined })
}
