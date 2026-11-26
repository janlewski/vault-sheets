// Modified 2026-11-24 by janlewski for Vault Sheets.

import type { ICellData, IWorkbookData } from '@univerjs/core'
import { CellValueType } from '@univerjs/core'
import Papa from 'papaparse'
import { randomString } from '../utils/uuid'
import { workbookShell, worksheetShell } from './snapshot'

export function detectDelimiter(text: string): ',' | ';' {
  const firstLine = text.split(/\r?\n/, 1)[0] ?? ''
  const commas = firstLine.split(',').length - 1
  const semis = firstLine.split(';').length - 1
  return semis > commas ? ';' : ','
}

function cellFromCsv(raw: string): ICellData | null {
  if (raw === '')
    return null
  if (raw.startsWith('='))
    return { f: raw }
  if (/^(true|false)$/i.test(raw)) {
    return {
      v: raw.toLowerCase() === 'true',
      t: CellValueType.BOOLEAN,
    }
  }
  if (/^-?\d+(?:\.\d+)?(?:e[+-]?\d+)?$/i.test(raw)) {
    return {
      v: Number(raw),
      t: CellValueType.NUMBER,
    }
  }
  return {
    v: raw,
    t: CellValueType.STRING,
  }
}

export function csvToWorkbook(text: string, options: { id?: string, name?: string, sheetName?: string } = {}): IWorkbookData {
  const delimiter = detectDelimiter(text)
  const parsed = Papa.parse<string[]>(text, {
    delimiter,
    skipEmptyLines: false,
  })
  const rows = parsed.data.filter((row, index) => {
    const isLast = index === parsed.data.length - 1
    return !(isLast && row.every(cell => cell === ''))
  })
  const cellData: Record<number, Record<number, ICellData>> = {}
  let maxColumn = 0
  rows.forEach((row, rowIndex) => {
    row.forEach((raw, columnIndex) => {
      const cell = cellFromCsv(raw ?? '')
      if (!cell)
        return
      cellData[rowIndex] ??= {}
      cellData[rowIndex][columnIndex] = cell
      maxColumn = Math.max(maxColumn, columnIndex + 1)
    })
  })
  const sheetId = randomString(6)
  return workbookShell({
    id: options.id,
    name: options.name,
    sheets: [
      worksheetShell(sheetId, options.sheetName || 'Sheet1', {
        cellData,
        rowCount: Math.max(rows.length, 100),
        columnCount: Math.max(maxColumn, 20),
      }),
    ],
  })
}

function cellText(cell: ICellData | undefined): string {
  if (!cell)
    return ''
  if (cell.f)
    return cell.f.startsWith('=') ? cell.f : `=${cell.f}`
  if (cell.v == null)
    return ''
  return String(cell.v)
}

export function sheetToCsv(data: IWorkbookData, sheetId?: string): string {
  const id = sheetId && data.sheets[sheetId] ? sheetId : data.sheetOrder[0]
  const sheet = id ? data.sheets[id] : undefined
  const cellData = sheet?.cellData ?? {}
  let maxRow = -1
  let maxColumn = -1
  for (const [rowKey, row] of Object.entries(cellData)) {
    maxRow = Math.max(maxRow, Number(rowKey))
    for (const colKey of Object.keys(row ?? {}))
      maxColumn = Math.max(maxColumn, Number(colKey))
  }
  if (maxRow < 0 || maxColumn < 0)
    return ''
  const matrix: string[][] = []
  for (let row = 0; row <= maxRow; row++) {
    const line: string[] = []
    for (let column = 0; column <= maxColumn; column++)
      line.push(cellText(cellData[row]?.[column]))
    matrix.push(line)
  }
  return Papa.unparse(matrix, { delimiter: ',' })
}
