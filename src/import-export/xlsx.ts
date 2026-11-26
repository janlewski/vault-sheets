// Modified 2026-11-24 by janlewski for Vault Sheets.
// Clean-room mapper between xlsx (exceljs) and Univer IWorkbookData.

import type { IBorderData, ICellData, IColumnData, IRange, IRowData, IStyleData, IWorkbookData, IWorksheetData } from '@univerjs/core'
import { BooleanNumber, BorderStyleTypes, CellValueType, HorizontalAlign, VerticalAlign, WrapStrategy } from '@univerjs/core'
import ExcelJS from 'exceljs'
import { randomString } from '../utils/uuid'
import { workbookShell, worksheetShell } from './snapshot'

const WIDTH_PX_PER_CHAR = 8
const PT_TO_PX = 96 / 72

const BORDER_FROM_EXCEL: Record<string, BorderStyleTypes> = {
  thin: BorderStyleTypes.THIN,
  hair: BorderStyleTypes.HAIR,
  dotted: BorderStyleTypes.DOTTED,
  dashed: BorderStyleTypes.DASHED,
  dashDot: BorderStyleTypes.DASH_DOT,
  dashDotDot: BorderStyleTypes.DASH_DOT_DOT,
  double: BorderStyleTypes.DOUBLE,
  medium: BorderStyleTypes.MEDIUM,
  mediumDashed: BorderStyleTypes.MEDIUM_DASHED,
  mediumDashDot: BorderStyleTypes.MEDIUM_DASH_DOT,
  mediumDashDotDot: BorderStyleTypes.MEDIUM_DASH_DOT_DOT,
  slantDashDot: BorderStyleTypes.SLANT_DASH_DOT,
  thick: BorderStyleTypes.THICK,
}

const BORDER_TO_EXCEL: Record<number, ExcelJS.BorderStyle> = {
  [BorderStyleTypes.THIN]: 'thin',
  [BorderStyleTypes.HAIR]: 'hair',
  [BorderStyleTypes.DOTTED]: 'dotted',
  [BorderStyleTypes.DASHED]: 'dashed',
  [BorderStyleTypes.DASH_DOT]: 'dashDot',
  [BorderStyleTypes.DASH_DOT_DOT]: 'dashDotDot',
  [BorderStyleTypes.DOUBLE]: 'double',
  [BorderStyleTypes.MEDIUM]: 'medium',
  [BorderStyleTypes.MEDIUM_DASHED]: 'mediumDashed',
  [BorderStyleTypes.MEDIUM_DASH_DOT]: 'mediumDashDot',
  [BorderStyleTypes.MEDIUM_DASH_DOT_DOT]: 'mediumDashDotDot',
  [BorderStyleTypes.SLANT_DASH_DOT]: 'slantDashDot',
  [BorderStyleTypes.THICK]: 'thick',
}

function argbToRgb(argb?: string): string | undefined {
  if (!argb)
    return undefined
  const hex = argb.replace('#', '').toUpperCase()
  if (hex.length === 8)
    return `#${hex.slice(2)}`
  if (hex.length === 6)
    return `#${hex}`
  return undefined
}

function rgbToArgb(rgb?: string | null | void): string | undefined {
  if (!rgb)
    return undefined
  const hex = rgb.replace('#', '').toUpperCase()
  if (hex.length === 6)
    return `FF${hex}`
  if (hex.length === 8)
    return hex
  return undefined
}

function columnIndex(label: string): number {
  let index = 0
  for (const char of label.toUpperCase()) {
    const code = char.charCodeAt(0)
    if (code < 65 || code > 90)
      return -1
    index = index * 26 + (code - 64)
  }
  return index - 1
}

function parseMerge(ref: string): IRange | null {
  const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(ref)
  if (!match)
    return null
  const startColumn = columnIndex(match[1])
  const startRow = Number(match[2]) - 1
  const endColumn = columnIndex(match[3])
  const endRow = Number(match[4]) - 1
  if (startColumn < 0 || endColumn < 0)
    return null
  return { startRow, endRow, startColumn, endColumn }
}

function formulaText(value: object): string | null {
  if ('formula' in value && typeof value.formula === 'string' && value.formula)
    return value.formula
  if ('sharedFormula' in value && typeof value.sharedFormula === 'string' && /[+\-*/():]/.test(value.sharedFormula))
    return value.sharedFormula
  return null
}

function cachedValue(result: unknown): Pick<ICellData, 'v' | 't'> {
  if (typeof result === 'number' && Number.isFinite(result))
    return { v: result, t: CellValueType.NUMBER }
  if (typeof result === 'boolean')
    return { v: result, t: CellValueType.BOOLEAN }
  if (typeof result === 'string')
    return { v: result, t: CellValueType.STRING }
  if (result instanceof Date)
    return { v: dateToSerial(result), t: CellValueType.NUMBER }
  return {}
}

function dateToSerial(date: Date): number {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds())
  const epoch = Date.UTC(1899, 11, 30)
  return (utc - epoch) / 86400000
}

function readStyle(cell: ExcelJS.Cell): IStyleData | null {
  const style: IStyleData = {}
  const font = cell.font
  if (font?.name)
    style.ff = font.name
  if (typeof font?.size === 'number')
    style.fs = font.size
  if (font?.bold)
    style.bl = BooleanNumber.TRUE
  if (font?.italic)
    style.it = BooleanNumber.TRUE
  if (font?.underline && font.underline !== 'none')
    style.ul = { s: BooleanNumber.TRUE }
  const fontColor = argbToRgb(font?.color?.argb)
  if (fontColor)
    style.cl = { rgb: fontColor }

  const fill = cell.fill
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    const background = argbToRgb(fill.fgColor?.argb)
    if (background)
      style.bg = { rgb: background }
  }

  const border = readBorder(cell.border)
  if (border)
    style.bd = border

  const alignment = cell.alignment
  if (alignment?.horizontal === 'left')
    style.ht = HorizontalAlign.LEFT
  else if (alignment?.horizontal === 'center')
    style.ht = HorizontalAlign.CENTER
  else if (alignment?.horizontal === 'right')
    style.ht = HorizontalAlign.RIGHT
  else if (alignment?.horizontal === 'justify')
    style.ht = HorizontalAlign.JUSTIFIED
  if (alignment?.vertical === 'top')
    style.vt = VerticalAlign.TOP
  else if (alignment?.vertical === 'middle')
    style.vt = VerticalAlign.MIDDLE
  else if (alignment?.vertical === 'bottom')
    style.vt = VerticalAlign.BOTTOM
  if (alignment?.wrapText)
    style.tb = WrapStrategy.WRAP

  if (cell.numFmt && cell.numFmt !== 'General')
    style.n = { pattern: cell.numFmt }

  return Object.keys(style).length > 0 ? style : null
}

function readBorder(border: Partial<ExcelJS.Borders> | undefined): IBorderData | null {
  if (!border)
    return null
  const data: IBorderData = {}
  for (const [excelKey, univerKey] of [['top', 't'], ['right', 'r'], ['bottom', 'b'], ['left', 'l']] as const) {
    const side = border[excelKey]
    if (!side?.style)
      continue
    const mapped = BORDER_FROM_EXCEL[side.style]
    if (mapped == null)
      continue
    data[univerKey] = {
      s: mapped,
      cl: { rgb: argbToRgb(side.color?.argb) ?? '#000000' },
    }
  }
  return Object.keys(data).length > 0 ? data : null
}

function readCell(cell: ExcelJS.Cell, styleId: string | undefined): ICellData | null {
  const value = cell.value
  if (value == null || value === '')
    return styleId ? { s: styleId } : null

  const data: ICellData = {}
  if (styleId)
    data.s = styleId

  if (typeof value === 'number') {
    data.v = value
    data.t = CellValueType.NUMBER
  }
  else if (typeof value === 'string') {
    data.v = value
    data.t = CellValueType.STRING
  }
  else if (typeof value === 'boolean') {
    data.v = value
    data.t = CellValueType.BOOLEAN
  }
  else if (value instanceof Date) {
    data.v = dateToSerial(value)
    data.t = CellValueType.NUMBER
    if (!cell.numFmt)
      data.s = styleId
  }
  else if (typeof value === 'object') {
    const formula = formulaText(value)
    if (formula) {
      data.f = formula.startsWith('=') ? formula : `=${formula}`
      Object.assign(data, cachedValue('result' in value ? value.result : undefined))
    }
    else if ('richText' in value && Array.isArray(value.richText)) {
      data.v = value.richText.map(part => part.text ?? '').join('')
      data.t = CellValueType.STRING
    }
    else if ('text' in value && typeof value.text === 'string') {
      data.v = value.text
      data.t = CellValueType.STRING
    }
    else if ('error' in value) {
      data.v = String(value.error)
      data.t = CellValueType.FORCE_STRING
    }
    else {
      return styleId ? { s: styleId } : null
    }
  }
  else {
    return styleId ? { s: styleId } : null
  }

  if (data.v == null && !data.f && !data.s)
    return null
  return data
}

class StyleTable {
  styles: Record<string, IStyleData> = {}
  private ids = new Map<string, string>()

  idFor(style: IStyleData | null): string | undefined {
    if (!style)
      return undefined
    const key = JSON.stringify(style)
    const existing = this.ids.get(key)
    if (existing)
      return existing
    const id = `s${this.ids.size + 1}`
    this.ids.set(key, id)
    this.styles[id] = style
    return id
  }
}

function safeSheetName(name: string, used: Set<string>): string {
  const cleaned = (name || 'Sheet').replace(/[\\/*?:[\]]/g, ' ').trim() || 'Sheet'
  let candidate = cleaned.slice(0, 31)
  let n = 2
  while (used.has(candidate)) {
    const suffix = ` ${n++}`
    candidate = cleaned.slice(0, 31 - suffix.length) + suffix
  }
  used.add(candidate)
  return candidate
}

export async function xlsxToWorkbook(buffer: ArrayBuffer, options: { id?: string, name?: string } = {}): Promise<IWorkbookData> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(buffer as unknown as ExcelJS.Buffer)
  const styles = new StyleTable()
  const sheets: Array<Partial<IWorksheetData> & { id: string }> = []

  workbook.eachSheet((worksheet) => {
    const cellData: Record<number, Record<number, ICellData>> = {}
    const rowData: Record<number, Partial<IRowData>> = {}
    const columnData: Record<number, Partial<IColumnData>> = {}
    let maxRow = 0
    let maxColumn = 0

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const rowIndex = rowNumber - 1
      const rowInfo: Partial<IRowData> = {}
      if (typeof row.height === 'number')
        rowInfo.h = Math.round(row.height * PT_TO_PX)
      if (row.hidden)
        rowInfo.hd = BooleanNumber.TRUE
      if (Object.keys(rowInfo).length > 0)
        rowData[rowIndex] = rowInfo

      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const style = readStyle(cell)
        const styleId = styles.idFor(style)
        const data = readCell(cell, styleId)
        if (!data)
          return
        const columnIndex = colNumber - 1
        cellData[rowIndex] ??= {}
        cellData[rowIndex][columnIndex] = data
        maxRow = Math.max(maxRow, rowNumber)
        maxColumn = Math.max(maxColumn, colNumber)
      })
      maxRow = Math.max(maxRow, row.hidden ? rowNumber : 0)
    })

    for (let index = 1; index <= worksheet.columnCount; index++) {
      const column = worksheet.getColumn(index)
      const info: Partial<IColumnData> = {}
      if (typeof column.width === 'number')
        info.w = Math.round(column.width * WIDTH_PX_PER_CHAR)
      if (column.hidden)
        info.hd = BooleanNumber.TRUE
      if (Object.keys(info).length > 0)
        columnData[index - 1] = info
    }

    const mergeData = (worksheet.model.merges ?? [])
      .map(ref => parseMerge(ref))
      .filter((range): range is IRange => range != null)

    sheets.push(worksheetShell(randomString(6), worksheet.name || `Sheet${sheets.length + 1}`, {
      cellData,
      rowData,
      columnData,
      mergeData,
      rowCount: Math.max(maxRow, worksheet.rowCount, 100),
      columnCount: Math.max(maxColumn, worksheet.columnCount, 20),
    }))
  })

  if (sheets.length === 0)
    sheets.push(worksheetShell(randomString(6), 'Sheet1'))

  const data = workbookShell({ id: options.id, name: options.name ?? workbook.title, sheets })
  data.styles = styles.styles
  return data
}

function resolveStyle(cellStyle: ICellData['s'], styles: IWorkbookData['styles']): IStyleData | null {
  if (!cellStyle)
    return null
  if (typeof cellStyle === 'string')
    return styles?.[cellStyle] ?? null
  return cellStyle
}

function applyStyle(cell: ExcelJS.Cell, style: IStyleData | null): void {
  if (!style)
    return
  const font: Partial<ExcelJS.Font> = {}
  if (style.ff)
    font.name = style.ff
  if (typeof style.fs === 'number')
    font.size = style.fs
  if (style.bl === BooleanNumber.TRUE)
    font.bold = true
  if (style.it === BooleanNumber.TRUE)
    font.italic = true
  if (style.ul?.s === BooleanNumber.TRUE)
    font.underline = true
  const fontColor = rgbToArgb(style.cl?.rgb)
  if (fontColor)
    font.color = { argb: fontColor }
  if (Object.keys(font).length > 0)
    cell.font = font

  const background = rgbToArgb(style.bg?.rgb)
  if (background) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: background },
    }
  }

  if (style.bd) {
    const border: Partial<ExcelJS.Borders> = {}
    for (const [univerKey, excelKey] of [['t', 'top'], ['r', 'right'], ['b', 'bottom'], ['l', 'left']] as const) {
      const side = style.bd[univerKey]
      if (!side)
        continue
      const mapped = BORDER_TO_EXCEL[side.s]
      if (!mapped)
        continue
      border[excelKey] = {
        style: mapped,
        color: { argb: rgbToArgb(side.cl?.rgb) ?? 'FF000000' },
      }
    }
    if (Object.keys(border).length > 0)
      cell.border = border
  }

  const alignment: Partial<ExcelJS.Alignment> = {}
  if (style.ht === HorizontalAlign.LEFT)
    alignment.horizontal = 'left'
  else if (style.ht === HorizontalAlign.CENTER)
    alignment.horizontal = 'center'
  else if (style.ht === HorizontalAlign.RIGHT)
    alignment.horizontal = 'right'
  else if (style.ht === HorizontalAlign.JUSTIFIED)
    alignment.horizontal = 'justify'
  if (style.vt === VerticalAlign.TOP)
    alignment.vertical = 'top'
  else if (style.vt === VerticalAlign.MIDDLE)
    alignment.vertical = 'middle'
  else if (style.vt === VerticalAlign.BOTTOM)
    alignment.vertical = 'bottom'
  if (style.tb === WrapStrategy.WRAP)
    alignment.wrapText = true
  if (Object.keys(alignment).length > 0)
    cell.alignment = alignment

  if (style.n?.pattern)
    cell.numFmt = style.n.pattern
}

export async function workbookToXlsx(data: IWorkbookData): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook()
  const usedNames = new Set<string>()
  const order = data.sheetOrder?.length ? data.sheetOrder : Object.keys(data.sheets ?? {})

  for (const sheetId of order) {
    const sheet = data.sheets?.[sheetId]
    if (!sheet)
      continue
    const worksheet = workbook.addWorksheet(safeSheetName(sheet.name || 'Sheet', usedNames))
    const cellData = (sheet.cellData ?? {}) as Record<string, Record<string, ICellData | null | undefined> | undefined>
    for (const [rowKey, row] of Object.entries(cellData)) {
      for (const [colKey, cell] of Object.entries(row ?? {})) {
        if (!cell)
          continue
        const excelCell = worksheet.getCell(Number(rowKey) + 1, Number(colKey) + 1)
        if (cell.f) {
          excelCell.value = {
            formula: cell.f.replace(/^=/, ''),
            result: cell.v ?? undefined,
          }
        }
        else if (typeof cell.v === 'string' || typeof cell.v === 'number' || typeof cell.v === 'boolean') {
          excelCell.value = cell.v
        }
        applyStyle(excelCell, resolveStyle(cell.s, data.styles))
      }
    }

    for (const [index, info] of Object.entries(sheet.rowData ?? {})) {
      const row = worksheet.getRow(Number(index) + 1)
      if (typeof info?.h === 'number')
        row.height = info.h / PT_TO_PX
      if (info?.hd === BooleanNumber.TRUE)
        row.hidden = true
    }
    for (const [index, info] of Object.entries(sheet.columnData ?? {})) {
      const column = worksheet.getColumn(Number(index) + 1)
      if (typeof info?.w === 'number')
        column.width = info.w / WIDTH_PX_PER_CHAR
      if (info?.hd === BooleanNumber.TRUE)
        column.hidden = true
    }
    for (const merge of sheet.mergeData ?? []) {
      if (merge.endRow < merge.startRow || merge.endColumn < merge.startColumn)
        continue
      worksheet.mergeCells(
        merge.startRow + 1,
        merge.startColumn + 1,
        merge.endRow + 1,
        merge.endColumn + 1,
      )
    }
  }

  if (workbook.worksheets.length === 0)
    workbook.addWorksheet('Sheet1')

  const output = await workbook.xlsx.writeBuffer()
  return output as ArrayBuffer
}
