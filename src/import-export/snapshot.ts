// Modified 2026-11-24 by janlewski for Vault Sheets.

import type { IWorkbookData, IWorksheetData } from '@univerjs/core'
import { BooleanNumber, LocaleType } from '@univerjs/core'
import { randomString } from '../utils/uuid'

export const DEFAULT_ROW_HEIGHT = 24
export const DEFAULT_COLUMN_WIDTH = 88

export function worksheetShell(id: string, name: string, partial: Partial<IWorksheetData> = {}): Partial<IWorksheetData> & { id: string } {
  return {
    tabColor: '',
    hidden: BooleanNumber.FALSE,
    rowCount: 100,
    columnCount: 20,
    zoomRatio: 1,
    scrollTop: 0,
    scrollLeft: 0,
    defaultColumnWidth: DEFAULT_COLUMN_WIDTH,
    defaultRowHeight: DEFAULT_ROW_HEIGHT,
    mergeData: [],
    cellData: {},
    rowData: {},
    columnData: {},
    showGridlines: BooleanNumber.TRUE,
    rowHeader: { width: 46 },
    columnHeader: { height: 20 },
    freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 },
    ...partial,
    id,
    name,
  }
}

export function workbookShell(options: {
  id?: string
  name?: string
  sheets: Array<Partial<IWorksheetData> & { id: string }>
}): IWorkbookData {
  const sheets: IWorkbookData['sheets'] = {}
  const sheetOrder: string[] = []
  for (const sheet of options.sheets) {
    sheets[sheet.id] = sheet
    sheetOrder.push(sheet.id)
  }
  return {
    id: options.id || randomString(8),
    name: options.name || 'workbook',
    appVersion: '0.1.0',
    locale: LocaleType.EN_US,
    styles: {},
    sheetOrder,
    sheets,
  }
}
