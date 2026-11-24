// Modified 2026-11-24 by janlewski for Vault Sheets.
import type { ICellData, IStyleData } from '@univerjs/core'
import { BooleanNumber, BorderStyleTypes, CellValueType, HorizontalAlign } from '@univerjs/core'
import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'
import { workbookToXlsx, xlsxToWorkbook } from '../../src/import-export/xlsx'

async function sampleXlsx(): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Data')
  const hello = sheet.getCell('A1')
  hello.value = 'Hello'
  hello.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FFFF0000' } }

  const amount = sheet.getCell('B1')
  amount.value = 12.5
  amount.numFmt = '0.00'

  const formula = sheet.getCell('C1')
  formula.value = { formula: 'B1+1', result: 13.5 }

  const merged = sheet.getCell('D1')
  merged.value = 'Merged'
  merged.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }
  merged.alignment = { horizontal: 'center' }
  merged.border = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } },
  }
  sheet.mergeCells('D1:E1')

  sheet.getColumn(1).width = 20
  sheet.getRow(1).height = 30
  const hiddenRow = sheet.getRow(3)
  hiddenRow.hidden = true
  hiddenRow.height = 15
  sheet.getColumn(3).hidden = true

  return workbook.xlsx.writeBuffer() as Promise<ArrayBuffer>
}

function styleOf(data: Awaited<ReturnType<typeof xlsxToWorkbook>>, cell: ICellData | undefined): IStyleData | null {
  if (!cell?.s)
    return null
  if (typeof cell.s === 'string')
    return data.styles[cell.s] ?? null
  return cell.s
}

describe('xlsx mapping', () => {
  it('reads values, a formula, a merge, styles, and sizes', async () => {
    const data = await xlsxToWorkbook(await sampleXlsx(), { name: 'sample' })
    const sheet = Object.values(data.sheets)[0]
    expect(sheet?.name).toBe('Data')

    const a1 = sheet?.cellData?.[0]?.[0]
    expect(a1?.v).toBe('Hello')
    expect(a1?.t).toBe(CellValueType.STRING)
    const a1Style = styleOf(data, a1)
    expect(a1Style?.ff).toBe('Arial')
    expect(a1Style?.fs).toBe(14)
    expect(a1Style?.bl).toBe(BooleanNumber.TRUE)
    expect(a1Style?.cl?.rgb?.toUpperCase()).toBe('#FF0000')

    const b1 = sheet?.cellData?.[0]?.[1]
    expect(b1?.v).toBe(12.5)
    expect(styleOf(data, b1)?.n?.pattern).toBe('0.00')

    const c1 = sheet?.cellData?.[0]?.[2]
    expect(c1?.f).toBe('=B1+1')
    expect(c1?.v).toBe(13.5)

    const d1 = sheet?.cellData?.[0]?.[3]
    expect(d1?.v).toBe('Merged')
    const d1Style = styleOf(data, d1)
    expect(d1Style?.bg?.rgb?.toUpperCase()).toBe('#FFFF00')
    expect(d1Style?.ht).toBe(HorizontalAlign.CENTER)
    expect(d1Style?.bd?.t?.s).toBe(BorderStyleTypes.THIN)

    expect(sheet?.mergeData).toEqual([
      { startRow: 0, endRow: 0, startColumn: 3, endColumn: 4 },
    ])
    expect(sheet?.columnData?.[0]?.w).toBe(160)
    expect(sheet?.rowData?.[0]?.h).toBe(40)
    expect(sheet?.rowData?.[2]?.hd).toBe(BooleanNumber.TRUE)
    expect(sheet?.columnData?.[2]?.hd).toBe(BooleanNumber.TRUE)
  })

  it('keeps those features across an xlsx round trip', async () => {
    const first = await xlsxToWorkbook(await sampleXlsx())
    const second = await xlsxToWorkbook(await workbookToXlsx(first))
    const sheet = Object.values(second.sheets)[0]
    expect(sheet?.cellData?.[0]?.[0]?.v).toBe('Hello')
    expect(styleOf(second, sheet?.cellData?.[0]?.[0])?.bl).toBe(BooleanNumber.TRUE)
    expect(styleOf(second, sheet?.cellData?.[0]?.[0])?.ff).toBe('Arial')
    expect(sheet?.cellData?.[0]?.[1]?.v).toBe(12.5)
    expect(styleOf(second, sheet?.cellData?.[0]?.[1])?.n?.pattern).toBe('0.00')
    expect(sheet?.cellData?.[0]?.[2]?.f?.replace(/^=/, '')).toBe('B1+1')
    expect(sheet?.cellData?.[0]?.[2]?.v).toBe(13.5)
    expect(sheet?.cellData?.[0]?.[3]?.v).toBe('Merged')
    expect(styleOf(second, sheet?.cellData?.[0]?.[3])?.bg?.rgb?.toUpperCase()).toBe('#FFFF00')
    expect(sheet?.mergeData).toEqual([
      { startRow: 0, endRow: 0, startColumn: 3, endColumn: 4 },
    ])
    expect(sheet?.columnData?.[0]?.w).toBe(160)
    expect(sheet?.rowData?.[0]?.h).toBe(40)
    expect(sheet?.rowData?.[2]?.hd).toBe(BooleanNumber.TRUE)
  })
})
