// Modified 2026-11-24 by janlewski for Vault Sheets.
import { CellValueType } from '@univerjs/core'
import { describe, expect, it } from 'vitest'
import { csvToWorkbook, detectDelimiter, sheetToCsv } from '../../src/import-export/csv'

describe('csv mapping', () => {
  it('detects a semicolon delimiter and keeps values and formulas', () => {
    expect(detectDelimiter('a;b;c\n1;2;3')).toBe(';')
    const data = csvToWorkbook('Name;Amount;Total\nRent;12.5;=B2+1\ntrue;false;note', { sheetName: 'Imported' })
    const sheet = Object.values(data.sheets)[0]
    expect(sheet?.name).toBe('Imported')
    expect(sheet?.cellData?.[0]?.[0]?.v).toBe('Name')
    expect(sheet?.cellData?.[1]?.[1]).toMatchObject({ v: 12.5, t: CellValueType.NUMBER })
    expect(sheet?.cellData?.[1]?.[2]?.f).toBe('=B2+1')
    expect(sheet?.cellData?.[2]?.[0]).toMatchObject({ v: true, t: CellValueType.BOOLEAN })

    const csv = sheetToCsv(data)
    const again = csvToWorkbook(csv)
    const round = Object.values(again.sheets)[0]
    expect(round?.cellData?.[1]?.[1]?.v).toBe(12.5)
    expect(round?.cellData?.[1]?.[2]?.f).toBe('=B2+1')
    expect(round?.cellData?.[2]?.[0]?.v).toBe(true)
    expect(round?.cellData?.[0]?.[0]?.v).toBe('Name')
  })
})
