// Modified 2026-11-24 by janlewski for Vault Sheets.

import type { IAccessor, ICommand, IWorkbookData } from '@univerjs/core'
import { CommandType, ICommandService, Injector } from '@univerjs/core'
import { FUniver } from '@univerjs/core/facade'
import '@univerjs/sheets/facade'
import { csvToWorkbook, sheetToCsv } from './csv'
import { downloadBlob, downloadStem, pickFile } from './file'
import { workbookToXlsx, xlsxToWorkbook } from './xlsx'

export const ImportXlsxStartCommand: ICommand = {
  id: 'vault-sheets.command.import-xlsx-start',
  type: CommandType.COMMAND,
  handler: () => true,
}
export const ImportXlsxFinishCommand: ICommand = {
  id: 'vault-sheets.command.import-xlsx-finish',
  type: CommandType.COMMAND,
  handler: () => true,
}
export const ExportXlsxStartCommand: ICommand = {
  id: 'vault-sheets.command.export-xlsx-start',
  type: CommandType.COMMAND,
  handler: () => true,
}
export const ExportXlsxFinishCommand: ICommand = {
  id: 'vault-sheets.command.export-xlsx-finish',
  type: CommandType.COMMAND,
  handler: () => true,
}
export const ImportCsvStartCommand: ICommand = {
  id: 'vault-sheets.command.import-csv-start',
  type: CommandType.COMMAND,
  handler: () => true,
}
export const ImportCsvFinishCommand: ICommand = {
  id: 'vault-sheets.command.import-csv-finish',
  type: CommandType.COMMAND,
  handler: () => true,
}
export const ExportCsvStartCommand: ICommand = {
  id: 'vault-sheets.command.export-csv-start',
  type: CommandType.COMMAND,
  handler: () => true,
}
export const ExportCsvFinishCommand: ICommand = {
  id: 'vault-sheets.command.export-csv-finish',
  type: CommandType.COMMAND,
  handler: () => true,
}

function replaceActiveWorkbook(accessor: IAccessor, data: IWorkbookData): void {
  const univerAPI = FUniver.newAPI(accessor.get(Injector))
  const current = univerAPI.getActiveWorkbook()
  const name = current?.getName() || data.name
  const oldId = current?.getId()
  const created = univerAPI.createWorkbook({ ...data, name })
  if (oldId && oldId !== created.getId())
    univerAPI.disposeUnit(oldId)
}

async function runImport(
  accessor: IAccessor,
  startId: string,
  finishId: string,
  load: () => Promise<IWorkbookData | null>,
): Promise<boolean> {
  const commandService = accessor.get(ICommandService)
  await commandService.executeCommand(startId)
  try {
    const data = await load()
    if (!data)
      return false
    replaceActiveWorkbook(accessor, data)
    return true
  }
  catch (error) {
    console.error(error)
    return false
  }
  finally {
    await commandService.executeCommand(finishId)
  }
}

export const ImportXlsxCommand: ICommand = {
  id: 'vault-sheets.command.import-xlsx',
  type: CommandType.COMMAND,
  handler: async (accessor) => {
    const picked = pickFile('.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return runImport(accessor, ImportXlsxStartCommand.id, ImportXlsxFinishCommand.id, async () => {
      const file = await picked
      if (!file)
        return null
      return xlsxToWorkbook(await file.arrayBuffer())
    })
  },
}

export const ExportXlsxCommand: ICommand = {
  id: 'vault-sheets.command.export-xlsx',
  type: CommandType.COMMAND,
  handler: async (accessor) => {
    const commandService = accessor.get(ICommandService)
    await commandService.executeCommand(ExportXlsxStartCommand.id)
    try {
      const univerAPI = FUniver.newAPI(accessor.get(Injector))
      const current = univerAPI.getActiveWorkbook()
      const snapshot = current?.save()
      if (!snapshot)
        return false
      const buffer = await workbookToXlsx(snapshot)
      downloadBlob(
        new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        `${downloadStem(snapshot.name)}.xlsx`,
      )
      return true
    }
    catch (error) {
      console.error(error)
      return false
    }
    finally {
      await commandService.executeCommand(ExportXlsxFinishCommand.id)
    }
  },
}

export const ImportCsvCommand: ICommand = {
  id: 'vault-sheets.command.import-csv',
  type: CommandType.COMMAND,
  handler: async (accessor) => {
    const picked = pickFile('.csv,text/csv')
    return runImport(accessor, ImportCsvStartCommand.id, ImportCsvFinishCommand.id, async () => {
      const file = await picked
      if (!file)
        return null
      const text = await file.text()
      const stem = file.name.replace(/\.csv$/i, '') || 'Sheet1'
      return csvToWorkbook(text, { sheetName: stem })
    })
  },
}

export const ExportCsvCommand: ICommand = {
  id: 'vault-sheets.command.export-csv',
  type: CommandType.COMMAND,
  handler: async (accessor) => {
    const commandService = accessor.get(ICommandService)
    await commandService.executeCommand(ExportCsvStartCommand.id)
    try {
      const univerAPI = FUniver.newAPI(accessor.get(Injector))
      const current = univerAPI.getActiveWorkbook()
      const snapshot = current?.save()
      if (!snapshot)
        return false
      const sheetId = current?.getActiveSheet()?.getSheetId()
      const csv = sheetToCsv(snapshot, sheetId)
      downloadBlob(new Blob([csv], { type: 'text/csv;charset=utf-8' }), `${downloadStem(snapshot.name)}.csv`)
      return true
    }
    catch (error) {
      console.error(error)
      return false
    }
    finally {
      await commandService.executeCommand(ExportCsvFinishCommand.id)
    }
  },
}

export const exchangeCommands: ICommand[] = [
  ImportXlsxStartCommand,
  ImportXlsxFinishCommand,
  ExportXlsxStartCommand,
  ExportXlsxFinishCommand,
  ImportCsvStartCommand,
  ImportCsvFinishCommand,
  ExportCsvStartCommand,
  ExportCsvFinishCommand,
  ImportXlsxCommand,
  ExportXlsxCommand,
  ImportCsvCommand,
  ExportCsvCommand,
]
