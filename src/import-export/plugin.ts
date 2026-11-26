// Modified 2026-11-24 by janlewski for Vault Sheets.

import type { IMenuItem } from '@univerjs/ui'
import { ICommandService, Inject, Injector, Plugin, UniverInstanceType } from '@univerjs/core'
import { IMenuManagerService, MenuItemType, RibbonStartGroup } from '@univerjs/ui'
import { t } from '../lang/helpers'
import {
  exchangeCommands,
  ExportCsvCommand,
  ExportXlsxCommand,
  ImportCsvCommand,
  ImportXlsxCommand,
} from './commands'

function menuItem(id: string, title: string): IMenuItem {
  return {
    id,
    title,
    type: MenuItemType.BUTTON,
  }
}

export class VaultSheetsExchangePlugin extends Plugin {
  static override pluginName = 'VAULT_SHEETS_EXCHANGE_PLUGIN'
  static override type = UniverInstanceType.UNIVER_SHEET

  constructor(
    @Inject(Injector) protected override readonly _injector: Injector,
  ) {
    super()
  }

  override onStarting(): void {
    const commandService = this._injector.get(ICommandService)
    const menuManager = this._injector.get(IMenuManagerService)
    for (const command of exchangeCommands)
      commandService.registerCommand(command)

    menuManager.mergeMenu({
      [RibbonStartGroup.OTHERS]: {
        [ImportXlsxCommand.id]: {
          order: 20,
          menuItemFactory: () => menuItem(ImportXlsxCommand.id, t('IMPORT_XLSX')),
        },
        [ImportCsvCommand.id]: {
          order: 21,
          menuItemFactory: () => menuItem(ImportCsvCommand.id, t('IMPORT_CSV')),
        },
        [ExportXlsxCommand.id]: {
          order: 22,
          menuItemFactory: () => menuItem(ExportXlsxCommand.id, t('EXPORT_XLSX')),
        },
        [ExportCsvCommand.id]: {
          order: 23,
          menuItemFactory: () => menuItem(ExportCsvCommand.id, t('EXPORT_CSV')),
        },
      },
    })
  }
}
