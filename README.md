# Vault Sheets

Edit spreadsheets in your vault. Vault Sheets is a free [Univer](https://github.com/dream-num/univer) spreadsheet editor with its own xlsx and csv import and export.

Vault Sheets is an independent project. It is **not** official Sheet Plus and it is **not** affiliated with ljcoder. It is a derivative of [obsidian-sheet-plus](https://github.com/ljcoder2015/obsidian-sheet-plus) (Apache-2.0, commit `f4d3ea7`). That link is attribution only. Documentation for this project is this repository.

## What it does

- Edit `.sheet` files: formulas, formatting, data validation, filters, comments, and the rest of the Univer spreadsheet UI.
- Import and export `.xlsx` and `.csv` in the browser. No separate license key.
- Keep the basic integrations that already ship in Apache-2.0 packages: in-sheet charts, embed a **sheet** in a note, outgoing links, vault images, and save.

## What it does not do

- Calendar, Kanban, and Dashboard views. Older `.sheet` files that still contain those tabs open on a short placeholder. The spreadsheet tab still edits and saves. You cannot create those views.
- MCP, agent, or REST API access.
- Embed **chart** links. A chart embed shows a short notice instead of a chart. Sheet and cell embeds still work.
- It is not listed in the official Community plugin directory. Do not expect it there.

### xlsx and csv scope

Import and export cover workbook sheets (names and order), cell values (text, number, boolean, date), formulas with cached values, merged cells, basic font / fill / border / alignment / wrap, column widths, row heights, hidden rows and columns when they are straightforward, and number formats.

These are skipped without failing the import: conditional formatting, data validation, filters, images, drawings, charts, comments, pivot tables, rich text beyond basic font styling, and sheet protection. Frozen panes are not preserved.

CSV import replaces the open workbook with a single sheet parsed as UTF-8. A leading `;` or `,` in the file selects the delimiter. Formula-looking cells (text starting with `=`) are stored as formulas. CSV export writes the **active** sheet only, values and formulas as text, comma-separated, with no styles.

## Install

Vault Sheets is not in the official Community plugin directory. Install it from a GitHub Release or with [BRAT](https://github.com/TfTHacker/obsidian42-brat).

### Manual

1. Download `main.js`, `manifest.json`, and `styles.css` from the [latest release](https://github.com/janlewski/vault-sheets/releases).
2. Create `<vault>/.obsidian/plugins/vault-sheets/`.
3. Put those three files in that folder. The names must stay exactly `main.js`, `manifest.json`, and `styles.css`.
4. Reload Community plugins and enable **Vault Sheets** under Installed.

### BRAT

1. Install BRAT.
2. Add `janlewski/vault-sheets` as a beta plugin.
3. Select a release. BRAT uses the three files attached to that release.

Minimum app version: 1.13.0.

## Development

```bash
pnpm install
pnpm test
pnpm build
```

`pnpm build` writes `dist/main.js`, `dist/manifest.json`, and `dist/styles.css`.

Pushing a `v*` or `0.*` tag runs the release workflow. It tests, builds, and attaches those three files to the GitHub Release. `versions.json` maps each plugin version to `minAppVersion`.

## License

Apache License 2.0. See [LICENSE](LICENSE), [NOTICE](NOTICE), and [THIRD_PARTY.md](THIRD_PARTY.md).
