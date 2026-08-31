/**
 * Excel (.xlsx) tools over SheetJS: `excel_create` writes a new workbook,
 * `excel_read` materializes sheets as rows of scalar cells, and
 * `excel_update` replaces/creates whole sheets and/or writes individual cell
 * values into an existing workbook.
 */

import * as XLSX from 'xlsx'
import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import type { ToolRunContext } from '@deepseek-ai/dsh-tools'
import { atomicWriteFile, assertMayCreate, loadZipGuarded, MAX_READ_CELLS, MAX_WRITE_CELLS, readOfficeBuffer, resolveOfficePath } from '../paths.ts'
import { CELL_VALUE_SCHEMA, FILE_RESULT_SCHEMA, ROW_SCHEMA, type CellRow, type CellValue } from './shared.ts'

interface SheetSpec {
  name: string
  rows: CellRow[]
}

interface ExcelCreateArgs {
  path: string
  sheets: SheetSpec[]
  overwrite?: boolean
}

interface ExcelReadArgs {
  path: string
  sheet?: string
  max_rows?: number
}

interface CellUpdate {
  sheet: string
  cell: string
  value: CellValue
}

interface ExcelUpdateArgs {
  path: string
  sheets?: SheetSpec[]
  cell_updates?: CellUpdate[]
}

const SHEET_RESULT_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    name: { type: 'string', required: true },
    rowCount: { type: 'integer', required: true },
    colCount: { type: 'integer', required: true },
  },
} as const

const EXCEL_CREATE_OUTPUT = {
  type: 'object',
  additionalProperties: false,
  properties: {
    ...FILE_RESULT_SCHEMA.properties,
    sheets: {
      type: 'array',
      required: true,
      items: SHEET_RESULT_SCHEMA,
    },
  },
} as const

const READ_SHEET_RESULT_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    name: { type: 'string', required: true },
    rows: {
      type: 'array',
      required: true,
      items: ROW_SCHEMA,
    },
    truncated: { type: 'boolean', required: true },
  },
} as const

const EXCEL_READ_OUTPUT = {
  type: 'object',
  additionalProperties: false,
  properties: {
    path: { type: 'string', required: true },
    sheets: {
      type: 'array',
      required: true,
      items: READ_SHEET_RESULT_SCHEMA,
    },
    sizeBytes: { type: 'integer', required: true },
  },
} as const

const CELL_UPDATE_RESULT_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    sheet: { type: 'string', required: true },
    cell: { type: 'string', required: true },
  },
} as const

const EXCEL_UPDATE_OUTPUT = {
  type: 'object',
  additionalProperties: false,
  properties: {
    ...FILE_RESULT_SCHEMA.properties,
    sheetNames: {
      type: 'array',
      required: true,
      items: { type: 'string' },
    },
    updatedSheets: {
      type: 'array',
      required: true,
      items: { type: 'string' },
    },
    cellUpdates: {
      type: 'array',
      required: true,
      items: CELL_UPDATE_RESULT_SCHEMA,
    },
  },
} as const

function validateSheetSpecs(sheets: SheetSpec[]): void {
  if (sheets.length === 0) throw new Error('sheets must contain at least one sheet')
  const seen = new Set<string>()
  let totalCells = 0
  let totalRows = 0
  for (const sheet of sheets) {
    if (sheet.name.trim() === '') throw new Error('sheet name must be a non-empty string')
    if (seen.has(sheet.name)) throw new Error(`duplicate sheet name "${sheet.name}" in one call`)
    seen.add(sheet.name)
    if (sheet.rows.length > 10_000) throw new Error(`sheet "${sheet.name}" has too many rows (maximum 10000)`)
    totalRows += sheet.rows.length
    for (const row of sheet.rows) {
      totalCells += row.length
      if (totalCells > MAX_WRITE_CELLS) throw new Error(`too many worksheet cells (maximum ${MAX_WRITE_CELLS})`)
    }
  }
  if (totalRows === 0) throw new Error('at least one row is required across the sheets')
}

function aoaToSheet(rows: CellRow[]): XLSX.WorkSheet {
  const aoa = rows.map(row => row.map(cell => cell === null ? '' : cell))
  const worksheet = XLSX.utils.aoa_to_sheet(aoa as unknown as unknown[][])
  materializeFormulas(worksheet)
  return worksheet
}

/**
 * A string cell that starts with '=' becomes a real formula cell, written as
 * `t="e"` with `<f>` and no cached value — the exact shape Excel and
 * LibreOffice use for uncached formulas, and the only shape SheetJS reads
 * back (a bare `<f>` without `t` is dropped on read). Excel computes the
 * value on open; excel_read returns such cells as '=…' strings until a
 * cached value exists.
 */
function formulaCellOf(cell: unknown): XLSX.CellObject | undefined {
  const candidate = cell as { t?: unknown; v?: unknown }
  if (candidate.t === 's' && typeof candidate.v === 'string' && candidate.v.startsWith('=')) {
    return { t: 'e', f: candidate.v.replace(/^=/, '') } as unknown as XLSX.CellObject
  }
  return undefined
}

/** Rewrite every '=…' string cell of a grid WE just wrote into a formula cell. */
function materializeFormulas(worksheet: XLSX.WorkSheet): void {
  for (const [address, cell] of Object.entries(worksheet)) {
    if (address.startsWith('!')) continue
    const formula = formulaCellOf(cell)
    if (formula !== undefined) worksheet[address] = formula
  }
}

function writeWorkbookBuffer(workbook: XLSX.WorkBook): Buffer {
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
  return Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer as ArrayBuffer)
}

function registerExcelCreate(ctx: Context): () => void {
  return ctx.tools.register(defineTool({
    name: 'excel_create',
    description:
      'Create a new .xlsx Excel workbook in the session workspace from structured sheets. '
      + 'Each sheet has a name and an array of rows; each row is an array of scalar cells (string, number, boolean, or null). '
      + 'Use excel_update to change an existing workbook without recreating it.',
    parameters: {
      path: {
        type: 'string',
        required: true,
        description: 'Output path. Relative paths resolve against the session workspace; the extension must be .xlsx.',
      },
      sheets: {
        type: 'array',
        required: true,
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            name: { type: 'string', required: true, description: 'Worksheet name (unique within this call).' },
            rows: {
              type: 'array',
              required: true,
              items: ROW_SCHEMA,
              description: 'Grid rows; the first row is typically a header row. String cells starting with = are written as formulas.',
            },
          },
        },
        description: 'Sheets to write, in tab order.',
      },
      overwrite: {
        type: 'boolean',
        description: 'Replace the file when it already exists. Defaults to false.',
      },
    },
    output: {
      schema: EXCEL_CREATE_OUTPUT,
      render: (_args, value: any) => [{
        type: 'text',
        text: `Created Excel workbook ${value.path} (${value.sizeBytes} bytes; ${value.sheets.length} sheet(s): ${value.sheets.map((sheet: any) => `${sheet.name} ${sheet.rowCount}x${sheet.colCount}`).join(', ')}).`,
      }],
    },
    presentCall: (args) => ({
      card: 'generic',
      title: `Create ${args.path}`,
      kind: 'edit',
      locations: [{ path: args.path }],
    }),
    async execute(args, exec: ToolRunContext) {
      const target = await resolveOfficePath(exec, args.path, ['.xlsx'], false)
      await assertMayCreate(target.absolute, args.overwrite ?? false)
      validateSheetSpecs(args.sheets)
      exec.signal.throwIfAborted()

      const workbook = XLSX.utils.book_new()
      const summaries: Array<{ name: string; rowCount: number; colCount: number }> = []
      for (const spec of args.sheets) {
        const worksheet = aoaToSheet(spec.rows)
        XLSX.utils.book_append_sheet(workbook, worksheet, spec.name)
        summaries.push({ name: spec.name, rowCount: spec.rows.length, colCount: spec.rows.length === 0 ? 0 : Math.max(...spec.rows.map(row => row.length)) })
      }

      const buffer = writeWorkbookBuffer(workbook)
      exec.signal.throwIfAborted()
      const sizeBytes = await atomicWriteFile(target.absolute, buffer)
      return { path: target.display, sizeBytes, sheets: summaries }
    },
  }))
}

/**
 * Materialize one worksheet as rows of scalar values, replicating SheetJS's
 * `sheet_to_json(header:1, raw:false)` output cell-for-cell (formatted text
 * via `w`, booleans as "TRUE"/"FALSE", gaps as null) with two deliberate
 * extensions: a formula cell without a cached value returns its formula as
 * `'=SUM(A1:A1)'` — symmetric with how writes accept `=…` strings — and rows
 * holding only such formulas are kept instead of being dropped as blank.
 */
function worksheetRows(worksheet: XLSX.WorkSheet): CellRow[] {
  const rangeRef = worksheet['!ref']
  if (typeof rangeRef !== 'string') return []
  const range = XLSX.utils.decode_range(rangeRef)
  const rows: CellRow[] = []
  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row: CellRow = []
    let hasValue = false
    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const cell = worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })] as
        | { t?: string; v?: unknown; w?: string; f?: string }
        | undefined
      const value = cellOf(cell)
      if (value !== null) hasValue = true
      row.push(value)
    }
    if (hasValue) rows.push(row)
  }
  return rows
}

function cellOf(cell: { t?: string; v?: unknown; w?: string; f?: string } | undefined): CellValue {
  if (cell === undefined) return null
  if (cell.t === 'e') return typeof cell.f === 'string' ? `=${cell.f}` : null
  if (cell.w !== undefined) return cell.w
  return cell.v !== undefined && cell.v !== null ? String(cell.v) : null
}

function registerExcelRead(ctx: Context): () => void {
  return ctx.tools.register(defineTool({
    name: 'excel_read',
    description:
      'Read one or all sheets of an existing .xlsx workbook and return each sheet as rows of scalar values (formatted strings). '
      + 'Formula cells return their cached value when one exists; formulas without a cached value return the formula as an "=SUM(…)" string. '
      + 'Rows are capped; the per-sheet `truncated` flag reports when more rows were not returned. '
      + 'Pass `sheet` to read a single named sheet.',
    parameters: {
      path: {
        type: 'string',
        required: true,
        description: 'Path to the .xlsx file, relative to the session workspace or absolute inside it.',
      },
      sheet: {
        type: 'string',
        description: 'Read only this worksheet by exact name. Omit to read every sheet.',
      },
      max_rows: {
        type: 'integer',
        description: 'Maximum rows returned per sheet. Defaults to 5000.',
      },
    },
    output: {
      schema: EXCEL_READ_OUTPUT,
      render: (_args, value: any) => [{
        type: 'text',
        text: value.sheets.map((sheet: any) =>
          `${sheet.name} (${sheet.rows.length} row(s)${sheet.truncated ? ', truncated' : ''}):\n`
          + JSON.stringify(sheet.rows),
        ).join('\n\n'),
      }],
    },
    presentCall: (args) => ({
      card: 'generic',
      title: `Read ${args.path}`,
      kind: 'read',
      locations: [{ path: args.path }],
    }),
    async execute(args, exec: ToolRunContext) {
      const target = await resolveOfficePath(exec, args.path, ['.xlsx'], true)
      const { buffer, sizeBytes } = await readOfficeBuffer(target.absolute, exec.signal)
      await loadZipGuarded(buffer, exec.signal)
      const workbook = XLSX.read(buffer, { type: 'buffer', cellFormula: true, cellHTML: false })
      const names = args.sheet === undefined ? workbook.SheetNames : [args.sheet]
      if (args.sheet !== undefined && !workbook.SheetNames.includes(args.sheet)) {
        throw new Error(`sheet "${args.sheet}" not found; available sheets: ${workbook.SheetNames.join(', ')}`)
      }

      const maxRows = Math.min(Math.max(args.max_rows ?? 5000, 1), 10_000)
      const sheets: Array<{ name: string; rows: CellRow[]; truncated: boolean }> = []
      let totalCells = 0
      let budgetExhausted = false

      for (const name of names) {
        const worksheet = workbook.Sheets[name]
        if (worksheet === undefined) continue
        const rawRows = worksheetRows(worksheet)
        const rows: CellRow[] = []
        let truncated = false
        for (const rawRow of rawRows) {
          if (budgetExhausted) break
          totalCells += rawRow.length
          if (totalCells > MAX_READ_CELLS) {
            truncated = true
            budgetExhausted = true
            break
          }
          rows.push(rawRow)
          if (rows.length >= maxRows) {
            truncated = rawRows.length > rows.length
            break
          }
        }
        if (rawRows.length > rows.length) truncated = true
        sheets.push({ name, rows, truncated })
      }

      return { path: target.display, sheets, sizeBytes }
    },
  }))
}

function registerExcelUpdate(ctx: Context): () => void {
  return ctx.tools.register(defineTool({
    name: 'excel_update',
    description:
      'Update an existing .xlsx workbook in place: replace or create whole sheets by name (`sheets`) and/or write individual scalar values into cells (`cell_updates`, e.g. "B2"). '
      + 'The workbook is rewritten by SheetJS, so unsupported features such as charts and macros may be lost; prefer excel_create for new workbooks. '
      + 'Provide at least one sheet or cell update.',
    parameters: {
      path: {
        type: 'string',
        required: true,
        description: 'Path to the existing .xlsx file.',
      },
      sheets: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            name: { type: 'string', required: true, description: 'Worksheet to replace; created when absent.' },
            rows: {
              type: 'array',
              required: true,
              items: ROW_SCHEMA,
              description: 'Replacement grid rows.',
            },
          },
        },
        description: 'Whole-sheet replacements (optional).',
      },
      cell_updates: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            sheet: { type: 'string', required: true, description: 'Worksheet name.' },
            cell: { type: 'string', required: true, description: 'Cell address in A1 notation, e.g. "B2".' },
            value: {
              ...CELL_VALUE_SCHEMA,
              required: true,
              description: 'Scalar value to write into the cell. A string starting with = is written as a formula.',
            },
          },
        },
        description: 'Individual cell writes (optional).',
      },
    },
    output: {
      schema: EXCEL_UPDATE_OUTPUT,
      render: (_args, value: any) => [{
        type: 'text',
        text: `Updated Excel workbook ${value.path} (${value.sizeBytes} bytes). Sheets now: ${value.sheetNames.join(', ')}. `
          + `Replaced/created sheets: ${value.updatedSheets.length === 0 ? '(none)' : value.updatedSheets.join(', ')}. `
          + `Cell writes: ${value.cellUpdates.length}.`,
      }],
    },
    presentCall: (args) => ({
      card: 'generic',
      title: `Update ${args.path}`,
      kind: 'edit',
      locations: [{ path: args.path }],
    }),
    async execute(args, exec: ToolRunContext) {
      const target = await resolveOfficePath(exec, args.path, ['.xlsx'], true)
      if ((args.sheets?.length ?? 0) === 0 && (args.cell_updates?.length ?? 0) === 0) {
        throw new Error('excel_update needs at least one entry in sheets or cell_updates')
      }

      const sheetSpecs = args.sheets ?? []
      if (sheetSpecs.length > 0) validateSheetSpecs(sheetSpecs)

      const { buffer } = await readOfficeBuffer(target.absolute, exec.signal)
      await loadZipGuarded(buffer, exec.signal)
      const workbook = XLSX.read(buffer, { type: 'buffer', cellFormula: true, cellHTML: false })
      const updatedSheets: string[] = []

      for (const spec of sheetSpecs) {
        const replacement = aoaToSheet(spec.rows)
        if (workbook.SheetNames.includes(spec.name)) {
          const index = workbook.SheetNames.indexOf(spec.name)
          workbook.Sheets[spec.name] = replacement
          void index
        } else {
          XLSX.utils.book_append_sheet(workbook, replacement, spec.name)
        }
        updatedSheets.push(spec.name)
      }

      const cellUpdates: Array<{ sheet: string; cell: string }> = []
      for (const update of args.cell_updates ?? []) {
        const worksheet = workbook.Sheets[update.sheet]
        if (worksheet === undefined) throw new Error(`sheet "${update.sheet}" not found for cell update; available sheets: ${workbook.SheetNames.join(', ')}`)
        try {
          XLSX.utils.decode_cell(update.cell)
        } catch {
          throw new Error(`invalid cell address "${update.cell}"; use A1 notation such as "B2"`)
        }
        XLSX.utils.sheet_add_aoa(worksheet, [[update.value === null ? '' : update.value]], { origin: update.cell })
        const formula = formulaCellOf(worksheet[update.cell])
        if (formula !== undefined) worksheet[update.cell] = formula
        cellUpdates.push({ sheet: update.sheet, cell: update.cell })
      }

      exec.signal.throwIfAborted()
      const sizeBytes = await atomicWriteFile(target.absolute, writeWorkbookBuffer(workbook))
      return {
        path: target.display,
        sizeBytes,
        sheetNames: workbook.SheetNames,
        updatedSheets,
        cellUpdates,
      }
    },
  }))
}

export function registerExcelTools(ctx: Context): () => void {
  const disposers = [registerExcelCreate(ctx), registerExcelRead(ctx), registerExcelUpdate(ctx)]
  return () => disposers.forEach(dispose => dispose())
}
