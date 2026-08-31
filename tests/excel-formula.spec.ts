/**
 * Formula tests for the Excel tools: writes turn '=…' strings into real `<f>`
 * formula cells (0.4.0), and excel_read returns cached values or, when a
 * formula has no cached value yet, the formula as an '=…' string again
 * (0.5.0). Write-side checks inspect the artifact because a freshly written
 * formula has no cached value to read.
 */

import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { readFile } from 'node:fs/promises'
import { describe, expect, test } from 'vitest'
import * as XLSX from 'xlsx'
import JSZip from 'jszip'
import { mountTools, run } from './harness.ts'

/** The worksheet XML of the first sheet inside a written .xlsx. */
async function firstSheetXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(path))
  const sheet = zip.file('xl/worksheets/sheet1.xml')
  if (sheet === null) throw new Error('no sheet1.xml in artifact')
  return sheet.async('string')
}

describe('excel formula writing', () => {
  test('excel_create writes =-strings as <f> formula cells', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-formula-create-'))
    try {
      const tools = mountTools()
      await run(tools, 'excel_create', {
        path: 'book.xlsx',
        sheets: [{ name: 'S', rows: [['qty', 'price', 'total'], [3, 4, '=B2*C2'], ['plain = not a formula']] }],
      }, root)

      const xml = await firstSheetXml(join(root, 'book.xlsx'))
      expect(xml).toContain('<f>B2*C2</f>')
      // A plain string that merely contains '=' stays a string cell.
      expect(xml).not.toContain('<f>plain')
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  test('excel_update cell_updates turn =-strings into formulas', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-formula-update-'))
    try {
      const tools = mountTools()
      await run(tools, 'excel_create', {
        path: 'book.xlsx',
        sheets: [{ name: 'S', rows: [['a', 'b'], [1, 2]] }],
      }, root)
      await run(tools, 'excel_update', {
        path: 'book.xlsx',
        cell_updates: [
          { sheet: 'S', cell: 'C2', value: '=A2+B2' },
          { sheet: 'S', cell: 'C3', value: 'plain text' },
        ],
      }, root)

      const xml = await firstSheetXml(join(root, 'book.xlsx'))
      expect(xml).toContain('<f>A2+B2</f>')
      expect(xml).not.toContain('<f>plain text</f>')
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  test('whole-sheet replacement through excel_update materializes formulas too', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-formula-replace-'))
    try {
      const tools = mountTools()
      await run(tools, 'excel_create', {
        path: 'book.xlsx',
        sheets: [{ name: 'S', rows: [['old']] }],
      }, root)
      await run(tools, 'excel_update', {
        path: 'book.xlsx',
        sheets: [{ name: 'S', rows: [['=SUM(1,2)']] }],
      }, root)

      const xml = await firstSheetXml(join(root, 'book.xlsx'))
      expect(xml).toContain('<f>SUM(1,2)</f>')
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })
})

describe('excel_read formula read-back (0.5.0)', () => {
  test('formulas without a cached value read back as =… strings', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-formula-read-'))
    try {
      const tools = mountTools()
      await run(tools, 'excel_create', {
        path: 'book.xlsx',
        sheets: [{ name: 'S', rows: [['qty', 'price', 'total'], [3, 4, '=B2*C2']] }],
      }, root)

      const read = await run(tools, 'excel_read', { path: 'book.xlsx', sheet: 'S' }, root) as any
      expect(read.sheets[0].rows[0]).toEqual(['qty', 'price', 'total'])
      expect(read.sheets[0].rows[1]).toEqual(['3', '4', '=B2*C2'])
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  test('formulas with a cached value read back as the cached value', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-formula-cache-'))
    try {
      // Build a workbook whose formula cell carries a cached value, like a
      // file last saved by Excel.
      const worksheet = XLSX.utils.aoa_to_sheet([['result']])
      worksheet.A1 = { t: 'n', v: 3, f: 'SUM(1,2)' } as XLSX.CellObject
      const workbook = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(workbook, worksheet, 'S')
      const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }) as Buffer
      const { writeFile } = await import('node:fs/promises')
      await writeFile(join(root, 'cached.xlsx'), buffer)

      const tools = mountTools()
      const read = await run(tools, 'excel_read', { path: 'cached.xlsx', sheet: 'S' }, root) as any
      expect(read.sheets[0].rows[0]).toEqual(['3'])
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  test('a row holding only uncached formulas is kept, not dropped as blank', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-formula-row-'))
    try {
      const tools = mountTools()
      await run(tools, 'excel_create', {
        path: 'book.xlsx',
        sheets: [{ name: 'S', rows: [['header'], ['=1+1'], ['tail']] }],
      }, root)

      const read = await run(tools, 'excel_read', { path: 'book.xlsx', sheet: 'S' }, root) as any
      expect(read.sheets[0].rows).toEqual([['header'], ['=1+1'], ['tail']])
      expect(read.sheets[0].truncated).toBe(false)
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  test('scalar formatting parity: booleans, empty strings, and gaps survive the manual walk', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-formula-parity-'))
    try {
      const tools = mountTools()
      await run(tools, 'excel_create', {
        path: 'book.xlsx',
        // a null cell writes as an empty string (old behavior, kept); the
        // short second row leaves B2 genuinely absent, which reads as null.
        sheets: [{ name: 'S', rows: [['text', true, false, '', 1.5], ['only-a']] }],
      }, root)

      const read = await run(tools, 'excel_read', { path: 'book.xlsx', sheet: 'S' }, root) as any
      expect(read.sheets[0].rows[0]).toEqual(['text', 'TRUE', 'FALSE', '', '1.5'])
      expect(read.sheets[0].rows[1]).toEqual(['only-a', null, null, null, null])
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })
})
