/**
 * Formula-writing tests (v0.4.0): string cells starting with '=' must land
 * in the workbook as real `<f>` formula cells (Excel computes them on open),
 * while plain strings stay plain strings. Verified at the artifact level
 * because a freshly written formula has no cached value for excel_read to
 * return yet (formula read-back is on the 0.5.0 roadmap).
 */

import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { readFile } from 'node:fs/promises'
import { describe, expect, test } from 'vitest'
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
