/**
 * word_read markdown-mode tests (0.5.0): the `format: "markdown"` option
 * renders headings, list items, and tables structurally, while the default
 * text mode stays byte-identical to the mammoth-pinned golden output.
 */

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { mountTools, run } from './harness.ts'
import { EXPECTED_WORD_CREATE_OUTPUT } from './word-parity.spec.ts'

/** A docx exercising headings, nested bullets, and a table with a pipe in a cell. */
const MARKDOWN_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Report</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Overview</w:t></w:r></w:p>
    <w:p><w:r><w:t>Plain body</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Deep detail</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>top item</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>nested item</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>a | b</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>d</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>two paragraphs</w:t></w:r></w:p><w:p><w:r><w:t>in one cell</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>short row</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>tail</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`

const EXPECTED_MARKDOWN = [
  '# Report',
  '# Overview',
  'Plain body',
  '### Deep detail',
  '- top item',
  '  - nested item',
  [
    '| a \\| b | c |',
    '| --- | --- |',
    '| d | two paragraphs in one cell |',
    '| short row |  |',
  ].join('\n'),
  'tail',
].join('\n\n')

describe('word_read markdown mode', () => {
  test('renders headings, nested bullets, and tables from a styled document', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-wmd-'))
    try {
      const zip = new JSZip()
      zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types/>')
      zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships/>')
      zip.file('word/document.xml', MARKDOWN_DOCUMENT_XML)
      await writeFile(join(root, 'styled.docx'), await zip.generateAsync({ type: 'nodebuffer' }))

      const tools = mountTools()
      const read = await run(tools, 'word_read', { path: 'styled.docx', format: 'markdown' }, root) as any
      expect(read.text).toBe(EXPECTED_MARKDOWN)
      expect(read.totalChars).toBe(EXPECTED_MARKDOWN.length)
      expect(read.truncated).toBe(false)
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  test('word_create fixture renders as markdown with a table and bullets', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-wmd-create-'))
    try {
      const tools = mountTools()
      await run(tools, 'word_create', {
        path: 'rich.docx',
        title: 'Golden Title',
        paragraphs: ['First body', 'Second body'],
        bullets: ['Bullet one', 'Bullet two'],
        table: { headers: ['H1', 'H2'], rows: [['a', 'b'], ['c', 'd']] },
      }, root)

      const read = await run(tools, 'word_read', { path: 'rich.docx', format: 'markdown' }, root) as any
      expect(read.text).toBe([
        '# Golden Title',
        'First body',
        'Second body',
        '- Bullet one',
        '- Bullet two',
        [
          '| H1 | H2 |',
          '| --- | --- |',
          '| a | b |',
          '| c | d |',
        ].join('\n'),
      ].join('\n\n'))
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  test('default format stays byte-identical to the pinned plain-text output', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-wmd-default-'))
    try {
      const tools = mountTools()
      await run(tools, 'word_create', {
        path: 'rich.docx',
        title: 'Golden Title',
        paragraphs: ['First body', '', 'Second body'],
        bullets: ['Bullet one', 'Bullet two'],
        table: { headers: ['H1', 'H2'], rows: [['a', 'b'], ['c', 'd']] },
      }, root)

      const read = await run(tools, 'word_read', { path: 'rich.docx' }, root) as any
      expect(read.text).toBe(EXPECTED_WORD_CREATE_OUTPUT)
      const explicit = await run(tools, 'word_read', { path: 'rich.docx', format: 'text' }, root) as any
      expect(explicit.text).toBe(EXPECTED_WORD_CREATE_OUTPUT)
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })
})
