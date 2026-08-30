/**
 * Golden parity tests pinning word_read's extraction contract (v0.3.0 stream
 * B). The expected strings below are the OBSERVED output of mammoth 1.11.0 on
 * 2026-08-30 — captured while mammoth was still the extractor, then frozen:
 *
 *   - every paragraph, including the last and empty ones, ends with `\n\n`;
 *   - `w:tab` renders as `\t`; `w:br` (and `w:cr`) are dropped entirely;
 *   - `w:noBreakHyphen` → `\u2011`, `w:softHyphen` → `\u00AD`;
 *   - hyperlink text survives, the URL does not;
 *   - table cells contribute their paragraphs with no cell separators;
 *   - headers/footers/footnotes are not part of the body text;
 *   - XML entities in `<w:t>` are decoded.
 *
 * word_read must keep producing exactly these bytes.
 */

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, test } from 'vitest'
import JSZip from 'jszip'
import { mountTools, run } from './harness.ts'

/** document.xml exercising tabs, breaks, both hyphens, hyperlinks, empty paragraphs, entities, and a table. */
const HAND_BUILT_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>alpha</w:t><w:tab/><w:t>beta</w:t></w:r></w:p>
    <w:p><w:r><w:t>line</w:t><w:br/><w:t>broken</w:t></w:r></w:p>
    <w:p><w:r><w:t>non</w:t><w:noBreakHyphen/><w:t>break</w:t></w:r><w:r><w:t> soft</w:t><w:softHyphen/><w:t>hy</w:t></w:r></w:p>
    <w:p><w:hyperlink r:id="rId1"><w:r><w:t>linked text</w:t></w:r></w:hyperlink></w:p>
    <w:p/>
    <w:p/>
    <w:p><w:r><w:t xml:space="preserve">a &amp; b &lt; c &#65; &#x42; &quot;q&quot; &apos;s&apos;</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>c00</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>c01</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>c10</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>c11</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:t>tail</w:t></w:r></w:p>
  </w:body>
</w:document>`

/** mammoth 1.11.0 on the `word_create` fixture (title, paragraphs, blank paragraph, bullets, 2x3 table). */
const EXPECTED_WORD_CREATE_OUTPUT = 'Golden Title\n\nFirst body\n\n\n\nSecond body\n\nBullet one\n\nBullet two\n\nH1\n\nH2\n\na\n\nb\n\nc\n\nd\n\n'

/** mammoth 1.11.0 on {@link HAND_BUILT_DOCUMENT_XML}. */
const EXPECTED_HAND_BUILT_OUTPUT = 'alpha\tbeta\n\nlinebroken\n\nnon\u2011break soft\u00ADhy\n\nlinked text\n\n\n\n\n\na & b < c A B "q" \'s\'\n\nc00\n\nc01\n\nc10\n\nc11\n\ntail\n\n'

async function buildHandBuiltDocx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types/>')
  zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships/>')
  zip.file('word/document.xml', HAND_BUILT_DOCUMENT_XML)
  zip.file('word/header1.xml', '<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>HEADER MUST NOT LEAK</w:t></w:r></w:p></w:hdr>')
  zip.file('word/footnotes.xml', '<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>FOOTNOTE MUST NOT LEAK</w:t></w:r></w:p></w:footnote></w:footnotes>')
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

describe('word_read golden parity', () => {
  test('word_create fixture extracts the mammoth-pinned text', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-parity-a-'))
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
      expect(read.totalChars).toBe(EXPECTED_WORD_CREATE_OUTPUT.length)
      expect(read.truncated).toBe(false)
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  test('hand-built fixture extracts the mammoth-pinned text', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-parity-b-'))
    try {
      const buffer = await buildHandBuiltDocx()
      await writeFile(join(root, 'hand.docx'), buffer)

      const tools = mountTools()
      const read = await run(tools, 'word_read', { path: 'hand.docx' }, root) as any
      expect(read.text).toBe(EXPECTED_HAND_BUILT_OUTPUT)
      expect(read.totalChars).toBe(EXPECTED_HAND_BUILT_OUTPUT.length)
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  test('docx without word/document.xml fails with a clear error', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-parity-c-'))
    try {
      const zip = new JSZip()
      zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types/>')
      const buffer = await zip.generateAsync({ type: 'nodebuffer' })
      await writeFile(join(root, 'empty.docx'), buffer)

      const tools = mountTools()
      await expect(run(tools, 'word_read', { path: 'empty.docx' }, root)).rejects.toThrow(/main document|word\/document\.xml/)
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })
})
