/**
 * Word (.docx) tools: `word_create` builds a styled document with the `docx`
 * package; `word_read` extracts plain text with `mammoth`.
 */

import { Document, HeadingLevel, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from 'docx'
import * as mammoth from 'mammoth'
import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import type { ToolRunContext } from '@deepseek-ai/dsh-tools'
import { atomicWriteFile, assertMayCreate, MAX_TEXT_CHARS, readOfficeBuffer, resolveOfficePath } from '../paths.ts'
import { FILE_RESULT_SCHEMA } from './shared.ts'

interface WordCreateArgs {
  path: string
  title?: string
  paragraphs?: string[]
  bullets?: string[]
  table?: { headers: string[]; rows: string[][] }
  overwrite?: boolean
}

interface WordReadArgs {
  path: string
  max_chars?: number
}

const WORD_CREATE_OUTPUT = {
  type: 'object',
  additionalProperties: false,
  properties: {
    ...FILE_RESULT_SCHEMA.properties,
    title: { type: 'string' },
    paragraphCount: { type: 'integer', required: true },
    bulletCount: { type: 'integer', required: true },
    tableRows: { type: 'integer', required: true },
  },
} as const

const WORD_READ_OUTPUT = {
  type: 'object',
  additionalProperties: false,
  properties: {
    path: { type: 'string', required: true },
    text: { type: 'string', required: true },
    totalChars: { type: 'integer', required: true },
    truncated: { type: 'boolean', required: true },
    sizeBytes: { type: 'integer', required: true },
  },
} as const

function wordCreateCounts(args: WordCreateArgs): { paragraphs: number; cells: number } {
  const tableCells = args.table === undefined
    ? 0
    : (args.table.headers.length + args.table.rows.length) * Math.max(1, args.table.headers.length)
  return {
    paragraphs: (args.title === undefined ? 0 : 1)
      + (args.paragraphs?.length ?? 0)
      + (args.bullets?.length ?? 0)
      + (args.table === undefined ? 0 : 1 + args.table.rows.length),
    cells: tableCells,
  }
}

async function buildDocx(args: WordCreateArgs): Promise<Buffer> {
  const children: (Paragraph | Table)[] = []

  if (args.title !== undefined && args.title.trim() !== '') {
    children.push(new Paragraph({ text: args.title, heading: HeadingLevel.TITLE }))
  }

  for (const text of args.paragraphs ?? []) {
    children.push(new Paragraph(text === '' ? { text: '' } : { children: [new TextRun(text)] }))
  }

  for (const item of args.bullets ?? []) {
    children.push(new Paragraph({ text: item, bullet: { level: 0 } }))
  }

  if (args.table !== undefined) {
    const widths = args.table.headers.length
    const cellWidth = widths > 0 ? Math.max(1, Math.floor(100 / widths)) : 100
    const rows = [
      new TableRow({
        children: args.table.headers.map(header => new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: header, bold: true })] })],
          width: { size: cellWidth, type: WidthType.PERCENTAGE },
        })),
      }),
      ...args.table.rows.map(row => new TableRow({
        children: row.map(cell => new TableCell({
          children: [new Paragraph({ text: cell })],
          width: { size: cellWidth, type: WidthType.PERCENTAGE },
        })),
      })),
    ]
    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows,
    }))
  }

  return Packer.toBuffer(new Document({ sections: [{ children }] }))
}

function registerWordCreate(ctx: Context): () => void {
  return ctx.tools.register(defineTool({
    name: 'word_create',
    description:
      'Create a Microsoft Word .docx document inside the session workspace from structured content. '
      + 'Supply paragraphs as plain text, optional bullet points, and one optional table (headers + string rows). '
      + 'The file is written atomically; pass overwrite: true to replace an existing file. '
      + 'Use word_read afterwards to verify the extracted text.',
    parameters: {
      path: {
        type: 'string',
        required: true,
        description: 'Output path. Relative paths resolve against the session workspace; the extension must be .docx.',
      },
      title: {
        type: 'string',
        description: 'Document title rendered as the title heading. Optional.',
      },
      paragraphs: {
        type: 'array',
        items: { type: 'string' },
        description: 'Body paragraphs in document order. Empty strings create blank paragraphs. Optional.',
      },
      bullets: {
        type: 'array',
        items: { type: 'string' },
        description: 'Bullet list items rendered after the paragraphs. Optional.',
      },
      table: {
        type: 'object',
        additionalProperties: false,
        properties: {
          headers: {
            type: 'array',
            items: { type: 'string' },
            required: true,
            description: 'Table column headers (bold).',
          },
          rows: {
            type: 'array',
            items: { type: 'array', items: { type: 'string' } },
            required: true,
            description: 'Table body rows; each row should match the header column count.',
          },
        },
        description: 'One optional table appended after the text content.',
      },
      overwrite: {
        type: 'boolean',
        description: 'Replace the file when it already exists. Defaults to false (existing files are refused).',
      },
    },
    output: {
      schema: WORD_CREATE_OUTPUT,
      render: (_args, value: any) => [{
        type: 'text',
        text: `Created Word document ${value.path} (${value.sizeBytes} bytes; ${value.paragraphCount} paragraphs, ${value.bulletCount} bullets, ${value.tableRows} table body rows).`,
      }],
    },
    presentCall: (args) => ({
      card: 'generic',
      title: `Create ${args.path}`,
      kind: 'edit',
      locations: [{ path: args.path }],
    }),
    async execute(args, exec: ToolRunContext) {
      const target = await resolveOfficePath(exec, args.path, ['.docx'], false)
      await assertMayCreate(target.absolute, args.overwrite ?? false)
      exec.signal.throwIfAborted()

      const { paragraphs: paragraphCount, cells } = wordCreateCounts(args)
      if (paragraphCount > 10_000) throw new Error('too many paragraphs/bullets/table rows (maximum 10000)')
      if (cells > 200_000) throw new Error('too many table cells (maximum 200000)')
      if (args.title === undefined && (args.paragraphs?.length ?? 0) === 0 && (args.bullets?.length ?? 0) === 0 && args.table === undefined) {
        throw new Error('word_create needs at least one of title, paragraphs, bullets, or table')
      }

      const buffer = await buildDocx(args)
      exec.signal.throwIfAborted()
      const sizeBytes = await atomicWriteFile(target.absolute, buffer)
      const result: {
        path: string
        sizeBytes: number
        title?: string
        paragraphCount: number
        bulletCount: number
        tableRows: number
      } = {
        path: target.display,
        sizeBytes,
        paragraphCount: (args.title === undefined || args.title.trim() === '' ? 0 : 1) + (args.paragraphs?.length ?? 0),
        bulletCount: args.bullets?.length ?? 0,
        tableRows: args.table?.rows.length ?? 0,
      }
      if (args.title !== undefined && args.title.trim() !== '') result.title = args.title
      return result
    },
  }))
}

function registerWordRead(ctx: Context): () => void {
  return ctx.tools.register(defineTool({
    name: 'word_read',
    description:
      'Extract plain text from an existing .docx Word document in the session workspace. '
      + 'Returns the full text up to the character limit and a truncated flag when more remains.',
    parameters: {
      path: {
        type: 'string',
        required: true,
        description: 'Path to the .docx file, relative to the session workspace or absolute inside it.',
      },
      max_chars: {
        type: 'integer',
        description: `Maximum characters to return. Defaults to ${MAX_TEXT_CHARS}.`,
      },
    },
    output: {
      schema: WORD_READ_OUTPUT,
      render: (_args, value: any) => [{
        type: 'text',
        text: value.text + (value.truncated ? `\n[text truncated; total ${value.totalChars} characters]` : ''),
      }],
    },
    presentCall: (args) => ({
      card: 'generic',
      title: `Read ${args.path}`,
      kind: 'read',
      locations: [{ path: args.path }],
    }),
    async execute(args, exec: ToolRunContext) {
      const target = await resolveOfficePath(exec, args.path, ['.docx'], true)
      const { buffer, sizeBytes } = await readOfficeBuffer(target.absolute, exec.signal)
      const result = await mammoth.extractRawText({ buffer })
      const totalChars = result.value.length
      const maxChars = Math.min(Math.max(args.max_chars ?? MAX_TEXT_CHARS, 1), MAX_TEXT_CHARS)
      const truncated = totalChars > maxChars
      const text = truncated ? result.value.slice(0, maxChars) : result.value
      return { path: target.display, text, totalChars, truncated, sizeBytes }
    },
  }))
}

export function registerWordTools(ctx: Context): () => void {
  const disposeCreate = registerWordCreate(ctx)
  const disposeRead = registerWordRead(ctx)
  return () => {
    disposeCreate()
    disposeRead()
  }
}
