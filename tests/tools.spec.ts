import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { describe, expect, test } from 'vitest'
import { Context } from '@deepseek-ai/cordis'
import type { ToolDefinition, ToolRunContext } from '@deepseek-ai/dsh-tools'
import { ToolRuntime, assertSupportedJsonSchema } from '@deepseek-ai/dsh-tools'
import { apply, Config } from '../src/index.ts'
import { resolveOfficePath } from '../src/paths.ts'

interface ToolRegistryLike {
  register(definition: ToolDefinition): () => void
}

interface AgentLike {
  session: { header: { cwd: string } }
}

/** 1x1 red PNG used by the ppt image-embedding tests. */
const TINY_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
)

function execFor(root: string): ToolRunContext {
  return {
    signal: new AbortController().signal,
    agent: { session: { header: { cwd: root } } } as AgentLike,
    callId: 'test-call',
    name: 'test',
    arguments: {},
  } as unknown as ToolRunContext
}

function mountTools(config?: { enablePptTools?: boolean }): Map<string, ToolDefinition> {
  const tools = new Map<string, ToolDefinition>()
  const context = {
    tools: {
      register(definition: ToolDefinition) {
        if (tools.has(definition.name)) throw new Error(`duplicate tool ${definition.name}`)
        tools.set(definition.name, definition)
        return () => tools.delete(definition.name)
      },
    } as ToolRegistryLike,
    effect(setup: () => () => void) {
      return setup()
    },
  } as unknown as Context
  apply(context, config)
  return tools
}

async function run(tools: Map<string, ToolDefinition>, name: string, args: Record<string, unknown>, root: string) {
  const tool = tools.get(name)
  expect(tool, `tool ${name} should be registered`).toBeDefined()
  return tool!.execute(args, execFor(root))
}

describe('tool registration', () => {
  test('registers all eight office tools exactly once', () => {
    const tools = mountTools()
    expect([...tools.keys()].sort()).toEqual([
      'excel_create',
      'excel_read',
      'excel_update',
      'ppt_create',
      'ppt_read',
      'word_create',
      'word_read',
      'word_update',
    ])
    for (const tool of tools.values()) {
      expect(() => assertSupportedJsonSchema(tool.parameters)).not.toThrow()
      expect(() => assertSupportedJsonSchema(tool.output.schema)).not.toThrow()
    }
  })

  test('registers against the real dsh-tools runtime', () => {
    const context = new Context()
    context.provide('systemPrompt', { tools() {}, section() {} })
    new ToolRuntime(context)
    apply(context)
    expect(context.tools.schemas().map(schema => schema.name).sort()).toEqual([
      'excel_create',
      'excel_read',
      'excel_update',
      'ppt_create',
      'ppt_read',
      'word_create',
      'word_read',
      'word_update',
    ])
  })
})

describe('enablePptTools config switch', () => {
  test('schema defaults enable ppt tools; explicit false disables them', () => {
    expect(Config({}).enablePptTools).toBe(true)
    expect(Config({ enablePptTools: false }).enablePptTools).toBe(false)
  })

  test('enablePptTools: false registers only the Word and Excel tools', () => {
    const tools = mountTools({ enablePptTools: false })
    expect([...tools.keys()].sort()).toEqual([
      'excel_create',
      'excel_read',
      'excel_update',
      'word_create',
      'word_read',
      'word_update',
    ])
    for (const tool of tools.values()) {
      expect(() => assertSupportedJsonSchema(tool.parameters)).not.toThrow()
      expect(() => assertSupportedJsonSchema(tool.output.schema)).not.toThrow()
    }
  })

  test('real runtime honors enablePptTools: false', () => {
    const context = new Context()
    context.provide('systemPrompt', { tools() {}, section() {} })
    new ToolRuntime(context)
    apply(context, { enablePptTools: false })
    expect(context.tools.schemas().map(schema => schema.name).sort()).toEqual([
      'excel_create',
      'excel_read',
      'excel_update',
      'word_create',
      'word_read',
      'word_update',
    ])
  })

  test('word and excel tools stay fully functional with ppt tools disabled', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-noppt-'))
    try {
      const tools = mountTools({ enablePptTools: false })
      const doc = await run(tools, 'word_create', { path: 'doc.docx', paragraphs: ['text'] }, root) as any
      expect(doc.sizeBytes).toBeGreaterThan(0)
      const docRead = await run(tools, 'word_read', { path: 'doc.docx' }, root) as any
      expect(docRead.text).toContain('text')

      await run(tools, 'excel_create', {
        path: 'book.xlsx',
        sheets: [{ name: 'S', rows: [['a', 1]] }],
      }, root)
      const book = await run(tools, 'excel_read', { path: 'book.xlsx' }, root) as any
      expect(book.sheets[0].rows[0]).toEqual(['a', '1'])
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })
})

describe('word tools', () => {
  test('create a docx and read its text back', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-word-'))
    try {
      const tools = mountTools()
      const created = await run(tools, 'word_create', {
        path: 'report.docx',
        title: 'Quarterly Report',
        paragraphs: ['Hello world', 'Second paragraph'],
        bullets: ['Alpha', 'Beta'],
        table: { headers: ['Name', 'Qty'], rows: [['Apples', '3'], ['Pears', '5']] },
      }, root) as any

      expect(created.path).toBe('report.docx')
      expect(created.paragraphCount).toBe(3)
      expect(created.bulletCount).toBe(2)
      expect(created.tableRows).toBe(2)
      expect((await stat(join(root, 'report.docx'))).isFile()).toBe(true)

      const read = await run(tools, 'word_read', { path: 'report.docx' }, root) as any
      expect(read.totalChars).toBeGreaterThan(0)
      expect(read.truncated).toBe(false)
      expect(read.text).toContain('Quarterly Report')
      expect(read.text).toContain('Hello world')
      expect(read.text).toContain('Alpha')
      expect(read.text).toContain('Apples')

      const limited = await run(tools, 'word_read', { path: 'report.docx', max_chars: 5 }, root) as any
      expect(limited.text).toHaveLength(5)
      expect(limited.truncated).toBe(true)
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  test('refuses to overwrite an existing docx unless asked', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-word-overwrite-'))
    try {
      const tools = mountTools()
      await run(tools, 'word_create', { path: 'a.docx', paragraphs: ['one'] }, root)
      await expect(run(tools, 'word_create', { path: 'a.docx', paragraphs: ['two'] }, root)).rejects.toThrow('already exists')
      const replaced = await run(tools, 'word_create', { path: 'a.docx', paragraphs: ['two'], overwrite: true }, root) as any
      expect(replaced.sizeBytes).toBeGreaterThan(0)
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })
})

describe('excel tools', () => {
  test('create, read, and update a workbook', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-excel-'))
    try {
      const tools = mountTools()
      const created = await run(tools, 'excel_create', {
        path: 'data.xlsx',
        sheets: [{ name: 'Data', rows: [['name', 'value'], ['alpha', 1], ['beta', true]] }],
      }, root) as any

      expect(created.sheets).toEqual([{ name: 'Data', rowCount: 3, colCount: 2 }])
      expect((await stat(join(root, 'data.xlsx'))).isFile()).toBe(true)

      const read = await run(tools, 'excel_read', { path: 'data.xlsx', sheet: 'Data' }, root) as any
      expect(read.sheets).toHaveLength(1)
      expect(read.sheets[0].name).toBe('Data')
      expect(read.sheets[0].rows[0]).toEqual(['name', 'value'])
      expect(read.sheets[0].rows[1]).toEqual(['alpha', '1'])

      const updated = await run(tools, 'excel_update', {
        path: 'data.xlsx',
        cell_updates: [{ sheet: 'Data', cell: 'B3', value: 99 }],
      }, root) as any
      expect(updated.cellUpdates).toEqual([{ sheet: 'Data', cell: 'B3' }])

      const reread = await run(tools, 'excel_read', { path: 'data.xlsx', sheet: 'Data' }, root) as any
      expect(reread.sheets[0].rows[2]).toEqual(['beta', '99'])
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  test('update can replace and append whole sheets', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-excel-sheets-'))
    try {
      const tools = mountTools()
      await run(tools, 'excel_create', {
        path: 'book.xlsx',
        sheets: [{ name: 'Old', rows: [['a']] }],
      }, root)
      await run(tools, 'excel_update', {
        path: 'book.xlsx',
        sheets: [
          { name: 'Old', rows: [['new'], ['rows']] },
          { name: 'New', rows: [['x', 'y']] },
        ],
      }, root)
      const read = await run(tools, 'excel_read', { path: 'book.xlsx' }, root) as any
      expect(read.sheets.map((sheet: any) => sheet.name)).toEqual(['Old', 'New'])
      expect(read.sheets[0].rows).toEqual([['new'], ['rows']])
      expect(read.sheets[1].rows).toEqual([['x', 'y']])
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })
})

describe('ppt tools', () => {
  test('create a deck and extract slide text', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-ppt-'))
    try {
      const tools = mountTools()
      const created = await run(tools, 'ppt_create', {
        path: 'deck.pptx',
        title: 'Product Deck',
        slides: [
          { title: 'Problem', bullets: ['Too slow', 'Too manual'], notes: 'Explain pain points' },
          { paragraphs: ['Solution summary'] },
        ],
      }, root) as any

      expect(created.slideCount).toBe(3)
      expect((await stat(join(root, 'deck.pptx'))).isFile()).toBe(true)

      const read = await run(tools, 'ppt_read', { path: 'deck.pptx' }, root) as any
      expect(read.slideCount).toBe(3)
      expect(read.slides).toHaveLength(3)
      expect(read.slides[1].paragraphs.join('\n')).toContain('Problem')
      expect(read.slides[1].paragraphs.join('\n')).toContain('Too slow')
      expect(read.slides[1].notes).toBeDefined()
      expect(read.slides[1].notes[0]).toContain('Explain pain points')
      expect(read.slides[2].paragraphs.join('\n')).toContain('Solution summary')
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })

  test('embeds images from the workspace and reports them on read', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-ppt-img-'))
    try {
      await writeFile(join(root, 'pic.png'), TINY_PNG)
      const tools = mountTools()
      const created = await run(tools, 'ppt_create', {
        path: 'with-image.pptx',
        slides: [
          { title: 'Chart slide', bullets: ['trend is up'], images: [{ path: 'pic.png' }] },
          { paragraphs: ['No images here'] },
        ],
      }, root) as any

      expect(created.slideCount).toBe(2)
      expect(created.imageCount).toBe(1)

      const read = await run(tools, 'ppt_read', { path: 'with-image.pptx' }, root) as any
      expect(read.slides[0].imageCount).toBe(1)
      expect(read.slides[1].imageCount).toBe(0)
      expect(read.slides[0].paragraphs.join('\n')).toContain('Chart slide')

      await expect(run(tools, 'ppt_create', {
        path: 'missing-image.pptx',
        slides: [{ images: [{ path: 'missing.png' }] }],
      }, root)).rejects.toThrow()
      await expect(run(tools, 'ppt_create', {
        path: 'bad-image.pptx',
        slides: [{ images: [{ path: 'pic.txt' }] }],
      }, root)).rejects.toThrow('expected .png or .jpg or .jpeg or .gif')
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })
})

describe('workspace confinement', () => {
  test('resolves relative paths and rejects escapes', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-paths-'))
    const outside = await mkdtemp(join(tmpdir(), 'dsh-office-outside-'))
    try {
      const resolved = await resolveOfficePath(execFor(root), 'sub/plan.docx', ['.docx'], false)
      expect(resolved.absolute).toBe(resolve(root, 'sub/plan.docx'))
      expect(resolved.display).toBe(join('sub', 'plan.docx'))

      await expect(resolveOfficePath(execFor(root), '../escape.docx', ['.docx'], false)).rejects.toThrow('escapes')
      await expect(resolveOfficePath(execFor(root), join(outside, 'escape.docx'), ['.docx'], false)).rejects.toThrow('escapes')
      await expect(resolveOfficePath(execFor(root), 'plan.xlsx', ['.docx'], false)).rejects.toThrow('expected .docx')
    } finally {
      await rm(root, { recursive: true, force: true })
      await rm(outside, { recursive: true, force: true })
    }
  })
})

describe('artifact bytes', () => {
  test('created files start with the expected zip signatures', async () => {
    const root = await mkdtemp(join(tmpdir(), 'dsh-office-zip-'))
    try {
      const tools = mountTools()
      await run(tools, 'word_create', { path: 'a.docx', paragraphs: ['zip'] }, root)
      await run(tools, 'excel_create', { path: 'a.xlsx', sheets: [{ name: 'S', rows: [['x']] }] }, root)
      await run(tools, 'ppt_create', { path: 'a.pptx', slides: [{ title: 'zip' }] }, root)

      for (const name of ['a.docx', 'a.xlsx', 'a.pptx']) {
        const bytes = await readFile(join(root, name))
        expect(bytes.subarray(0, 2).toString('latin1')).toBe('PK')
      }
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  })
})
