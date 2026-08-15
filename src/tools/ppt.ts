/**
 * PowerPoint (.pptx) tools: `ppt_create` builds slide decks with pptxgenjs
 * (text, bullets, speaker notes, and PNG/JPG/GIF images); `ppt_read` unzips a
 * deck with JSZip and extracts paragraph text, speaker notes, and per-slide
 * image counts without any native dependency.
 */

import { stat } from 'node:fs/promises'
import pptxgen from 'pptxgenjs'
import JSZip from 'jszip'
import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import type { ToolRunContext } from '@deepseek-ai/dsh-tools'
import { atomicWriteFile, assertMayCreate, MAX_TEXT_CHARS, readOfficeBuffer, resolveOfficePath } from '../paths.ts'
import { FILE_RESULT_SCHEMA } from './shared.ts'

/** Image formats pptxgenjs embeds natively. */
const IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif'] as const

/** One embedded image may not exceed this size. */
const MAX_IMAGE_BYTES = 20 * 1024 * 1024

/** Safety cap for images on a single slide. */
const MAX_IMAGES_PER_SLIDE = 20

interface SlideImageSpec {
  path: string
  x?: number
  y?: number
  w?: number
  h?: number
  sizing?: 'contain' | 'cover'
}

interface ResolvedSlideImage extends SlideImageSpec {
  absolute: string
}

interface SlideSpec {
  title?: string
  paragraphs?: string[]
  bullets?: string[]
  notes?: string
  images?: SlideImageSpec[]
}

interface PptCreateArgs {
  path: string
  title?: string
  slides?: SlideSpec[]
  overwrite?: boolean
}

interface PptReadArgs {
  path: string
  max_chars?: number
}

const SLIDE_SUMMARY_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    index: { type: 'integer', required: true },
    title: { type: 'string' },
    paragraphs: {
      type: 'array',
      required: true,
      items: { type: 'string' },
    },
    notes: {
      type: 'array',
      items: { type: 'string' },
    },
    imageCount: { type: 'integer', required: true },
  },
} as const

const PPT_CREATE_OUTPUT = {
  type: 'object',
  additionalProperties: false,
  properties: {
    ...FILE_RESULT_SCHEMA.properties,
    title: { type: 'string' },
    slideCount: { type: 'integer', required: true },
    imageCount: { type: 'integer', required: true },
  },
} as const

const PPT_READ_OUTPUT = {
  type: 'object',
  additionalProperties: false,
  properties: {
    path: { type: 'string', required: true },
    slideCount: { type: 'integer', required: true },
    slides: {
      type: 'array',
      required: true,
      items: SLIDE_SUMMARY_SCHEMA,
    },
    truncated: { type: 'boolean', required: true },
    sizeBytes: { type: 'integer', required: true },
  },
} as const

function assertPositiveCoordinate(value: number | undefined, label: string, slideIndex: number, imageIndex: number): void {
  if (value === undefined) return
  if (!Number.isFinite(value) || value <= 0 || value > 100) {
    throw new Error(`slide ${slideIndex} image ${imageIndex + 1} ${label} must be a positive number of inches (0-100)`)
  }
}

function validateSlideSpecs(slides: SlideSpec[]): void {
  if (slides.length === 0) throw new Error('slides must contain at least one slide')
  if (slides.length > 200) throw new Error('too many slides (maximum 200)')
  for (const [slideIndex, slide] of slides.entries()) {
    const hasContent = (slide.title?.trim().length ?? 0) > 0
      || (slide.paragraphs?.length ?? 0) > 0
      || (slide.bullets?.length ?? 0) > 0
      || (slide.images?.length ?? 0) > 0
    if (!hasContent) throw new Error(`slide ${slideIndex + 1} is empty; give it a title, paragraphs, bullets, or images`)
    if ((slide.paragraphs?.length ?? 0) + (slide.bullets?.length ?? 0) > 500) {
      throw new Error(`slide ${slideIndex + 1} has too many text blocks (maximum 500)`)
    }
    const images = slide.images ?? []
    if (images.length > MAX_IMAGES_PER_SLIDE) {
      throw new Error(`slide ${slideIndex + 1} has too many images (maximum ${MAX_IMAGES_PER_SLIDE})`)
    }
    for (const [imageIndex, image] of images.entries()) {
      assertPositiveCoordinate(image.x, 'x', slideIndex + 1, imageIndex)
      assertPositiveCoordinate(image.y, 'y', slideIndex + 1, imageIndex)
      assertPositiveCoordinate(image.w, 'w', slideIndex + 1, imageIndex)
      assertPositiveCoordinate(image.h, 'h', slideIndex + 1, imageIndex)
      if (image.sizing !== undefined && (image.w === undefined || image.h === undefined)) {
        throw new Error(`slide ${slideIndex + 1} image ${imageIndex + 1} uses sizing; provide both w and h`)
      }
    }
  }
}

/**
 * Resolve every image path against the session workspace and verify it is an
 * existing, supported, bounded image file. PptxGenJS reads the file itself;
 * we only pre-flight the path and size here.
 */
async function resolveSlideImages(exec: ToolRunContext, slides: SlideSpec[]): Promise<ResolvedSlideImage[][]> {
  return Promise.all(slides.map(async (slide, slideIndex) => {
    const images = slide.images ?? []
    return Promise.all(images.map(async (image, imageIndex) => {
      const resolved = await resolveOfficePath(exec, image.path, IMAGE_EXTENSIONS, true)
      const info = await stat(resolved.absolute)
      if (info.size > MAX_IMAGE_BYTES) {
        throw new Error(`slide ${slideIndex + 1} image ${imageIndex + 1} "${image.path}" is ${info.size} bytes; maximum embedded image size is ${MAX_IMAGE_BYTES} bytes`)
      }
      return { ...image, absolute: resolved.absolute }
    }))
  }))
}

function addSlideContent(pptx: pptxgen, spec: SlideSpec, first: boolean, images: ResolvedSlideImage[]): void {
  const slide = pptx.addSlide()
  const hasTitle = spec.title !== undefined && spec.title.trim() !== ''

  if (first && hasTitle) {
    // The first titled slide gets the classic centered title treatment.
    slide.addText(spec.title!, {
      x: 0.8, y: 1.2, w: 8.4, h: 1.2,
      fontSize: 32, bold: true, align: 'center', color: '1F3864',
    })
  } else if (hasTitle) {
    slide.addText(spec.title!, {
      x: 0.8, y: 0.35, w: 8.4, h: 0.9,
      fontSize: 26, bold: true, color: '1F3864',
    })
  }

  const top = first && hasTitle ? 2.7 : hasTitle ? 1.5 : 0.8
  let y = top

  if ((spec.paragraphs?.length ?? 0) > 0) {
    for (const paragraph of spec.paragraphs!) {
      if (y > 6.4) break
      slide.addText(paragraph, { x: 0.8, y, w: 8.4, h: 0.7, fontSize: 18, valign: 'top' })
      y += 0.8
    }
    y += 0.2
  }

  if ((spec.bullets?.length ?? 0) > 0) {
    slide.addText(spec.bullets!.map(item => ({ text: item, options: { bullet: true } })), {
      x: 0.8, y, w: 8.4, h: Math.min(4.5, Math.max(1, spec.bullets!.length * 0.6)),
      fontSize: 18, valign: 'top', lineSpacingMultiple: 1.2,
    })
  }

  if (images.length > 0) {
    const automatic = images.filter(image =>
      image.x === undefined && image.y === undefined && image.w === undefined && image.h === undefined)
    const automaticHeight = Math.max(0.6, Math.min(3.2, (6.6 - Math.min(y, 6.4)) / Math.max(1, automatic.length)))
    let imageY = Math.min(y + 0.25, 6.5)
    for (const image of images) {
      const options: {
        path: string
        x?: number
        y?: number
        w?: number
        h?: number
        sizing?: { type: 'contain' | 'cover'; w: number; h: number }
      } = { path: image.absolute }
      if (image.x !== undefined) options.x = image.x
      if (image.y !== undefined) options.y = image.y
      if (image.w !== undefined) options.w = image.w
      if (image.h !== undefined) options.h = image.h
      if (image.sizing !== undefined) {
        options.sizing = { type: image.sizing, w: image.w!, h: image.h! }
      }
      if (automatic.includes(image)) {
        options.x = 0.8
        options.y = imageY
        options.w = 8.4
        options.h = automaticHeight
        options.sizing = { type: image.sizing ?? 'contain', w: 8.4, h: automaticHeight }
        imageY += automaticHeight + 0.15
      }
      slide.addImage(options)
    }
  }

  if (spec.notes !== undefined && spec.notes.trim() !== '') {
    slide.addNotes(spec.notes)
  }
}

async function buildPptx(args: PptCreateArgs, imagesBySlide: ResolvedSlideImage[][]): Promise<Buffer> {
  const pptx = new pptxgen()
  pptx.layout = 'LAYOUT_WIDE'
  pptx.author = 'DeepSeek Harness'
  pptx.company = 'DSH'
  pptx.subject = args.title ?? 'Presentation'
  pptx.title = args.title ?? 'Presentation'

  let first = true
  if (args.title !== undefined && args.title.trim() !== '') {
    addSlideContent(pptx, { title: args.title }, true, [])
    first = false
  }
  const slides = args.slides ?? []
  for (let index = 0; index < slides.length; index += 1) {
    addSlideContent(pptx, slides[index]!, first, imagesBySlide[index] ?? [])
    first = false
  }

  const output = await pptx.write({ outputType: 'nodebuffer' })
  return Buffer.isBuffer(output) ? output : Buffer.from(output as ArrayBuffer)
}

function decodeXmlEntities(value: string): string {
  return value.replace(/&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (entity, code: string) => {
    if (code === 'amp') return '&'
    if (code === 'lt') return '<'
    if (code === 'gt') return '>'
    if (code === 'quot') return '"'
    if (code === 'apos') return "'"
    const number = code.startsWith('#x') ? Number.parseInt(code.slice(2), 16) : Number.parseInt(code.slice(1), 10)
    return Number.isFinite(number) ? String.fromCodePoint(number) : entity
  })
}

/** Extract one `<a:p>` paragraph as a plain-text string. */
function paragraphText(paragraphXml: string): string {
  const runs: string[] = []
  const runPattern = /<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g
  for (const match of paragraphXml.matchAll(runPattern)) runs.push(match[1] ?? '')
  const text = decodeXmlEntities(runs.join('').replace(/<a:br\b[^>]*\/>/g, '\n'))
  return text
}

/**
 * Split slide/notes XML into paragraph strings. `skipFields` drops
 * auto-generated field paragraphs (slide-number placeholders), which carry no
 * author content and would otherwise surface as stray digits in `ppt_read`.
 */
function extractParagraphs(xml: string, skipFields: boolean): string[] {
  const paragraphs: string[] = []
  const pattern = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g
  for (const match of xml.matchAll(pattern)) {
    const paragraph = match[1] ?? ''
    if (skipFields && /<a:fld\b/.test(paragraph)) continue
    const text = paragraphText(paragraph)
    if (text.trim() !== '') paragraphs.push(text)
  }
  return paragraphs
}

function decodeRelationshipTarget(xml: string): string | undefined {
  const match = /Target="([^"]*notesSlides\/notesSlide(\d+)\.xml)"/.exec(xml)
  if (match === null) return undefined
  return `ppt/notesSlides/notesSlide${match[2]}.xml`
}

function slideIndex(name: string): number {
  const match = /slide(\d+)\.xml$/.exec(name)
  return match === null ? Number.MAX_SAFE_INTEGER : Number.parseInt(match[1]!, 10)
}

/** Count image relationships on one slide. */
async function countSlideImages(zip: JSZip, slideNumber: number): Promise<number> {
  const relationship = zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`)
  if (relationship === null) return 0
  const xml = await relationship.async('string')
  return [...xml.matchAll(/Type="[^"]*\/image"/g)].length
}

interface SlideReadData {
  xmls: string[]
  notes: Array<string | undefined>
  imageCounts: number[]
}

async function readSlideXml(zip: JSZip): Promise<SlideReadData> {
  const slideFiles = zip.file(/ppt\/slides\/slide[0-9]+\.xml$/)
  slideFiles.sort((a, b) => slideIndex(a.name) - slideIndex(b.name))
  const xmls = await Promise.all(slideFiles.map(file => file.async('string')))

  const notes = await Promise.all(Array.from({ length: xmls.length }, async (_, index) => {
    const slideNumber = index + 1
    const relationship = zip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`)
    let notesName = `ppt/notesSlides/notesSlide${slideNumber}.xml`
    if (relationship !== null) {
      const target = decodeRelationshipTarget(await relationship.async('string'))
      if (target !== undefined) notesName = target
    }
    const noteFile = zip.file(notesName)
    if (noteFile === null) return undefined
    const paragraphs = extractParagraphs(await noteFile.async('string'), true)
    return paragraphs.length === 0 ? undefined : paragraphs.join('\n')
  }))

  const imageCounts = await Promise.all(Array.from({ length: xmls.length }, (_, index) => countSlideImages(zip, index + 1)))
  return { xmls, notes, imageCounts }
}

function registerPptCreate(ctx: Context): () => void {
  return ctx.tools.register(defineTool({
    name: 'ppt_create',
    description:
      'Create a PowerPoint .pptx presentation in the session workspace. Optionally start with a title slide, then add slides with a title, body paragraphs, bullet points, speaker notes, and PNG/JPG/GIF images. '
      + 'Image paths are workspace files; give x/y/w/h in inches to position them explicitly, or omit them for automatic centered placement below the text. '
      + 'The deck uses the 16:9 widescreen layout.',
    parameters: {
      path: {
        type: 'string',
        required: true,
        description: 'Output path. Relative paths resolve against the session workspace; the extension must be .pptx.',
      },
      title: {
        type: 'string',
        description: 'Deck title. When provided, a title slide is inserted before the explicit slides.',
      },
      slides: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            title: { type: 'string', description: 'Slide title.' },
            paragraphs: {
              type: 'array',
              items: { type: 'string' },
              description: 'Body paragraphs rendered as plain text boxes.',
            },
            bullets: {
              type: 'array',
              items: { type: 'string' },
              description: 'Bullet list items rendered after the paragraphs.',
            },
            notes: { type: 'string', description: 'Speaker notes for this slide.' },
            images: {
              type: 'array',
              items: {
                type: 'object',
                additionalProperties: false,
                properties: {
                  path: {
                    type: 'string',
                    required: true,
                    description: 'Image file inside the session workspace: .png, .jpg, .jpeg, or .gif (absolute paths must stay inside the workspace).',
                  },
                  x: { type: 'number', description: 'Left position in inches on the 13.33x7.5 slide. Omit for automatic placement.' },
                  y: { type: 'number', description: 'Top position in inches. Omit for automatic placement.' },
                  w: { type: 'number', description: 'Display width in inches. Omit to use the native image size.' },
                  h: { type: 'number', description: 'Display height in inches. Omit to use the native image size.' },
                  sizing: {
                    type: 'string',
                    enum: ['contain', 'cover'],
                    description: 'Fit mode inside the w x h box. Requires w and h; defaults to contain for automatic placement.',
                  },
                },
              },
              description: 'Images to embed on this slide, drawn after the text content.',
            },
          },
        },
        description: 'Slides in presentation order. Optional when a title is provided.',
      },
      overwrite: {
        type: 'boolean',
        description: 'Replace the file when it already exists. Defaults to false.',
      },
    },
    output: {
      schema: PPT_CREATE_OUTPUT,
      render: (_args, value: any) => [{
        type: 'text',
        text: `Created PowerPoint ${value.path} (${value.sizeBytes} bytes; ${value.slideCount} slide(s), ${value.imageCount} image(s)).`,
      }],
    },
    presentCall: (args) => ({
      card: 'generic',
      title: `Create ${args.path}`,
      kind: 'edit',
      locations: [{ path: args.path }],
    }),
    async execute(args, exec: ToolRunContext) {
      const target = await resolveOfficePath(exec, args.path, ['.pptx'], false)
      await assertMayCreate(target.absolute, args.overwrite ?? false)
      if ((args.slides?.length ?? 0) > 0) validateSlideSpecs(args.slides!)
      if (args.title === undefined && (args.slides?.length ?? 0) === 0) {
        throw new Error('ppt_create needs a title or at least one slide')
      }
      exec.signal.throwIfAborted()

      const slides = args.slides ?? []
      const imagesBySlide = await resolveSlideImages(exec, slides)
      const imageCount = imagesBySlide.reduce((sum, images) => sum + images.length, 0)
      const buffer = await buildPptx(args, imagesBySlide)
      exec.signal.throwIfAborted()
      const sizeBytes = await atomicWriteFile(target.absolute, buffer)
      const result: { path: string; sizeBytes: number; title?: string; slideCount: number; imageCount: number } = {
        path: target.display,
        sizeBytes,
        slideCount: (args.title !== undefined && args.title.trim() !== '' ? 1 : 0) + slides.length,
        imageCount,
      }
      if (args.title !== undefined && args.title.trim() !== '') result.title = args.title
      return result
    },
  }))
}

function registerPptRead(ctx: Context): () => void {
  return ctx.tools.register(defineTool({
    name: 'ppt_read',
    description:
      'Extract text from an existing .pptx presentation: every slide\'s paragraphs, speaker notes, and embedded image count, in slide order. '
      + 'Use it to understand or summarize a deck before editing it.',
    parameters: {
      path: {
        type: 'string',
        required: true,
        description: 'Path to the .pptx file, relative to the session workspace or absolute inside it.',
      },
      max_chars: {
        type: 'integer',
        description: `Maximum characters returned across the deck. Defaults to ${MAX_TEXT_CHARS}.`,
      },
    },
    output: {
      schema: PPT_READ_OUTPUT,
      render: (_args, value: any) => [{
        type: 'text',
        text: value.slides.map((slide: any) =>
          `Slide ${slide.index}${slide.title !== undefined ? ` — ${slide.title}` : ''} (images: ${slide.imageCount}):\n`
          + slide.paragraphs.map((paragraph: string) => `- ${paragraph}`).join('\n')
          + (slide.notes !== undefined ? `\nNotes: ${slide.notes.join(' | ')}` : ''),
        ).join('\n\n') + (value.truncated ? '\n[text truncated]' : ''),
      }],
    },
    presentCall: (args) => ({
      card: 'generic',
      title: `Read ${args.path}`,
      kind: 'read',
      locations: [{ path: args.path }],
    }),
    async execute(args, exec: ToolRunContext) {
      const target = await resolveOfficePath(exec, args.path, ['.pptx'], true)
      const { buffer, sizeBytes } = await readOfficeBuffer(target.absolute, exec.signal)
      const zip = await JSZip.loadAsync(buffer)
      const { xmls, notes, imageCounts } = await readSlideXml(zip)
      if (xmls.length === 0) throw new Error('the .pptx contains no slides')

      const maxChars = Math.min(Math.max(args.max_chars ?? MAX_TEXT_CHARS, 1), MAX_TEXT_CHARS)
      const slides: Array<{ index: number; title?: string; paragraphs: string[]; notes?: string[]; imageCount: number }> = []
      let totalChars = 0
      let truncated = false

      for (let index = 0; index < xmls.length; index += 1) {
        const paragraphs = extractParagraphs(xmls[index]!, false)
        const noteText = notes[index]
        const noteParagraphs = noteText === undefined || noteText.trim() === '' ? undefined : [noteText]
        const body = paragraphs
        const remainingChars = Math.max(0, maxChars - totalChars)
        let slideChars = 0
        const bounded = body.map((paragraph) => {
          if (slideChars >= remainingChars) return ''
          const retained = paragraph.slice(0, remainingChars - slideChars)
          slideChars += retained.length
          return retained
        })
        const noteBounded = noteParagraphs === undefined ? undefined : [noteParagraphs[0]!.slice(0, Math.max(0, remainingChars - slideChars))]
        const bodyChars = body.reduce((sum, paragraph) => sum + paragraph.length, 0)
        const noteChars = noteParagraphs?.[0]?.length ?? 0
        totalChars += slideChars + (noteBounded?.[0]?.length ?? 0)
        if (bodyChars + noteChars > slideChars + (noteBounded?.[0]?.length ?? 0)) truncated = true
        const slide: { index: number; title?: string; paragraphs: string[]; notes?: string[]; imageCount: number } = {
          index: index + 1,
          paragraphs: bounded.filter(paragraph => paragraph !== ''),
          imageCount: imageCounts[index] ?? 0,
        }
        if (noteBounded !== undefined) slide.notes = noteBounded
        slides.push(slide)
      }

      return { path: target.display, slideCount: slides.length, slides, truncated, sizeBytes }
    },
  }))
}

export function registerPptTools(ctx: Context): () => void {
  const disposers = [registerPptCreate(ctx), registerPptRead(ctx)]
  return () => disposers.forEach(dispose => dispose())
}
