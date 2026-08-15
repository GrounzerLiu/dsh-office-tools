declare module 'pptxgenjs' {
  export interface PptxTextBlock {
    text: string
    options?: { bullet?: boolean }
  }

  export interface PptxTextOptions {
    x?: number
    y?: number
    w?: number
    h?: number
    fontSize?: number
    bold?: boolean
    align?: 'left' | 'center' | 'right'
    color?: string
    valign?: 'top' | 'middle' | 'bottom'
    lineSpacingMultiple?: number
    [key: string]: unknown
  }

  export interface PptxImageOptions {
    path: string
    x?: number
    y?: number
    w?: number
    h?: number
    sizing?: { type: 'contain' | 'cover'; w: number; h: number }
  }

  export interface PptxSlide {
    addText(text: string | PptxTextBlock[], options?: PptxTextOptions): void
    addNotes(text: string): void
    addImage(options: PptxImageOptions): void
  }

  export interface PptxWriteOptions {
    outputType: 'nodebuffer'
  }

  export default class PptxGenJS {
    layout: string
    author: string
    company: string
    subject: string
    title: string
    addSlide(): PptxSlide
    write(options: PptxWriteOptions): Promise<string | ArrayBuffer | Buffer>
  }
}
