declare module 'mammoth' {
  export interface MammothMessage {
    type: string
    message: string
    error?: Error
  }

  export interface RawTextResult {
    value: string
    messages: MammothMessage[]
  }

  export function extractRawText(input: { buffer: Buffer }): Promise<RawTextResult>
}
