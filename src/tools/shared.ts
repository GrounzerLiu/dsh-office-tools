/**
 * Shared schema fragments and cell types for the Office tool suite. Keeping
 * the schemas in one place keeps the seven tool contracts consistent.
 */

import type { ValueSchemaSpec } from '@deepseek-ai/dsh-tools'

/** One cell value accepted by the Excel tools. */
export const CELL_VALUE_SCHEMA = {
  oneOf: [
    { type: 'string' },
    { type: 'number' },
    { type: 'boolean' },
    { type: 'null' },
  ],
} as const satisfies ValueSchemaSpec

/** One spreadsheet row. */
export const ROW_SCHEMA = {
  type: 'array',
  items: CELL_VALUE_SCHEMA,
} as const satisfies ValueSchemaSpec

/** Common success echo for a created/replaced file. */
export const FILE_RESULT_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    path: { type: 'string', required: true },
    sizeBytes: { type: 'integer', required: true },
  },
} as const satisfies ValueSchemaSpec

export type CellValue = string | number | boolean | null

export type CellRow = CellValue[]

/**
 * Decode the XML entities that can legally appear in OOXML text content: the
 * five predefined names plus decimal/hex character references.
 */
export function decodeXmlEntities(value: string): string {
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

