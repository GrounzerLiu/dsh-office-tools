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

