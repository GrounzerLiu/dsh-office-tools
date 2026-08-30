/**
 * Workspace-confined path resolution and durable-write helpers shared by all
 * seven Office tools.
 *
 * Every tool binds to the CALLING agent's session cwd (the authoritative
 * `session.header.cwd`) — the model never gets to pick a root. Relative paths
 * resolve against that cwd, absolute paths are accepted only when they stay
 * inside it, and a realpath check on the nearest existing ancestor closes the
 * symlink escape hatch.
 */

import { mkdir, readFile, realpath, rename, rm, stat, writeFile } from 'node:fs/promises'
import { dirname, extname, isAbsolute, join, relative, resolve } from 'node:path'
import { randomUUID } from 'node:crypto'
import JSZip from 'jszip'
import type { ToolRunContext } from '@deepseek-ai/dsh-tools'

/** Hard cap for reading an existing Office file into memory. */
export const MAX_OFFICE_FILE_BYTES = 50 * 1024 * 1024

/**
 * Declared uncompressed ceiling for one zip entry inside an Office file. The
 * 50 MiB read cap only bounds the COMPRESSED bytes; deflate can expand a file
 * a thousandfold, so {@link loadZipGuarded} also checks the archive's own
 * declared sizes before any entry is inflated.
 */
export const MAX_ZIP_ENTRY_BYTES = 256 * 1024 * 1024

/** Declared uncompressed ceiling summed over all entries of one archive. */
export const MAX_ZIP_TOTAL_BYTES = 512 * 1024 * 1024

/** Maximum entries (files + directories) in one archive. */
export const MAX_ZIP_ENTRIES = 100_000

/** Cap for text materialized into a single tool result. */
export const MAX_TEXT_CHARS = 200_000

/** Cap for worksheet cells materialized into a single tool result. */
export const MAX_READ_CELLS = 200_000

/** Cap for worksheet cells accepted by one create/update call. */
export const MAX_WRITE_CELLS = 200_000

export interface ResolvedOfficePath {
  /** The path exactly as the model passed it. */
  input: string
  /** Absolute path used for every fs operation. */
  absolute: string
  /** Path rendered back to the model (workspace-relative when possible). */
  display: string
  /** Lowercased extension including the leading dot. */
  ext: string
}

function workspaceRootOf(exec: ToolRunContext): string {
  const cwd = exec.agent?.session.header.cwd
  if (cwd === undefined || cwd === '') {
    throw new Error('office tools require an active session with a working directory (session.header.cwd is empty)')
  }
  return resolve(cwd)
}

function displayPathOf(root: string, absolute: string): string {
  const rel = relative(root, absolute)
  return rel !== '' && !rel.startsWith('..') && !isAbsolute(rel) ? rel : absolute
}

/**
 * Reject paths that lexically escape the session workspace. The caller should
 * combine this with {@link assertRealAncestorWithin} for symlink safety.
 */
function assertLexicallyWithin(root: string, candidate: string): void {
  const rel = relative(root, candidate)
  if (rel === '' || rel.startsWith('..') || isAbsolute(rel)) {
    throw new Error(`path "${candidate}" escapes the session workspace "${root}"`)
  }
}

/**
 * Walk up from `target` to the nearest existing ancestor and verify its
 * realpath is still inside the real workspace root. The workspace root itself
 * may be a symlink (resolve it first), but no path component may hop outside.
 */
async function assertRealAncestorWithin(root: string, target: string): Promise<void> {
  let cursor = target
  for (;;) {
    try {
      await stat(cursor)
      break
    } catch (error: unknown) {
      const code = error instanceof Error && 'code' in error ? (error as NodeJS.ErrnoException).code : undefined
      if (code !== 'ENOENT' && code !== 'ENOTDIR') throw error
      const parent = dirname(cursor)
      if (parent === cursor) throw new Error(`cannot resolve a real parent for "${target}"`)
      cursor = parent
    }
  }
  const realRoot = resolve(await realpath(root))
  const realAncestor = resolve(await realpath(cursor))
  if (realAncestor === realRoot) return
  assertLexicallyWithin(realRoot, realAncestor)
}

/**
 * Resolve one model-supplied path to a workspace-confined absolute file path.
 *
 * @param exec - the running tool call (carries the session cwd + abort signal).
 * @param rawPath - the model-supplied path string.
 * @param allowedExts - acceptable lowercased extensions WITH dots (e.g. `.docx`).
 * @param mustExist - when true, stat the file and reject directories.
 */
export async function resolveOfficePath(
  exec: ToolRunContext,
  rawPath: string,
  allowedExts: readonly string[],
  mustExist: boolean,
): Promise<ResolvedOfficePath> {
  exec.signal.throwIfAborted()
  if (rawPath.trim() === '') throw new Error('path must be a non-empty string')

  const root = workspaceRootOf(exec)
  const candidate = resolve(isAbsolute(rawPath) ? rawPath : join(root, rawPath))
  assertLexicallyWithin(root, candidate)
  await assertRealAncestorWithin(root, candidate)

  const ext = extname(candidate).toLowerCase()
  if (!allowedExts.includes(ext)) {
    throw new Error(`expected ${allowedExts.join(' or ')} file, got extension "${ext || '(none)'}"`)
  }

  if (mustExist) {
    const info = await stat(candidate)
    if (!info.isFile()) throw new Error(`"${candidate}" is not a regular file`)
  }

  return { input: rawPath, absolute: candidate, display: displayPathOf(root, candidate), ext }
}

/**
 * Read a bounded Office file into memory, observing tool-call cancellation.
 */
export async function readOfficeBuffer(path: string, signal: AbortSignal): Promise<{ buffer: Buffer; sizeBytes: number }> {
  signal.throwIfAborted()
  const info = await stat(path)
  if (!info.isFile()) throw new Error(`"${path}" is not a regular file`)
  if (info.size > MAX_OFFICE_FILE_BYTES) {
    throw new Error(`"${path}" is ${info.size} bytes; office tools refuse files larger than ${MAX_OFFICE_FILE_BYTES} bytes`)
  }
  const buffer = await readFile(path, { signal })
  signal.throwIfAborted()
  return { buffer, sizeBytes: buffer.byteLength }
}

/**
 * jszip@^3.10.1 internal: after `loadAsync`, every entry's `_data` is the
 * CompressedObject built from the central directory, whose `uncompressedSize`
 * is the DECLARED inflated size — readable without inflating anything. The
 * dependency is pinned to ^3.10.1 on purpose; re-verify this shape on any
 * jszip major bump. (jszip's own index.d.ts acknowledges the field.)
 */
interface JszipLoadedEntry {
  name: string
  dir: boolean
  _data?: { uncompressedSize?: unknown }
}

/** Overridable budgets so tests can trip the guard with tiny values. */
export interface ZipGuardLimits {
  maxEntryBytes?: number
  maxTotalBytes?: number
  maxEntries?: number
}

/**
 * Load a zip-backed Office file with jszip and refuse archives whose central
 * directory already declares more uncompressed content than we are ever
 * willing to inflate (the zip-bomb guard), or more entries than we will walk.
 * Returns the loaded instance so callers never parse the buffer twice.
 */
export async function loadZipGuarded(buffer: Buffer, signal: AbortSignal, limits?: ZipGuardLimits): Promise<JSZip> {
  signal.throwIfAborted()
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(buffer)
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error)
    throw new Error(`not a readable zip archive (Office files must be valid .docx/.xlsx/.pptx zips): ${reason}`)
  }
  signal.throwIfAborted()

  const maxEntryBytes = limits?.maxEntryBytes ?? MAX_ZIP_ENTRY_BYTES
  const maxTotalBytes = limits?.maxTotalBytes ?? MAX_ZIP_TOTAL_BYTES
  const maxEntries = limits?.maxEntries ?? MAX_ZIP_ENTRIES

  const entries = Object.values(zip.files) as JszipLoadedEntry[]
  if (entries.length > maxEntries) {
    throw new Error(`zip archive holds ${entries.length} entries; office tools refuse archives with more than ${maxEntries}`)
  }
  let totalBytes = 0
  for (const entry of entries) {
    const declared = entry._data?.uncompressedSize
    if (typeof declared !== 'number' || !Number.isFinite(declared)) continue
    if (declared > maxEntryBytes) {
      throw new Error(`zip entry "${entry.name}" declares ${declared} uncompressed bytes; office tools refuse entries above ${maxEntryBytes} bytes`)
    }
    totalBytes += declared
    if (totalBytes > maxTotalBytes) {
      throw new Error(`zip archive declares more than ${maxTotalBytes} uncompressed bytes in total (at least ${totalBytes} after "${entry.name}"); refusing to inflate it`)
    }
  }
  return zip
}

/**
 * OOXML parts are plain element trees; a DOCTYPE/ENTITY declaration is never
 * legitimate in one. Our extractors never resolve entities, but we refuse to
 * look at such a part at all so entity-expansion payloads die at the door.
 */
export function assertNoXmlDtd(xml: string, label: string): void {
  if (/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml)) {
    throw new Error(`${label} contains a DOCTYPE/ENTITY declaration; office tools refuse such XML parts`)
  }
}

/**
 * Read one zip part as text and refuse DTD/entity-bearing XML before any
 * caller parses it. Returns null when the archive has no such part.
 */
export async function readZipXmlPart(zip: JSZip, name: string, signal: AbortSignal): Promise<string | null> {
  signal.throwIfAborted()
  const file = zip.file(name)
  if (file === null) return null
  const xml: string = await file.async('string')
  assertNoXmlDtd(xml, name)
  return xml
}

/**
 * Write a buffer through a same-directory temp file + rename so a failed
 * generation never leaves a half-written Office document behind.
 */
export async function atomicWriteFile(path: string, buffer: Buffer): Promise<number> {
  const directory = dirname(path)
  await mkdir(directory, { recursive: true })
  const temp = join(directory, `.${randomUUID()}.office-tmp`)
  try {
    await writeFile(temp, buffer, { flag: 'wx', mode: 0o600 })
    await rename(temp, path)
    return buffer.byteLength
  } catch (error) {
    await rm(temp, { force: true }).catch(() => undefined)
    throw error
  }
}

/**
 * Reject an overwrite when `overwrite` is false and the target already exists.
 * Callers use this BEFORE doing expensive generation so the model gets a fast
 * refusal instead of wasted work.
 */
export async function assertMayCreate(path: string, overwrite: boolean): Promise<void> {
  if (overwrite) return
  try {
    await stat(path)
    throw new Error(`"${path}" already exists; pass overwrite: true to replace it`)
  } catch (error: unknown) {
    const code = error instanceof Error && 'code' in error ? (error as NodeJS.ErrnoException).code : undefined
    if (code === 'ENOENT' || code === 'ENOTDIR') return
    if (error instanceof Error && error.message.includes('already exists')) throw error
    throw error
  }
}
