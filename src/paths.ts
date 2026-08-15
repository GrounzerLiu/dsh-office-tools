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
import type { ToolRunContext } from '@deepseek-ai/dsh-tools'

/** Hard cap for reading an existing Office file into memory. */
export const MAX_OFFICE_FILE_BYTES = 50 * 1024 * 1024

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
