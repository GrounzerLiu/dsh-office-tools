/**
 * ESM host build for dsh-office-tools.
 *
 * The harness profile resolves `main` (`lib/index.js`). Office libraries
 * (docx / xlsx / pptxgenjs / jszip) are bundled into the single host
 * artifact so a profile install never needs to resolve their internals;
 * @deepseek-ai/dsh-* and cordis stay external (the profile's healed
 * node_modules provides them). Type declarations are emitted by tsc.
 *
 * The CJS office library `xlsx` contains dynamic `require("fs")` /
 * `require("stream")` calls. The banner installs a real CommonJS `require`
 * for this ESM artifact so those calls resolve Node builtins instead of
 * hitting esbuild's "Dynamic require is not supported" throw.
 */

import { build } from 'esbuild'
import { mkdirSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
mkdirSync('lib', { recursive: true })

await build({
  entryPoints: ['src/index.ts'],
  outfile: 'lib/index.js',
  bundle: true,
  format: 'esm',
  platform: 'node',
  target: ['node22'],
  sourcemap: true,
  external: ['@deepseek-ai/cordis', '@deepseek-ai/dsh-*'],
  banner: {
    js: "import { createRequire as __createRequire } from 'node:module'; var require = __createRequire(import.meta.url);",
  },
  logLevel: 'info',
})

execFileSync('npx', ['tsc', '-p', 'tsconfig.json'], { stdio: 'inherit', shell: true })
