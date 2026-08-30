/**
 * Shared in-process tool harness: mounts the plugin on a fake Cordis context
 * and runs tools against a temp workspace root, the same way tools.spec.ts
 * drives the real registry.
 */

import type { Context } from '@deepseek-ai/cordis'
import type { ToolDefinition, ToolRunContext } from '@deepseek-ai/dsh-tools'
import { apply } from '../src/index.ts'

interface ToolRegistryLike {
  register(definition: ToolDefinition): () => void
}

interface AgentLike {
  session: { header: { cwd: string } }
}

export function execFor(root: string): ToolRunContext {
  return {
    signal: new AbortController().signal,
    agent: { session: { header: { cwd: root } } } as AgentLike,
    callId: 'test-call',
    name: 'test',
    arguments: {},
  } as unknown as ToolRunContext
}

export function mountTools(config?: { enablePptTools?: boolean }): Map<string, ToolDefinition> {
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

export async function run(tools: Map<string, ToolDefinition>, name: string, args: Record<string, unknown>, root: string) {
  const tool = tools.get(name)
  if (tool === undefined) throw new Error(`tool ${name} should be registered`)
  const exec = execFor(root)
  const invoke = tool.execute.bind(tool)
  return invoke(args, exec)
}
