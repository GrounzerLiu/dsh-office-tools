# AGENTS.md — dsh-office-tools

本文件给未来的 agent / 维护者看。先读 `docs/DEVELOPMENT.md`，里面有完整项目总结、架构与改进路线。

## 一句话定位

面向 DeepSeek Harness（DSH）的 host 插件，注册 7 个模型可调用工具，让 agent 在会话工作区内创建、读取、更新 Word (.docx)、Excel (.xlsx)、PowerPoint (.pptx)，并支持 PPT 图片嵌入。

## 硬性纪律

- 禁止修改 DSH 源码；插件只通过 `cordis.patch.yml` + profile 挂载。
- 所有文件读写必须限制在 `exec.agent.session.header.cwd` 内，路径解析统一走 `src/paths.ts`。
- 工具参数/输出必须通过 `defineTool` 声明；输出必须是 lossless JSON，**不能出现 `undefined` 属性**。
- 改源码后必须 `pnpm run check`（typecheck + 10 个 vitest + build），并提交 `lib/`。
- `lib/index.js` 是构建产物，不要手改。Office 依赖会被 esbuild 内联；`@deepseek-ai/*` 和 `cordis` 保持 external。
- 不要提交凭据、token、`node_modules/`、日志。`.npmrc` 仅本地使用，已被 git exclude。

## 常用命令

```bash
pnpm run check        # typecheck + tests + build
pnpm run build        # 只重新打包 lib
pnpm run test         # 只跑测试
pnpm run typecheck    # 只做类型检查
```

## 关键文件

| 文件 | 作用 |
|---|---|
| `src/index.ts` | 插件入口：`name/inject/apply`，注册 7 个工具 |
| `src/paths.ts` | 工作区路径安全、大小上限、原子写 |
| `src/tools/word.ts` | `word_create` / `word_read` |
| `src/tools/excel.ts` | `excel_create` / `excel_read` / `excel_update` |
| `src/tools/ppt.ts` | `ppt_create` / `ppt_read`（含图片嵌入/读取） |
| `src/tools/shared.ts` | 共享 schema 与 CellValue 类型 |
| `src/mammoth.d.ts` | mammoth 缺失类型的手写声明 |
| `src/pptxgenjs-shim.d.ts` | pptxgenjs 类型 shim |
| `tests/tools.spec.ts` | 10 个测试，含真实 `ToolRuntime` 注册验证 |
| `docs/hub-registration.md` | awesome / dsh-market / dsh-hub 收录材料 |
| `docs/DEVELOPMENT.md` | 完整开发总结 |

## 发布与索引状态

- GitHub: <https://github.com/kw78/dsh-office-tools>
- npm: `dsh-office-tools`
- awesome-dsh-plugin: 已合并（PR #405）
- dsh-hub / Atlas: 未收录，材料在 `docs/hub-registration.md`
