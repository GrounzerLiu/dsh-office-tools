# dsh-office-tools 开发总结

> 面向未来 agent / 维护者：本文记录本项目从 0 到发布的全过程、架构决策、已知问题与改进路线。

## 1. 目标

让 DeepSeek Harness（DSH）里的 agent 直接调用 Office 工具，完成真实文档任务：

- 创建、读取 Word `.docx`
- 创建、读取、更新 Excel `.xlsx`
- 创建、读取 PowerPoint `.pptx`
- PPT 支持 PNG/JPG/GIF 图片嵌入与布局

不修改 DSH 源码，不调用 LibreOffice / Word / PowerPoint 外部进程，全部通过纯 JS 库读写 OOXML。

## 2. 框架分析结论

DSH 是 Cordis 插件内核：

- host 插件导出 `name` / `inject` / `apply`；
- 工具通过 `@deepseek-ai/dsh-tools` 的 `ctx.tools.register(defineTool(...))` 注册；
- `defineTool` 声明模型可见 `parameters`、强制校验的 `output.schema`、模型文本投影 `output.render`；
- 执行链路：模型调用 → 参数校验 → `tools/pre-execute` → guards → `tools/execute` → 工具体 → 输出校验 → render → post-execute → result；
- 工具通过 `exec.agent.session.header.cwd` 获取当前会话工作区，通过 `exec.signal` 协作取消。

因此正确做法是新增独立 host 插件，而不是改 agent loop。

## 3. 工具清单

| 工具 | 功能 | 依赖 |
|---|---|---|
| `word_create` | 创建 `.docx`：标题、段落、项目符号、一个表格 | `docx` |
| `word_read` | 提取 `.docx` 纯文本 | `mammoth` |
| `excel_create` | 创建多 sheet `.xlsx` | `xlsx` (SheetJS) |
| `excel_read` | 读取一个或全部 sheet 为标量行 | `xlsx` |
| `excel_update` | 替换/新建整张 sheet，或按 A1 地址写单元格 | `xlsx` |
| `ppt_create` | 创建 16:9 `.pptx`，含标题页/段落/项目符号/备注/图片 | `pptxgenjs` |
| `ppt_read` | 按页提取段落、备注与图片数量 | `jszip` |

## 4. 关键设计

### 4.1 路径安全（`src/paths.ts`）

- 所有路径从 `exec.agent.session.header.cwd` 解析；
- 相对路径留在工作区，绝对路径必须仍在工作区内；
- 最近存在祖先做 `realpath` 校验，防符号链接逃逸；
- 读文件上限 50 MiB；图片单张上限 20 MiB。

### 4.2 原子写与覆盖保护

- 同目录临时文件 + `rename`；
- `overwrite` 默认 `false`；
- 创建前先做存在性检查，避免昂贵生成后才发现冲突。

### 4.3 PPT 图片

- 只允许 `.png/.jpg/.jpeg/.gif`，必须位于工作区；
- 支持显式 `x/y/w/h` 英寸坐标，也支持自动排到文本下方；
- `sizing: contain | cover`，自动模式默认 `contain`；
- `ppt_read` 通过 slide rels 统计每页图片数量，并过滤页码占位符。

### 4.4 构建（`build.mjs`）

- esbuild 把 Office 依赖内联进 `lib/index.js`；
- `@deepseek-ai/*` 和 `cordis` external；
- 使用 `createRequire(import.meta.url)` banner 解决 `xlsx` / `mammoth` 的 CJS 动态 `require("fs"/"stream")` 在 ESM bundle 中崩溃的问题；
- tsc 只发 `lib/types` 声明。

### 4.5 类型缺口

- `mammoth` 无官方类型：`src/mammoth.d.ts`；
- `pptxgenjs` 官方 d.ts 在 NodeNext 下默认导出不可构造：`src/pptxgenjs-shim.d.ts`。

## 5. 测试

`tests/tools.spec.ts`，10 个用例：

- 7 个工具恰好注册一次；
- 所有 schema 通过 `assertSupportedJsonSchema`；
- 在真实 `ToolRuntime` 上注册成功；
- Word/Excel/PPT 创建、读取、更新闭环；
- PPT 图片嵌入、缺失图片/错误扩展名报错；
- 工作区路径逃逸拒绝；
- 生成文件 ZIP 签名 `PK` 校验。

## 6. 发布与生态状态

| 项 | 状态 |
|---|---|
| GitHub | <https://github.com/kw78/dsh-office-tools> |
| npm | `dsh-office-tools@0.1.0` |
| topics | `dsh`, `dsh-plugin`, `deepseek-harness`, `office` 等 |
| CI | GitHub Actions `pnpm run check`，全绿 |
| tag | `v0.1.0` |
| awesome-dsh-plugin | PR #405 已合并 |
| dsh-market | 随 awesome 列表同步 |
| dsh-hub / Atlas | 未收录，候选 entry 在 `docs/hub-registration.md` |

## 7. 安装方式

```bash
# npm（推荐）
dsh plugin --profile web add dsh-office-tools

# GitHub 源码
dsh plugin --profile web add github:kw78/dsh-office-tools
```

安装后重启 DSH。host 插件不会热加载。

## 8. 已知问题与风险

1. `xlsx@0.18.5` 在 npm audit 中报历史高危漏洞；功能和当前场景可用，长期建议换 SheetJS 官方 CDN 新版本或替代解析器。
2. `excel_update` 会由 SheetJS 重写工作簿，图表、宏等高级特性可能丢失。
3. 不支持旧 OLE 格式 `.doc/.xls/.ppt`。
4. `ppt_read` 只提取文本和图片数量，不解析表格、SmartArt、图片内容。
5. Word/PPT 没有“原位编辑”能力，只有整体创建和文本读取。
6. npm 的 `@deepseek-ai/*` 为 optional peerDependencies；裸环境 import 会失败是预期的，必须在 DSH profile 内加载。
7. 当前机器上 web profile 已通过 link 接入插件，但运行中的 DSH 进程未重启时不会加载新 host 插件。

## 9. 改进路线建议

- 升级/替换 `xlsx` 消除 audit 告警；
- 支持 `.doc/.xls/.ppt` 降级读取或 LibreOffice 转换；
- 增加 `excel_format`（条件格式、列宽、图表）能力；
- 增加 `word_update`（模板替换、追加段落）；
- 增加 `ppt_update`（向现有 deck 追加 slide）；
- 读取工具返回结构化 markdown/HTML，而非纯文本；
- 增加 PDF 工具，把 Office 文档导出 PDF；
- 给每个工具加 `presentCall/presentResult` 文件位置定位；
- 补充 npm publish / GitHub release 的自动化 workflow。

## 10. 维护者操作手册

### 10.1 改代码

```bash
pnpm run check     # 必须全绿
git add -A
git commit -m "fix: ..."
git push
```

### 10.2 发布新版本

```bash
npm version patch   # 或 minor/major
pnpm run check
git push --follow-tags
npm publish --registry=https://registry.npmjs.org
```

### 10.3 验证插件真实执行

使用 `tests/tools.spec.ts` 或真实 profile：

```bash
# 在 DSH web profile 重启后，模型直接调用工具即可
```

## 11. 安全注意事项

- 不要在仓库、文档、日志中提交 npm/GitHub token；
- 已用于发布的两个 npm token 应尽快在 npm 后台 revoke；
- 工具本身只访问会话工作区，但插件代码运行在用户权限下，发布前请再次 review 依赖。
