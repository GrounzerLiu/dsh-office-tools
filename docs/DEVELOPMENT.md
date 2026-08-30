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
| `word_read` | 提取 `.docx` 纯文本（mammoth 契约的自研提取器，0.3.0） | `jszip` |
| `word_update` | 向现有 `.docx` 追加段落/项目符号/表格（0.4.0） | `docx` + `jszip` |
| `excel_create` | 创建多 sheet `.xlsx`（`=…` 字符串写成公式，0.4.0） | `xlsx` (SheetJS) |
| `excel_read` | 读取一个或全部 sheet 为标量行 | `xlsx` |
| `excel_update` | 替换/新建整张 sheet，或按 A1 地址写单元格（支持公式） | `xlsx` |
| `ppt_create` | 创建 16:9 `.pptx`，含标题页/段落/项目符号/备注/图片 | `pptxgenjs` |
| `ppt_read` | 按页提取段落、备注与图片数量 | `jszip` |

## 4. 关键设计

### 4.1 路径安全（`src/paths.ts`）

- 所有路径从 `exec.agent.session.header.cwd` 解析；
- 相对路径留在工作区，绝对路径必须仍在工作区内；
- 最近存在祖先做 `realpath` 校验，防符号链接逃逸；
- 读文件上限 50 MiB；图片单张上限 20 MiB。

### 4.2 zip 炸弹守卫与 XML 拒绝（`src/paths.ts`，0.3.0）

- 50 MiB 上限只约束压缩后体积，deflate 可放大千倍；`loadZipGuarded(buffer, signal, limits?)` 在 `JSZip.loadAsync` 之后、任何解压之前，读取 jszip 3.10.1 内部的 `zip.files[name]._data.uncompressedSize`（即 central directory 声明值，零解压可读；该私有字段用局部 interface 断言访问，依赖锁 `^3.10.1`，升级需复核）；
- 预算：单条目 ≤256 MiB、整包声明总量 ≤512 MiB、条目数 ≤100 000，超限 throw 并附实际值与上限；`limits` 参数仅供测试注入小预算；
- 四条读路径全部接入：`ppt_read`（替换原 `loadAsync`）、`excel_read` / `excel_update`（`XLSX.read` 前预检）、`word_read`（流 B 后即唯一加载路径）；生成型工具不涉及；
- `assertNoXmlDtd(xml, label)`：OOXML 部件永不含 DOCTYPE/ENTITY，读到即拒绝（我们的正则提取器本就不展开实体，这是把实体膨胀载荷挡在门外）；`readZipXmlPart(zip, name, signal)` 统一「读部件文本 + DTD 拒绝」，word/ppt 共用；
- 伪 zip（随机字节）得到友好的 `not a readable zip archive` 报错，而非 jszip 原始堆栈。

### 4.3 原子写与覆盖保护

- 同目录临时文件 + `rename`；
- `overwrite` 默认 `false`；
- 创建前先做存在性检查，避免昂贵生成后才发现冲突。

### 4.4 PPT 图片

- 只允许 `.png/.jpg/.jpeg/.gif`，必须位于工作区；
- 支持显式 `x/y/w/h` 英寸坐标，也支持自动排到文本下方；
- `sizing: contain | cover`，自动模式默认 `contain`；
- `ppt_read` 通过 slide rels 统计每页图片数量，并过滤页码占位符。

### 4.5 word_read 自研提取器（0.3.0）

- `word_read` 用 `loadZipGuarded` 取 `word/document.xml`，`extractDocxText`（`src/tools/word.ts`）用与 `ppt.ts` 同款正则手法全局抽取 `<w:p>`、段内按文档序拼 `<w:t>`（实体解码复用 `shared.ts` 的 `decodeXmlEntities`），每段尾 `\n\n`；表格/超链接/文本框内的 `<w:p>` 被同一正则按文档序捕获；
- 行为契约逐字节对齐 mammoth 1.11.0 raw text，由 `tests/word-parity.spec.ts` 的 golden 常量钉死（mammoth 卸载前捕获）：段尾一律 `\n\n` 含末段与空段、`w:tab`→`\t`、`w:br`/`w:cr` 丢弃、`w:noBreakHyphen`→U+2011、`w:softHyphen`→U+00AD、hyperlink 留文本弃 URL、表格格间无分隔、页眉页脚脚注不包含；
- 收益：mammoth 及其引入的 bluebird 整棵移除，bundle 3.2 MB → 2.4 MB。

### 4.6 word_update 与 Excel 公式写入（0.4.0）

- `word_update`（`src/tools/word-update.ts` + `word.ts` 的 `buildAppendFragment`/`appendBeforeSectPr`）：追加内容用与 `word_create` 完全相同的 `buildDocx` 路径生成（样式天然一致、转义由 `docx` 包负责），从临时文档抽出 body 子元素、剥掉其尾部 `<w:sectPr>`，缝合进原文档 `word/document.xml` 的 `<w:sectPr>` 之前（页面设置必须保持 body 最后一个子元素）；原 zip 经 `loadZipGuarded` 加载、只替换 document.xml、`generateAsync` DEFLATE 重打包、原子写回；
- bullet 复用文档已有的 numbering（`word_create` 产出的文件必有 numId=1）；无 numbering 的文档追加 bullet 会显示为普通段落（Word 对缺失引用优雅降级，不报修复错误）；
- Excel 公式：SheetJS 会把 `=…` 字符串存成纯文本，`formulaCellOf`/`materializeFormulas`（`excel.ts`）把本次写入的此类单元格改写为 `{ f: … }` 公式单元格，产物携带 `<f>`，Excel 打开时计算；无缓存值，`excel_read` 暂返回空（公式回读在 0.5.0 规划）。

### 4.7 构建（`build.mjs`）

- esbuild 把 Office 依赖与 `@deepseek-ai/schemastery` 内联进 `lib/index.js`（schemastery 用于 Loader 校验 `Config`，与 dsh-notification 同策略）；
- `@deepseek-ai/dsh-*` 和 `cordis` 保持 external；
- `xlsx` 固定来自 SheetJS 官方 CDN tarball（0.20.3），npm 上无修复版（见 0.2.0 变更与第 8 节）；
- 使用 `createRequire(import.meta.url)` banner 解决 `xlsx` 的 CJS 动态 `require("fs"/"stream")` 在 ESM bundle 中崩溃的问题（0.3.0 起 mammoth 已移除，该 banner 只服务 xlsx）；
- tsc 只发 `lib/types` 声明；
- npm 包 `files` 只装 `lib/*.js` + `lib/types`，`lib/index.js.map` 留在 git 用于构建调试（2.0 MB → 550 kB packed）。

### 4.8 类型缺口

- `pptxgenjs` 官方 d.ts 在 NodeNext 下默认导出不可构造：`src/pptxgenjs-shim.d.ts`（mammoth 的补丁已随依赖一起删除）。

### 4.9 配置开关（`src/index.ts`）

- 插件导出 schemastery `Config`，Loader 加载时校验并应用默认值；
- `enablePptTools`（默认 `true`）：`false` 时不注册 `ppt_create`/`ppt_read`，用于与 dsh-ppt 等注册同名 `ppt_create` 的专用演示插件共存（DSH 拒绝同名工具重复注册）；
- `apply(ctx, config)` 内先 `Config(config ?? {})` 解析，未传配置时行为与 0.1.0 完全一致。

## 5. 测试

`tests/tools.spec.ts`（14 例）+ `tests/zip-guard.spec.ts`（7 例，0.3.0）+ `tests/word-parity.spec.ts`（3 例，0.3.0）+ `tests/excel-formula.spec.ts`（3 例，0.4.0）+ `tests/word-update.spec.ts`（5 例，0.4.0），共享挂载器 `tests/harness.ts`：

- 8 个工具恰好注册一次；
- 所有 schema 通过 `assertSupportedJsonSchema`；
- 在真实 `ToolRuntime` 上注册成功；
- `enablePptTools: false` 只注册 6 个 Word/Excel 工具（含真实 ToolRuntime 验证与 Word/Excel 功能闭环）；
- Word/Excel/PPT 创建、读取、更新闭环；
- PPT 图片嵌入、缺失图片/错误扩展名报错；
- 工作区路径逃逸拒绝；
- 生成文件 ZIP 签名 `PK` 校验；
- zip 守卫：高压缩比包默认预算放行、注入小预算触发单条目/总量/条目数拒绝、伪 zip 友好报错、read 工具对伪 zip 拒绝、slide XML 带 DOCTYPE 拒绝；
- word golden 对拍：`word_create` fixture 与手工构造 docx（tab/br/两种连字符/hyperlink/空段/实体/表格/页眉脚注不泄漏）逐字节等于 mammoth 1.11.0 冻结输出；缺 `word/document.xml` 报错；
- word_update：追加段落/项目符号/表格后 `word_read` 全文逐字节校验、插入位置在 sectPr 之前且保留原包其它部件、XML 特殊字符转义、no-op/伪 zip/超上限拒绝、无 sectPr 文档兜底；
- Excel 公式：create/update 单元格/整表替换三条路径的产物含 `<f>`，普通字符串不误转。

## 6. 发布与生态状态

| 项 | 状态 |
|---|---|
| GitHub | <https://github.com/kw78/dsh-office-tools> |
| npm | `dsh-office-tools@0.1.0`（0.2.0 已就绪未发布；0.3.0 待发布） |
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

1. ~~`xlsx@0.18.5` 的 CVE-2023-30533 / CVE-2024-22363~~：0.2.0 起已迁移至 SheetJS 官方 CDN 0.20.3 tarball，且仅作构建期依赖（内联进 bundle，运行时不解析）。
2. ~~zip 炸弹面~~：0.3.0 起读路径全部经 `loadZipGuarded` 预检声明大小（256 MiB/条目、512 MiB/整包、100 000 条目），见 4.2。残余面：central directory 声明造假的 zip（声明小、实际膨胀大）不在声明预算内——jszip 解压结束时会因 `uncompressed data size mismatch` 报错，但内存峰值在报错前已发生；50 MiB 压缩上限把最坏情形约束在可控量级。
3. `excel_read` / `excel_update` 把整个 buffer 交给 SheetJS，无法逐部件做 DTD 检查；SheetJS 自带解析器不做实体展开/外部实体。
4. `excel_update` 会由 SheetJS 重写工作簿，图表、宏等高级特性可能丢失。
5. 不支持旧 OLE 格式 `.doc/.xls/.ppt`。
6. `ppt_read` 只提取文本和图片数量，不解析表格、SmartArt、图片内容。
7. Word 有 `word_update` 追加能力（0.4.0），但仍没有原位编辑/删除/模板替换；PPT 没有更新能力。
8. npm 的 `@deepseek-ai/*` 为 optional peerDependencies；裸环境 import 会失败是预期的，必须在 DSH profile 内加载。
9. 当前机器上 web profile 已通过 link 接入插件，但运行中的 DSH 进程未重启时不会加载新 host 插件。

## 9. 改进路线建议

版本化的路线与取舍记录见 [docs/ROADMAP.md](ROADMAP.md)（0.5.0 读取增强、0.6.0 工程化与生态、远期候选）。历史上的候选清单：

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
