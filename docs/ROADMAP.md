# dsh-office-tools 多版本路线图

> 版本节奏与取舍的单一事实来源。v0.4.0（含 0.3.0）已实施；后续版本按需启动，逐流推进、独立提交、测试全绿才前进。

---

# v0.3.0「lighter & safer」

按 A → C → B → D 四个流推进。

## 流 A · npm 包瘦身

- `package.json` 的 `files`：`"lib"` → `"lib/*.js", "lib/types"`，把 ~6 MB 的 `index.js.map` 剔出 npm 包（保留在 git 里用于构建调试）。
- 验证：`npm pack --dry-run`（预期 packed ~2.0 MB → ~1.3 MB）+ `pnpm run check`。

## 流 C · zip 炸弹守卫

- 问题：50 MiB 上限只限压缩后体积，解压可放大百倍，inflate 打爆宿主内存。
- 依据（已核实源码）：jszip 3.10.1 的 `zip.files[name]._data.uncompressedSize` 即 central directory 声明大小，`loadAsync` 后可读、零解压。
- `src/paths.ts` 新增 `loadZipGuarded(buffer, signal): Promise<JSZip>` + 常量：单条目 ≤256 MiB、总量 ≤512 MiB、条目数 ≤100 000，超限即 throw（错误含实际值与上限）；返回 zip 实例避免二次解析；`_data` 私有字段用局部 interface 断言 + 注释锁 `^3.10.1`。
- 接入：`ppt_read`（替换现有 loadAsync）、`excel_read`/`excel_update`（XLSX.read 前预检）、`word_read`（mammoth 前预检，流 B 后成为唯一路径）。生成型工具不涉及。
- 测试：高压缩比 zip（测试内造 ~10 MB 重复文本）+ 可注入小上限触发拒绝；正常文件回归；伪 zip 友好报错。

## 流 B · word_read 去 mammoth 化

- 收益：mammoth 及其唯一引入的 bluebird（bundle 内 172 处引用、Mimosa 误报主力）整棵消失，预期 bundle 3.2 MB → 2.3~2.6 MB。诚实预期：jszip 自带 setimmediate 仍在，误报大减而非清零。
- 第 1 步 golden 对拍（mammoth 还在时）：新增 `tests/word-parity.spec.ts`，`word_create` 生成 fixture（标题/段落/空段/bullets/表格）+ JSZip 手工构造 docx 覆盖 `w:tab`、`w:br`、hyperlink、hyphen、连续空段；把 mammoth 输出硬编码为期望常量。行为契约（已核实 mammoth 源码）：段尾一律 `\n\n` 含末段；`w:tab`→`\t`；`w:br`→空字符串（丢失）；表格格间无分隔；hyperlink 留文本弃 URL；hyphen 映射 `\u2011`/`\u00AD`；页眉页脚脚注不包含。
- 第 2 步替换：`word_read` 改为 `loadZipGuarded` → `word/document.xml` → 复用 `ppt.ts` 正则手法（全局抽 `<w:p>`、段内拼 `<w:t>`、实体解码复用 `decodeXmlEntities`、每段尾 `\n\n`）。表格/超链接内段落被同一正则按文档序捕获，天然对齐。
- 第 3 步清理：删 mammoth 依赖与 `src/mammoth.d.ts`，更新 `build.mjs` 注释。
- 验证：对拍全绿 + `grep -ci bluebird lib/index.js` = 0 + 体积实测。

## 流 D · 发版收尾

- 版本 0.3.0（package.json + dsh.plugin.json）、CHANGELOG（Security/Changed/Perf 附实测数字）、README 双语安全边界、DEVELOPMENT.md 4/5/8 节、本地 AGENTS.md。
- `pnpm run check` + `npm pack --dry-run` 复核。
- 发版动作按 0.2.0 流程（push → tag → `npm publish --otp` → GitHub Release → 重启 DSH）。

---

# v0.4.0「写入增强」（已实施）

路线图原本未单列 0.4.0 细案；0.5.0 的「公式回读」与远期的「word 模板替换」都指向 0.4.0 = `word_update` + Excel 公式写入，据此实施：

- **`word_update`**：向现有 `.docx` 追加段落/项目符号/一个表格。追加片段由与 `word_create` 相同的 `buildDocx` 路径生成（样式一致、转义由 docx 包负责），剥掉临时文档的 `<w:sectPr>` 后缝合进原文档 `<w:sectPr>` 之前，经 zip 守卫加载、原子写回。bullet 复用文档已有 numbering（word_create 产物必有）；无 numbering 的文档降级为普通段落。与 word_create 同上限（10 000 段/20 万格）。
- **Excel 公式写入**：`=…` 前缀字符串在 excel_create / excel_update（整表与 cell_updates）三条路径写成真 `<f>` 公式单元格（SheetJS 默认会存成纯文本，需显式转换）。无缓存值，Excel 打开时计算；回读属 0.5.0。

---

# v0.5.0「读取增强」规划

| 需求 | 说明 | 难度 |
|---|---|---|
| `word_read` 富模式 | `format: "markdown"` 可选参：按 `w:pStyle`（Heading1/2/3 → `#`/`##`/`###`）、列表样式 → `-`，输出结构化 markdown，模型消费效率高于纯文本；默认仍纯文本（向后兼容） | 中 |
| `ppt_read` 表格文本 | 现有 XML walker 上加 `a:tbl` 行列抽取（`\|` 分隔或嵌套数组） | 低-中 |
| `ppt_read` 图片 alt | 抽 `descr` 属性随 imageCount 一起返回 | 低 |
| `excel_read` 公式回读 | 公式写入的配套：返回缓存值为主，无缓存值时返回公式串并标记（行为用测试钉死） | 低 |
| 读取结果 bounded 语义统一 | 各 read 工具的 truncated 标志与预算口径统一文档化 | 低 |

# v0.6.0「工程化与生态」规划

| 需求 | 说明 | 难度 |
|---|---|---|
| README 演示 GIF + 示例 prompt | "一句话生成季度报告三件套"演示，社区插件转化率关键 | 低 |
| CI 矩阵 | node 20/22（engines 声称 ≥20 但只测 22） | 低 |
| npm provenance + 发布工作流 | GitHub Actions OIDC 发布，`--provenance`，包页显示构建来源徽章，顺便把发版动作自动化 | 中 |
| dsh-hub / Atlas 提交 | 材料早就在 `docs/hub-registration.md`，一直没交 | 低 |
| peer 依赖实测刷新 | 本机 dsh 已 0.1.1-rc.2，验证 `@deepseek-ai/*` 的 `^0.1.0-rc.6` 范围语义后放宽/提升 | 低 |
| 开关泛化（可选） | `enableWordTools` / `enableExcelTools`，与 `enablePptTools` 对称；仅当有用户提出同类共存需求再做 | 低 |

# 远期候选（按需评估，不承诺）

- **`ppt_update`**（向现有 deck 追加 slide）：需克隆 slide XML + `_rels` + `[Content_Types].xml` 三处联动，OOXML 手术里最硬的一块。仅当出现真实需求再啃。
- **PDF 导出**：纯 JS 无满意方案，得破"不调外部进程"纪律 → **建议永久放弃**，除非未来出现可靠的纯 JS 渲染器。
- **`.doc/.xls/.ppt` 旧 OLE 格式**：解析复杂、攻击面大、需求稀少 → 建议不做；确有需求时引导用户先转换。
- **`word_update` 模板替换**（占位符变量注入）：word_update 打底后可加 `template_fill` 模式。

# 持续卫生（不占版本号，每次发版前过一遍）

1. 瞄一眼 cdn.sheetjs.com 是否有 >0.20.3 新版，有则升级 + 跑全套测试；
2. 依赖刷新（docx/jszip/pptxgenjs minor）+ `pnpm run check`；
3. Mimosa 提交门禁被拦 → 重试一次（会话基线吸收），持续被拦再显式深扫；
4. CHANGELOG / README / DEVELOPMENT.md / AGENTS.md 四处同步。

# 风险与对策（全程适用）

- 对拍差异：golden 测试先钉行为再动手；
- jszip `_data` 私有字段：锁 `^3.10.1` + 断言注释；
- 体积数字以实测为准，预期值仅参考；
- 每个流独立提交，任何一步出问题可单独 revert，不影响已合入部分。
