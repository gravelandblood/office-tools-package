# Template Detective Skill

日期：2026-06-01

## 定位

`template-detective` 不是一个简单 CLI，而是一套用于 DOCX 模板反推的重型工作流。它把模板探测拆成四层：

1. 确定性抽取：从 DOCX/OOXML 中提取段落、run、表格、section、样式、编号、页眉页脚、格式指纹和证据位置。
2. 规范化 IR：把重复格式压缩成 format atoms，把文档结构、数据字段、规则和冲突用统一 schema 表达。
3. 规则归纳：先用确定性启发式生成静态文本、slot、loop、冲突候选，再交给 LLM 做语义命名、条件解释和冲突分类。
4. 渲染验证：用样例数据回放生成 DOCX，再做文本和格式指纹比较。

当前 repo 中已有的 `template infer-format/render/compare-format` 是底层保真工具；真正的泛化模板探测应使用 `template profile`、`template analyze` 和 `skills/template-detective` 工作流。

## CLI 入口

抽取单个 DOCX 的全量 profile：

```powershell
node packages/cli/bin/office-tools.js template profile C:\path\sample.docx --out C:\path\profile.json
node packages/cli/bin/office-tools.js template profile C:\path\sample.docx --summary
```

把一个或多个 DOCX 分析成 Template Detective IR：

```powershell
node packages/cli/bin/office-tools.js template analyze C:\path\a.docx C:\path\b.docx --out-ir C:\path\template-ir.json
node packages/cli/bin/office-tools.js template analyze C:\path\a.docx C:\path\b.docx --summary
```

## Profile 内容

`profile.json` 包含：

- Word 文本 part：正文、页眉、页脚、脚注、尾注、批注。
- 段落：文本、样式 ID、段落属性、编号、对齐、缩进、spacing、run 列表、格式指纹、证据路径。
- run：文本、run 属性、直接格式、字段/图片/换行等标记、格式指纹。
- 表格：表格属性、grid、行、单元格、边框、底纹、合并、宽度和格式指纹。
- section：页边距、纸张大小、分栏、页眉页脚引用。
- styles/numbering：样式与编号摘要。
- signals：用于提示泛化风险的初步信号。

## Analyze 内容

`template analyze` 生成的 IR 包含：

- `sources`：样例 DOCX 和 profile 指纹。
- `formatAtoms`：段落、run、表格、行、单元格等格式原子的去重集合。
- `structure`：带证据位置的文档结构节点。
- `dataSchema`：从多样本文本变化中推断出的字段和数组候选。
- `rules`：静态文本、slot、loop 的确定性候选规则。
- `conflicts`：可选块、同位置格式差异、表格结构差异等未解决冲突。
- `profileSignals`：单文档内部的直接格式漂移、表格行格式变体等信号。

当前版本的 `analyze` 是可运行的启发式 IR 生成器，不是最终模板编译器。它的价值在于把后续 LLM 需要看的材料压缩成可验证、可追溯的结构化证据。

## 已验证样本

单样本验证：`C:\Code\doctemplate_codex\历史沿革报告.docx`

- structureNodes：343
- formatAtoms：39
- rules：331
- arrays：10
- profileSignals：13

双样本验证：`历史沿革报告.docx` + `法律尽调报告律师样板.docx`

- structureNodes：677
- formatAtoms：134
- rules：340
- fields：304
- arrays：10
- conflicts：355
- profileSignals：25

双样本结果说明当前按位置对齐的启发式太粗，会把不同报告类型中的段落和表格硬配在一起，所以冲突数量很高。这不是失败，而是下一阶段结构对齐的输入。

## 后续工程拆分

### P1：锚点式多样本对齐

当前 `analyze` 先按 part + index 做粗对齐。下一步需要加入：

- 标题相似度。
- 表格表头相似度。
- 段落格式指纹相似度。
- 编号层级和章节路径。
- 页眉页脚与正文分域对齐。
- 缺失块/新增块的 optional 判断。

### P2：LLM 规则归纳

基于 IR 让 LLM 输出：

- 静态文本规则。
- slot 字段规则和语义命名建议。
- table row / paragraph group loop。
- 条件块。
- 格式规则。
- 冲突分类：稳定规则、条件规则、疑似样例异常、无法解决的冲突。

LLM 只给出候选规则，不能直接改 DOCX。所有规则必须带 evidence。

### P3：模板编译与回放

把规则编译成可渲染模板，并对每个样例做 replay：

```powershell
node packages/cli/bin/office-tools.js template compare-format original.docx rendered.docx
```

未通过 replay 的规则不能默认进入稳定模板，只能进入 report 或待确认项。

## 设计原则

- 先抽全量细节，再压缩摘要。
- 格式推断以 OOXML 和指纹为准，不靠截图猜。
- LLM 用来做语义命名、条件解释和冲突判断。
- 样例内部格式不一致时，不静默归一化；必须标出差异来源和泛化风险。
- 输出必须同时包含模板、数据 schema、规则报告和验证结果。
