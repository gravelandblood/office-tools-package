# Template Detective Skill

日期：2026-06-01

## 定位

`template-detective` 不是一个简单 CLI，而是一套用于 DOCX 模板反推的重型工作流。它把模板探测拆成四层：

1. 确定性抽取：从 DOCX/OOXML 中提取段落、run、表格、section、样式、编号、页眉页脚、格式指纹和证据位置。
2. 规范化 IR：把重复格式压缩成 format atoms，把文档结构、数据字段、规则和冲突用统一 schema 表达。
3. LLM 规则归纳：在规范化 profile/IR 上推断静态文本、动态字段、循环、条件块和格式规则。
4. 渲染验证：用样例数据回放生成 DOCX，再做文本和格式指纹比较。

当前 repo 中已有的 `template infer-format/render/compare-format` 是底层保真工具；真正的泛化模板探测应使用新增的 `template profile` 和 `skills/template-detective` 工作流。

## 新增能力

```powershell
node packages/cli/bin/office-tools.js template profile C:\path\sample.docx --out C:\path\profile.json
node packages/cli/bin/office-tools.js template profile C:\path\sample.docx --summary
```

`profile.json` 包含：

- Word 文本 part：正文、页眉、页脚、脚注、尾注、批注。
- 段落：文本、样式 ID、段落属性、编号、对齐、缩进、spacing、run 列表、格式指纹、证据路径。
- run：文本、run 属性、直接格式、字段/图片/换行等标记、格式指纹。
- 表格：表格属性、grid、行、单元格、边框、底纹、合并、宽度和格式指纹。
- section：页边距、纸张大小、分栏、页眉页脚引用。
- styles/numbering：样式与编号摘要。
- signals：用于提示泛化风险的初步信号。

## 已验证样本

已在 `C:\Code\doctemplate_codex\历史沿革报告.docx` 上运行：

- partCount：23
- wordTextParts：正文、footer1/footer2、header1/header2、footnotes、comments
- paragraphs：333
- runs：553
- tables：10
- sections：2
- textNodes：543
- signals：检测到大量直接 run 格式、多个表格行格式变体、页眉页脚相关段落样式变体

这说明 profile 层已经能覆盖后续 LLM 归纳需要的主要证据，但它还不是最终的模板推理器。

## 后续工程拆分

### P1：多样本对齐

输入多个 profile，对齐段落、表格、标题、标签和重复块，输出候选结构树。

需要重点处理：

- 文本相同但格式不同。
- 格式相同但文本不同。
- 表格行数不同。
- 可选段落或可选章节。
- 页眉页脚与正文中的同名样式。

### P2：规则归纳

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
