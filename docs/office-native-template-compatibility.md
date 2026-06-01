# Office Native Template Compatibility

日期：2026-06-01

## 核心原则

Template Detective 不是要发明一套替代 Word 的模板语言。它的定位是：

1. 从历史 DOCX 样例中探测结构、字段、循环、条件和冲突。
2. 生成可解释的 Template IR。
3. 把 IR 编译成 Office 原生模板补丁计划。
4. 只有 Office 原生能力覆盖不了的部分，才进入 sidecar metadata。

换句话说，最终模板应尽量还是普通 Word 模板：`.docx/.dotx`、WordprocessingML、styles、content controls、custom XML parts、repeating section controls。

## 能沿用就沿用

必须沿用：

- `styles.xml` 中已有样式。
- `numbering.xml` 中已有编号。
- 段落、run、表格、行、单元格的原始 OOXML 格式。
- section、页眉、页脚、页边距、纸张大小、分栏。
- 图片、绘图、域代码、书签、批注等已有结构。

默认策略是 wrap，而不是 rewrite：给已有范围包一层 content control，或增加 custom XML binding，不重建原文档结构。

## Office 原生映射表

| IR 规则 | Office 原生目标 | 绑定方式 | 说明 |
|---|---|---|---|
| `staticText` | 原始 WordprocessingML | 无 | 不改内容，不改格式。 |
| `slot` | plain-text 或 rich-text content control | custom XML part | 字段写入 `w:tag` / `w:alias`，可绑定 XPath。 |
| `loop` | repeating section content control | custom XML part + renderer | 稳定表格行或块范围优先用重复节控件。 |
| `conditional` | group/rich-text content control | sidecar condition | Word 没有一等条件模板语义，用原生控件标范围，sidecar 表达条件。 |
| `format` | styles / 原 OOXML 属性 | 无或 sidecar audit | 不新造样式系统。 |
| `conflict` | 不自动 patch | sidecar gap | 未解决冲突只进入报告，不静默修改模板。 |

## Sidecar 允许范围

sidecar 只允许补 Office 原生语义缺口：

- 条件表达式。
- 渲染器需要的循环数据路径。
- 冲突分类和人工确认状态。
- LLM 命名建议。
- 验证结果和可接受差异。

sidecar 不允许替代 Word 的格式系统：

- 不写 CSS-like 段落样式。
- 不重定义字体、边框、编号、页眉页脚。
- 不把 OOXML 结构复制成另一套私有文档模型。

## 新增命令

先生成 IR：

```powershell
node packages/cli/bin/office-tools.js template analyze a.docx b.docx --out-ir template-ir.json
```

再生成 Office 原生补丁计划：

```powershell
node packages/cli/bin/office-tools.js template plan-office template-ir.json --out-plan office-plan.json
node packages/cli/bin/office-tools.js template plan-office template-ir.json --summary
```

把安全补丁编译回 Office 原生 DOCX：

```powershell
node packages/cli/bin/office-tools.js template compile-office baseline.docx --plan office-plan.json --out template.docx
```

`office-plan.json` 包含：

- `patches`：真正需要应用到 DOCX 的结构补丁计划，例如 content control、repeating section、条件范围标记。
- `preserves`：确认应原样保留的静态 WordprocessingML 项，不属于实际补丁。
- `gaps`：不能安全自动应用的冲突和缺口。
- `nativeTargets`：每类 IR 规则对应的 Office 原生能力。
- `backendDecision`：编译、渲染、验证分别走哪条后端。
- `compileReadiness`：每个 patch 是否已经具备可安全写入 DOCX 的精确 range。
- `summary`：slot、loop、actual patch、static preserve、sidecar gap 数量。

## 编译顺序

1. 保留原 DOCX 作为 baseline。
2. 对稳定 `slot` 范围包 content control。
3. 给 slot 添加 `w:tag`、`w:alias`，可选 custom XML binding。
4. 对稳定表格行或块范围包 repeating section content control。
5. 对条件块只包 native content control，条件表达式进入 sidecar。
6. 对 unresolved conflict 不改 DOCX，只写 gap。
7. 渲染样例并运行 `template compare-format`。

## 自定义部分走哪条后端

默认选择：走 OfficeCLI / OOXML patcher。

| 自定义部分 | 推荐后端 | 原因 |
|---|---|---|
| 给稳定字段包 content control | OfficeCLI / OOXML patcher | 纯 DOCX 结构修改，可 headless、可验证。 |
| 写入 `w:tag` / `w:alias` | OfficeCLI / OOXML patcher | 标准 WordprocessingML 修改。 |
| 添加 custom XML part 和关系 | OfficeCLI / OOXML patcher | 标准 OPC/OOXML 包操作。 |
| 绑定 content control 到 custom XML XPath | OfficeCLI / OOXML patcher | 写 `w:dataBinding` 即可，不需要启动 Word/WPS。 |
| 包 repeating section content control | OfficeCLI / OOXML patcher，先保守支持稳定表格行 | 本质是 OOXML 范围补丁，但范围边界必须严格验证。 |
| 条件块渲染 | OfficeCLI / OOXML patcher + sidecar | Word 没有一等条件模板语义，渲染器按 sidecar 删除/保留 native content control 范围。 |
| 复杂循环渲染 | OfficeCLI / OOXML patcher + sidecar | 克隆 OOXML 块并保持样式；需要 compare-format 回归。 |
| 域更新、目录刷新、分页观察 | WPS COM / Word COM / JSAPI | 需要真实 Office/WPS 排版或对象模型行为。 |
| 活动文档交互、用户选区插入 | WPS JSAPI | 属于打开中的 WPS 会话上下文。 |
| PDF 转 Word、压缩等产品 UI 能力 | WPS UIA | 模板渲染不使用 UIA。 |

因此，自定义规则不是另起一套模板引擎，而是一个 **Office-native patch/render layer**：

1. 编译期：用 OfficeCLI/OOXML patcher 把可表达部分写成 content controls、custom XML、repeating sections。
2. 渲染期：优先让 Office 原生绑定消费数据；原生不支持的条件和复杂循环，由 sidecar 驱动 OOXML patcher 做删除、克隆、替换。
3. 验证期：先用 headless `compare-format`；需要真实排版时再调用 COM/JSAPI 打开、更新域、另存或导出。

## 官方可直接使用的目标形态

探测完成后的模板应满足：

- Word/WPS 打开时仍是普通 DOCX/DOTX。
- 用户能看到并编辑 content controls。
- 支持使用 Word/WPS/Office 生态识别 `w:sdt`、`w:tag`、`w:alias`、custom XML binding。
- 即使没有我们的渲染器，静态内容和基础字段控件仍然可见、可人工填写。
- 有我们的渲染器时，sidecar 才增强条件、复杂循环、冲突裁决和批量数据填充。

## 禁止事项

- 禁止为了模板化而重建整个 DOCX。
- 禁止把样式搬到私有 JSON 里再生成新样式。
- 禁止对低置信对齐结果生成字段或循环补丁。
- 禁止对 unresolved conflict 静默选择一个样式。
- 禁止让 LLM 直接输出可写入 DOCX 的最终 OOXML；LLM 只能产生候选规则和解释，写入必须经过确定性补丁器。

## 当前边界

当前已落地：

- `template plan-office`：生成 Office 原生补丁计划。
- `template compile-office`：输入 baseline DOCX + office-plan.json，只对安全 scalar slot patch 包 content controls，并写入 custom XML part / relationship / content type。
- 烟测覆盖：多样本分析、plan 生成、safe slot 编译、custom XML 生成、编译后段落/表格结构未丢失。

仍然保守跳过：

- value 不在单个安全 run 里的 slot。
- 需要跨 run 或跨段落包裹的字段。
- 表格循环的 repeating section 实际写入。
- 条件块的实际删除/保留渲染。
- unresolved conflict 的自动裁决。

注意：不能在当前 IR 只定位到段落/表格块时贸然写入 DOCX。例如 `公司名称：某某公司` 这种字段，原生 content control 应包住冒号后的 value，而不是把整个段落含 label 都变成可替换字段。因此 `plan-office` 会为每个 patch 输出 `compileReadiness`，只有定位到 exact run/text-node/table-row/block range 后，patcher 才能实际修改 DOCX。
