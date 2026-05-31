# WPS UIA 能力与 CLI 设计

日期：2026-05-31

这份文档把 WPS UI Automation 限定为产品后端，而不是一堆暴露给用户的启动参数。

## 为什么需要 WPS UIA

WPS UIA 只应该用于稳定文件 API、COM 或 JSAPI 无法覆盖，但 WPS 桌面产品已经提供的能力。

适合 UIA 的能力：

- WPS 通过可见桌面 UI 暴露该功能。
- 功能依赖 WPS 账号、会员、云端或产品转换能力。
- 用本地确定性库很难复刻效果，例如扫描 PDF OCR 后的版式还原。
- 流程可以抽象成：启动、配置、开始、等待产物、清理窗口。

不适合 UIA 的能力：

- 普通 OOXML 读写。
- 本地库可稳定完成的 PDF 合并、拆分、旋转。
- 需要大量人工判断的编辑类操作。
- 任何试图绕过 WPS 授权的行为。

## 已验证能力

| 能力 | 命令 | 状态 | 说明 |
|---|---|---|---|
| PDF 转 Word | `pdf to-word` | 已落地并验证 | 支持 `--out`，转换后自动清理窗口。 |
| PDF 转 Excel | `pdf to-excel` | 已落地并验证 | 本机真实生成 `.xlsx`，复用 WPS PDF 转换窗口。 |
| PDF 转 PPT | `pdf to-ppt` | 已落地并验证 | 本机真实生成 `.pptx`，复用 WPS PDF 转换窗口。 |
| PDF 转图片型 PDF | `pdf to-image-pdf` | 已落地并验证 | 本机真实生成扫描版 PDF，支持 `--out` 和自动清理窗口。 |
| PDF 压缩 | `pdf compress` | 已落地并验证 | 支持 `--out` 和压缩质量，完成后自动清理 `PDF压缩` 窗口。 |
| 文件瘦身 | `file slim` | 已落地并验证 | 本机真实生成瘦身副本，支持 `--out` 和自动清理窗口。 |

验证样本：

- 输入：`C:\Code\test\pdfreadtest\工商底档-2023.10.26变更登记.pdf`
- 输出：
  - `C:\Code\test\pdfreadtest\cli-product-pdf2excel.xlsx`
  - `C:\Code\test\pdfreadtest\cli-product-pdf2ppt.pptx`
  - `C:\Code\test\pdfreadtest\cli-product-regression-word.docx`
  - `C:\Code\test\pdfreadtest\cli-product-image-pdf.pdf`
  - `C:\Code\test\pdfreadtest\cli-product-compress-regression.pdf`
  - `C:\Code\test\pdfreadtest\cli-product-file-slim.pdf`

`pdf to-image-pdf` 回归结果：

- WPS 默认产物命名：`<输入名>_扫描版.pdf`
- 产品命令输出：`C:\Code\test\pdfreadtest\cli-product-image-pdf.pdf`
- 输出大小：`16827307` bytes
- 收尾验证：未残留 `WPS PDF转换` / `WPS Office` 产物窗口，未残留 `office-tools-pdf2imagepdf-*` 临时目录。

`pdf compress` 回归结果：

- 输入大小：`3407255` bytes
- `standard` 输出大小：`1882184` bytes
- 节省空间：`1525071` bytes
- 收尾验证：未残留 `PDF压缩` / WPS 产物窗口，未残留 `office-tools-pdfcompress-*` 临时目录。

`file slim` 回归结果：

- WPS 默认产物命名：`(已瘦身)<输入名>`
- 产品命令输出：`C:\Code\test\pdfreadtest\cli-product-file-slim.pdf`
- 输入大小：`3407255` bytes
- 输出大小：`1374857` bytes
- 节省空间：`2032398` bytes
- 收尾验证：未残留 `文件瘦身` / `WPS Office` 产物窗口，未残留 `office-tools-fileslim-*` 临时目录。

## WPS 本机证据

`addons\kapplist\shellext\menu_2.xml` 暴露了 PDF 工具入口：

| id | appId | pluginName | action | 格式 |
|---|---|---|---|---|
| `pdf2doc` | `kpdf2wordv2` | `kpdf2wordv3` | `ConvertToWord` | pdf |
| `pdf2xls` | `kpdf2wordv2` | `kpdf2wordv3` | `ConvertToExcel` | pdf |
| `pdf2ppt` | `kpdf2wordv2` | `kpdf2wordv3` | `ConvertToPowerPoint` | pdf |
| `pdf2imgpdf` | `kpdf2wordv2` | `kpdf2wordv3` | `ConvertToImgPDF` | pdf |
| `pdfmerge` | `kpdf2wordv2` | `kpdf2wordv3` | `Merge` | pdf |
| `pdfsplit` | `kpdf2wordv2` | `kpdf2wordv3` | `Split` | pdf |
| `pdfcompress` | `batchcompress` | `pdfbatchcompression` | | pdf |
| `batchcompress` | `kdocumentslimming` | `kbatchcompress` | | doc/docx/xls/xlsx/ppt/pptx/pdf 等 |

`addons\kpdf2wordv3\runinfo.json` 进一步说明：

- `kpdf2wordv2` 是可打开 `WPS PDF转换` 主窗口的 appframework 入口。
- `ConvertToExcel` 和 `ConvertToPowerPoint` 可以通过同一转换窗口完成。
- `kpdf2excel` / `kpdf2ppt` 是 platformApp 入口，但本机验证时只打开泛化 `WPS Office` 容器，不作为首选产品路径。

## 已探索能力

| 能力 | 证据 | 状态 | CLI 设计 |
|---|---|---:|---|
| PDF 压缩 | `batchcompress` 可打开具名 `PDF压缩` 窗口；UIA 暴露 `高品质`、`标准品质`、`中等品质`、`低品质` 和 `开始压缩`。 | 已产品化 | `office-tools pdf compress input.pdf --out output.pdf --level standard` |
| 图片 OCR | `kocrtool` / `kwebocrtool` / `Pic2Xls` 入口存在。 | 暂不产品化 | 建议保留草案 `office-tools image ocr input.png --out output.txt`，但必须先验证稳定输出路径。 |
| 图片转 Excel | `kocrtool` + `Pic2Xls` 入口存在。 | 暂不产品化 | 建议保留草案 `office-tools image to-excel input.png --out output.xlsx`，但当前窗口主要是 CEF WebView，未找到稳定控件。 |
| 图片转 PDF | 右键动词 `转换为PDF格式` 存在，`photo2pdf` appId 存在。 | 暂不产品化 | 草案 `office-tools image to-pdf input.png --out output.pdf`；本机探测中 Shell 动词调用会卡住，未观察到产物闭环。 |
| OFD 转 PDF | `kwpsofdentry` / `ofd2pdf` 入口存在，Shell 扩展命令为 WPSOFD 导出 PDF。 | 暂不产品化 | 草案 `office-tools ofd to-pdf input.ofd --out output.pdf`；本机只打开 OFD 容器，未观察到默认输出文件。 |
| PDF 转图片型 PDF | `ConvertToImgPDF` 可打开 `WPS PDF转换`，UIA 暴露 `开始转换`；默认产物为 `_扫描版.pdf`。 | 已产品化 | `office-tools pdf to-image-pdf input.pdf --out output.pdf` |
| PDF 转图片 | `kexportimage` 入口存在。 | P2 候选 | `office-tools pdf to-image input.pdf --out-dir pages --format png`，但本地库通常也可完成，需比较价值。 |
| 文件瘦身 | `kdocumentslimming` / `kbatchcompress` 可打开具名 `文件瘦身` 窗口；UIA 暴露 `另存为副本` 和 `开始瘦身`；默认产物为 `(已瘦身)<输入名>`。 | 已产品化 | `office-tools file slim input --out output` |
| PDF 拆分/合并 | `kpdf2wordv2` 的 `Split` / `Merge` 可打开 `WPS PDF转换`，UIA 暴露 `PDF拆分` / `PDF合并` 区域。 | 暂不产品化 | 本地库更适合稳定 CLI；UIA 仅保留为 WPS 产品后端备选。 |

## 不优先产品化

| 能力 | 原因 |
|---|---|
| PDF 编辑 | 交互性强，难以形成稳定无人工 CLI。 |
| PDF 签名/表单填写 | 用户意图和授权语义复杂。 |
| WPS AI 阅读/总结/生成 | 依赖账号、云端、界面变化大。 |
| PDF 合并/拆分 | 本地库通常更稳定；UIA 只作为备选后端。 |
| 批量打印 | 涉及真实打印机、副作用大，不适合默认自动化。 |

## 产品 CLI 原则

1. 命令描述用户结果，不暴露 WPS 内部 `action`、`appId`、`InstanceId`。
2. `--out` 和 `--out-dir` 是一等参数，必须严格生效。
3. UIA 命令默认清理本次任务打开的窗口和自动打开的产物页。
4. `--backend auto` 是默认值；只有确实需要时才选择 `wps-uia`。
5. 所有命令默认返回结构化 JSON，便于 Agent 和 MCP 调用。
6. 原始 WPS 启动参数只放在 `wps-uia raw ...` 下。

## 已实现命令

```bash
office-tools pdf to-word input.pdf --out output.docx
office-tools pdf to-excel input.pdf --out output.xlsx
office-tools pdf to-ppt input.pdf --out output.pptx
office-tools pdf to-image-pdf input.pdf --out output.pdf
office-tools pdf compress input.pdf --out output.pdf --level standard
office-tools file slim input.pdf --out output.pdf
```

通用参数：

| 参数 | 含义 |
|---|---|
| `--out <path>` | 精确输出文件路径。 |
| `--backend <auto|wps-uia>` | 后端选择，默认 `auto`。 |
| `--overwrite` | 覆盖已有输出。 |
| `--timeout <seconds>` | 等待转换产物的超时时间。 |
| `--cleanup <auto|always|never>` | UIA 清理策略，默认 `auto`。 |
| `--cleanup-seconds <seconds>` | 清理窗口等待时间。 |
| `--verbose` | 返回 staging、action、runner 等诊断信息。 |

`pdf compress` 额外参数：

| 参数 | 含义 |
|---|---|
| `--level <high|standard|medium|low>` | 压缩质量，默认 `standard`。值越偏向 `high` 越保留质量，越偏向 `low` 越追求小体积。 |

`file slim` 暂不暴露瘦身档位。WPS UI 中存在 `自定义设置`，但它会随文件类型变化；在 Word、Excel、PPT、PDF 的设置语义逐项验证前，产品 CLI 只承诺“另存为副本并瘦身”。

## 输出路径语义

WPS 转换器、压缩器和文件瘦身工具默认把产物写在输入文件旁边。产品 CLI 为了支持精确 `--out`，使用 staging 流程：

1. 把输入 PDF 复制到临时目录。
2. 按目标输出文件名重命名临时 PDF。
3. 启动 WPS 转换、压缩或文件瘦身，让 WPS 在临时目录生成默认产物。
4. 清理转换窗口、压缩窗口、瘦身窗口和自动打开的产物页，让文件解除锁定。
5. 把产物移动到 `--out` 指定位置。
6. 删除临时目录。

## Raw/诊断命令

```bash
office-tools wps-uia env
office-tools wps-uia verbs input.pdf
office-tools wps-uia windows
office-tools wps-uia dump-window --title "WPS PDF转换" --max-depth 8
office-tools wps-uia raw pdf-converter input.pdf --action ConvertToWord
```

这些命令用于探索和排障，不是产品主路径。

## 下一步

1. 用 `wps-uia dump-window` 继续验证图片 OCR / 图片转 Excel 的 WebView 工具是否存在稳定可调用控件。
2. 对 `ofd to-pdf` 继续寻找可直接产出 PDF 的 runner 参数；当前不建议暴露用户命令。
3. 研究 PDF 转可搜索 PDF：它与 `to-image-pdf` 同属转换窗口，但可能依赖 OCR/会员授权，必须先确认授权失败时的错误语义。
4. 如确实需要 WPS 版 `pdf split` / `pdf merge`，先定义本地库后端的产品语义，再把 UIA 作为兼容后端补上。
