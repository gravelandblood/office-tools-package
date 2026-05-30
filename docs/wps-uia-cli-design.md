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

验证样本：

- 输入：`C:\Code\test\pdfreadtest\工商底档-2023.10.26变更登记.pdf`
- 输出：
  - `C:\Code\test\pdfreadtest\cli-product-pdf2excel.xlsx`
  - `C:\Code\test\pdfreadtest\cli-product-pdf2ppt.pptx`
  - `C:\Code\test\pdfreadtest\cli-product-regression-word.docx`

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

`addons\kpdf2wordv3\runinfo.json` 进一步说明：

- `kpdf2wordv2` 是可打开 `WPS PDF转换` 主窗口的 appframework 入口。
- `ConvertToExcel` 和 `ConvertToPowerPoint` 可以通过同一转换窗口完成。
- `kpdf2excel` / `kpdf2ppt` 是 platformApp 入口，但本机验证时只打开泛化 `WPS Office` 容器，不作为首选产品路径。

## 可继续产品化的能力

| 能力 | 证据 | 建议优先级 | CLI 草案 |
|---|---|---:|---|
| PDF 压缩 | 可打开具名 `PDF压缩` 窗口 | P1 | `office-tools pdf compress input.pdf --out output.pdf --level medium` |
| 图片 OCR | `kocrtool` / `kwebocrtool` / `Pic2Xls` 入口存在 | P1 | `office-tools image ocr input.png --out output.txt` |
| 图片转 Excel | `kocrtool` + `Pic2Xls` 入口存在 | P1 | `office-tools image to-excel input.png --out output.xlsx` |
| OFD 转 PDF | `kwpsofdentry` / `ofd2pdf` 入口存在 | P2 | `office-tools ofd to-pdf input.ofd --out output.pdf` |
| PDF 转图片 | `kexportimage` 入口存在 | P2 | `office-tools pdf to-image input.pdf --out-dir pages --format png` |
| 文件瘦身 | 可打开具名 `文件瘦身` 窗口 | P2 | `office-tools file slim input.pdf --out output.pdf --level medium` |

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

## 输出路径语义

WPS 转换器默认把产物写在输入文件旁边。产品 CLI 为了支持精确 `--out`，使用 staging 流程：

1. 把输入 PDF 复制到临时目录。
2. 按目标输出文件名重命名临时 PDF。
3. 启动 WPS 转换，让 WPS 在临时目录生成默认产物。
4. 清理转换窗口和自动打开的产物页，让文件解除锁定。
5. 把产物移动到 `--out` 指定位置。
6. 删除临时目录。

## Raw/诊断命令

```bash
office-tools wps-uia env
office-tools wps-uia verbs input.pdf
office-tools wps-uia windows
office-tools wps-uia raw pdf-converter input.pdf --action ConvertToWord
```

这些命令用于探索和排障，不是产品主路径。

## 下一步

1. 把 `pdf compress` 做成可产物闭环的 P1 命令。
2. 找图片 OCR / 图片转 Excel 的稳定启动和输出路径。
3. 收集 OFD 样本，验证 `ofd to-pdf`。
4. 为 UIA 加窗口 dump / 控件快照能力，降低后续版本漂移风险。
