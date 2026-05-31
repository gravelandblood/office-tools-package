# office-tools-package

面向 Agent、CLI 和 MCP 的 Office/WPS 自动化工具包实验项目。

这个项目的目标不是做一堆零散脚本，而是把多条办公自动化路线整理成统一、可组合、可验证的工具层：

- **OfficeCLI**：用于无头、确定性的 OOXML 文件级读写。
- **WPS JSAPI**：用于 WPS 插件、侧边栏、活动文档、选区和窗口上下文。
- **WPS COM**：用于真实 WPS 应用对象模型、打开、保存、导出等行为。
- **WPS UIA**：用于只能通过 WPS 桌面界面使用的产品能力，例如 PDF 转 Office、OCR、压缩等。

当前已经落地的是 WPS UI Automation 后端里的 PDF 转 Office、图片型 PDF、PDF 压缩与文件瘦身能力。

## 当前状态

- npm workspace 包结构
- `office-tools` CLI 原型
- WPS UIA PDF 转 Word / Excel / PPT / 图片型 PDF / 压缩 / 文件瘦身
- OfficeCLI / WPS JSAPI / WPS UIA 能力规划
- WPS 插件桥接和 RPC 调用原型
- Windows 桌面 WPS 自动化诊断脚本

## 安装

```powershell
npm install
```

查看能力矩阵：

```powershell
npm run cli -- capabilities
```

## PDF 转 Office

使用 WPS UIA 后端转换 PDF：

```powershell
node packages/cli/bin/office-tools.js pdf to-word C:\path\input.pdf --out C:\path\output.docx
node packages/cli/bin/office-tools.js pdf to-excel C:\path\input.pdf --out C:\path\output.xlsx
node packages/cli/bin/office-tools.js pdf to-ppt C:\path\input.pdf --out C:\path\output.pptx
node packages/cli/bin/office-tools.js pdf to-image-pdf C:\path\input.pdf --out C:\path\output.pdf
```

常用参数：

```powershell
node packages/cli/bin/office-tools.js pdf to-word C:\path\input.pdf `
  --out C:\path\output.docx `
  --timeout 240 `
  --overwrite
```

默认情况下，转换完成后会清理本次任务打开的 WPS 转换窗口和自动打开的文档页，避免影响下一次自动化任务。

## PDF 压缩

使用 WPS UIA 后端压缩 PDF，并把输出文件放到指定路径：

```powershell
node packages/cli/bin/office-tools.js pdf compress C:\path\input.pdf --out C:\path\compressed.pdf
```

压缩质量可选：

```powershell
node packages/cli/bin/office-tools.js pdf compress C:\path\input.pdf `
  --out C:\path\compressed.pdf `
  --level standard `
  --overwrite
```

`--level` 支持 `high`、`standard`、`medium`、`low`。默认使用 `standard`。命令会在临时目录中运行 WPS 压缩器，再把产物移动到 `--out` 指定位置，并清理本次任务打开的 `PDF压缩` 窗口。

## 文件瘦身

使用 WPS 的文件瘦身工具减小 Office/PDF 文件体积：

```powershell
node packages/cli/bin/office-tools.js file slim C:\path\input.pdf --out C:\path\slimmed.pdf --overwrite
```

命令会在临时目录中运行 WPS 文件瘦身，等待默认产物 `(已瘦身)<文件名>`，再移动到 `--out` 指定位置，并清理本次任务打开的 `文件瘦身` 和 `WPS Office` 窗口。

## 验证与打包

静态验证：

```powershell
npm run verify
```

本机 WPS UIA 回归需要 Windows、已安装 WPS，并提供一个本地 PDF 样本：

```powershell
npm run smoke:wps-uia -- --sample C:\path\input.pdf
```

只跑部分命令：

```powershell
npm run smoke:wps-uia -- --sample C:\path\input.pdf --commands to-image-pdf,file-slim
```

检查 CLI 包发布内容：

```powershell
npm run pack:cli
```

当前是 npm workspaces 多包结构，`@office-tools/cli` 依赖同 workspace 下的后端包；正式发布时需要按后端包到 CLI 包的顺序一起发布，或后续再做单包 bundle。

调试 WPS UI 时可以使用：

```powershell
node packages/cli/bin/office-tools.js wps-uia raw pdf-converter C:\path\input.pdf --no-cleanup
```

WPS UIA 探索/诊断命令：

```powershell
node packages/cli/bin/office-tools.js wps-uia env
node packages/cli/bin/office-tools.js wps-uia verbs C:\path\input.pdf
node packages/cli/bin/office-tools.js wps-uia windows
```

## 后端选择原则

优先使用确定性、无头的文件后端；只有能力确实存在于桌面产品 UI 中时，才使用 UIA。

| 场景 | 推荐后端 |
|---|---|
| `.docx/.xlsx/.pptx` 文件级读写 | OfficeCLI |
| PDF 合并、拆分、旋转等页面级处理 | 待接入成熟 PDF 后端 |
| CI / 批处理 / MCP 结构化工具 | OfficeCLI |
| WPS 活动文档、选区、插件侧边栏 | WPS JSAPI |
| WPS 打开、保存、导出、对象模型行为 | WPS COM |
| WPS PDF 会员工具、UI-only 能力 | WPS UIA |

## WPS UIA 方向

已经验证并产品化：

- `pdf to-word`
- `pdf to-excel`
- `pdf to-ppt`
- `pdf to-image-pdf`
- `pdf compress`
- `file slim`

已经探索但暂不产品化：

- `image ocr` / `image to-excel`：入口存在，但当前 UI 主要是 WebView，未观察到稳定无人工输出闭环。
- `image to-pdf`：右键动词存在，但 Shell 动词调用在本机探测中会卡住，暂不包装成 CLI。
- `ofd to-pdf`：入口存在，但本机验证只打开 OFD 容器，未观察到默认输出文件。
- `pdf split` / `pdf merge`：入口存在且可打开 `WPS PDF转换` 窗口，但本地库更适合做稳定 CLI；UIA 暂作为备选研究路线。

新增诊断能力：

- `wps-uia dump-window --title <text>`：输出 UIA 窗口控件树，用于判断新 WPS 工具是否有稳定控件和可自动化闭环。

不优先做成稳定 CLI：

- PDF 编辑、签名、表单填写等强交互能力
- WPS AI 阅读/总结/生成类能力
- 本地库能稳定完成的普通 PDF 合并、拆分、旋转

## 开源边界

- 本项目不绕过 WPS 会员或授权限制。
- UIA 自动化只面向用户已经登录、已经授权的本地桌面会话。
- WPS UIA/COM 后端仅适用于 Windows 桌面环境。
- OfficeCLI 作为独立适配后端，不与本项目强绑定。

## 文档

- [后端能力规划](docs/backend-capability-plan.md)
- [WPS UIA 能力与 CLI 设计](docs/wps-uia-cli-design.md)

## 许可证

MIT
