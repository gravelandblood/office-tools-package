# office-tools-package

一个面向 Agent、CLI 和 MCP 的 Office/WPS 自动化工具包实验项目。

这个项目的目标不是只做一个脚本，而是把多种文档自动化路线统一成一个可组合的工具层：

- **OfficeCLI**：用于无头、确定性的 OOXML 文件级读写。
- **WPS JSAPI**：用于 WPS 插件/侧边栏、活动文档、选区和窗口上下文。
- **WPS COM**：用于真实 WPS 应用对象模型、打开/保存/导出等行为。
- **WPS UIA**：用于 WPS 只暴露在桌面界面里的能力，例如会员 PDF 转换工具。

当前最先落地的是 WPS UI Automation 后端的 **PDF 转 Word** 能力。

## 当前状态

这是一个早期开放源码原型，已经包含：

- npm workspace 包结构
- `office-tools` CLI 原型
- WPS UIA PDF 转 Word 后端
- OfficeCLI / WPS JSAPI / WPS UIA 能力规划
- WPS 插件桥接和 RPC 调用原型
- Windows 桌面 WPS 自动化诊断脚本

## 安装

```powershell
npm install
```

查看能力矩阵：

```powershell
node packages/cli/bin/office-tools.js capabilities
```

## PDF 转 Word

使用 WPS UIA 后端转换 PDF：

```powershell
node packages/cli/bin/office-tools.js pdf to-word C:\path\input.pdf --backend wps-uia
```

常用参数：

```powershell
node packages/cli/bin/office-tools.js pdf to-word C:\path\input.pdf `
  --output C:\path\output.docx `
  --launch-mode shell `
  --timeout 240
```

默认情况下，转换完成后会清理本次任务打开的 WPS 窗口、转换窗口和自动打开的文档页，避免影响下一次自动化任务。调试 WPS UI 时可以使用：

```powershell
node packages/cli/bin/office-tools.js pdf to-word C:\path\input.pdf --no-cleanup
```

## 启动模式

WPS UIA PDF 转 Word 后端支持三种启动模式：

- `shell`：调用 Windows 中注册的 WPS PDF 转 Word 右键菜单/ Shell verb。
- `native`：调用已观测到的 `wps.exe Run /InstanceId=kpdf2wordv2 ...` 入口。
- `cloud`：调用已观测到的 `wpscloudsvr.exe /app_id=kpdf2wordv2 ...` 入口。

`cloud` 模式中的 `app_params` 结构仍需要按 WPS 版本继续逆向/观测，因此目前通过 `--cloud-app-params` 和 `--cloud-arg` 透传。

## 目录结构

```text
office-tools-package/
  docs/
    backend-capability-plan.md
    office-tools-package-blueprint.md
    wps-bridge-research.md
  packages/
    cli/
    backend-officecli/
    backend-wps-jsapi/
    backend-wps-uia/
    backend-wps-com/
    mcp-server/
  examples/
    wps-addin-bridge/
    wps-rpc-invoke-prototype/
```

## 后端选择原则

优先使用确定性、无头的文件后端；只有在能力确实只存在于桌面产品 UI 中时，才使用 UIA。

| 场景 | 推荐后端 |
|---|---|
| `.docx/.xlsx/.pptx` 文件级读写 | OfficeCLI |
| CI / 批处理 / MCP 结构化工具 | OfficeCLI |
| WPS 活动文档、选区、插件侧边栏 | WPS JSAPI |
| WPS 打开、保存、导出、对象模型行为 | WPS COM |
| WPS PDF 会员工具、UI-only 功能 | WPS UIA |

详细规划见 [docs/backend-capability-plan.md](docs/backend-capability-plan.md)。

## 开源边界

- 本项目不会绕过 WPS 会员或授权限制。
- UIA 自动化只面向用户已经登录、已经授权的本地桌面会话。
- WPS UIA/COM 后端仅适用于 Windows 桌面环境。
- OfficeCLI 作为独立后端适配，不与本项目强绑定。

## 许可证

MIT
