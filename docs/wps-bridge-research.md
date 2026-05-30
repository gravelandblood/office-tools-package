# WPS / Office Automation Research

Date: 2026-05-30

## Executive Summary

The best direction is not a single automation backend. A useful open-source
`office-tools-package` should combine four layers:

- **OfficeCLI** for headless OOXML file operations.
- **WPS COM** for real WPS application/object-model behavior.
- **WPS JSAPI add-in bridge** for sidebar/add-in integration and active document
  context.
- **WPS UI Automation** for WPS product features that are not publicly exposed,
  especially PDF/member tooling.

The scanned PDF to Word experiment proved that WPS member PDF conversion can be
automated on a logged-in Windows desktop without JSAPI: shell verb/native WPS
entrypoint opens the conversion UI, then UI Automation invokes the exposed
`Start conversion` button.

## OfficeCLI vs WPS COM vs JSAPI vs UIA

OfficeCLI is strongest when the task is file-level, OOXML-native, headless, and
agent-friendly. It is suitable for `.docx`, `.xlsx`, and `.pptx` inspection and
modification, with commands such as get/query/set/add/remove/batch/dump/raw,
validation, rendering, and MCP serving.

WPS COM is strongest when the real WPS application matters: document opening,
saving, export, calculation, Word/Excel object-model behavior, review state,
comments, fields, and compatibility.

WPS JSAPI is strongest when the operation belongs inside a WPS add-in/sidebar or
needs the active WPS window, active selection, task pane, or web-to-WPS callback.

WPS UIA is strongest when WPS exposes a feature through UI but not through COM or
JSAPI. It requires a real Windows desktop session and is less stable than a
formal API, but it can cover important product features.

| Scenario | Preferred backend |
|---|---|
| Batch OOXML edits in CI | OfficeCLI |
| Agent JSON/MCP file tools | OfficeCLI |
| Word/Excel object model operations | WPS COM |
| Real WPS export/render/compatibility | WPS COM |
| Sidebar or add-in integration | WPS JSAPI |
| Active document/selection in WPS UI | WPS JSAPI or COM |
| Member PDF to Word | WPS UIA + shell/native entrypoint |
| WPS product feature with no public API | WPS UIA |
| Local helper launched from add-in | JSAPI `OAAssist.ShellExecute` or NativeX |

## WPS Add-In / JSAPI Official Model

WPS add-ins run inside the WPS client. JSAPI code is loaded from an add-in web
entry such as `index.html`; that in-WPS page can access `window.Application` and
WPS extension objects.

External systems do not directly execute JSAPI in a separate process. They call
into WPS through the local WPS add-in service exposed by `wpsjs-rpc-sdk-new`.

The core external call is:

```js
WpsInvoke.InvokeAsHttp(clientType, addinName, functionName, param, callback, showToFront, jsPluginsXml, silentMode)
```

Observed local service ports:

- HTTP: `58890`
- HTTPS: `58891`
- WebSocket alternative: `58892` / `58893`

For independent WPS instances, the SDK exposes:

- `new WpsClient("wps")`
- `StartWpsInSilentMode(addinName, callback)`
- `InvokeAsHttp(addinName, functionName, param, callback, showToFront)`
- `ShowToFront(addinName, callback)`
- `CloseSilentClient(addinName, callback)`

## Native Escape Hatches

`wps.OAAssist.ShellExecute(command)` is available inside the tested WPS add-in
page. It can launch local commands from add-in JavaScript.

NativeX is the heavier extension path. WPS documentation describes it as a way
to extend JSAPI with native modules such as C++/Ruby/Python-backed DLL/SO code,
registered through `ksomisc`, then called from JSAPI through `CreateObject`.

## Local Prototype

Created:

- `codex-bridge-addin/`
- `codex-bridge-tools/`

Add-in name:

- `CodexBridge`

Implemented add-in functions:

- `CodexBridge_Ping(param)`
- `CodexBridge_GetActiveDocument(param)`
- `CodexBridge_OpenDocument(param)`
- `CodexBridge_LaunchPdfConvert(param)`

Validation page:

- `http://127.0.0.1:48990/invoke.generated.html`

Add-in service:

- `http://127.0.0.1:3889/`

Validation XML:

- `http://127.0.0.1:48990/jsplugins.xml`

## Verified: JSAPI Bridge

1. WPS local RPC service can be started and reached at `127.0.0.1:58890`.
2. `WpsAddonMgr.verifyStatus` returns OK for our `ribbon.xml`.
3. `WpsAddonMgr.enable` returns OK and writes `publish.xml`.
4. Direct `WpsInvoke.InvokeAsHttp` against already-running WPS returned
   `No Plugin named: CodexBridge` until the target WPS process loaded the add-in.
5. `WpsClient.StartWpsInSilentMode` with `jsPluginsXml` successfully loaded the
   add-in in an independent WPS client.
6. `CodexBridge_Ping` returned:
   - `hasWindowApplication: true`
   - `hasGlobalWps: true`
   - `hasOAAssist: true`
   - `hasShellExecute: true`
   - `hasWebNotify: true`
7. `CodexBridge_GetActiveDocument` returned successfully.
8. `CodexBridge_LaunchPdfConvert` called `wps.OAAssist.ShellExecute` and returned
   `launched: true` for the WPS PDF conversion runner command.

## Verified: WPS Member PDF to Word Through UIA

The real WPS right-click shell verb for "convert to Word" (`转为Word`) launches:

```text
wpscloudsvr.exe Run /InstanceId:WpsCloudSvr ... /app_id=kpdf2wordv2 /app_params=<base64> /src=public_rclickmenu
```

Then it launches:

```text
wps.exe Run /InstanceId=kpdf2wordv2 ... /action=ConvertToWord /file=<pdf>
```

The WPS PDF conversion window title is `WPS PDF转换`.

UI Automation exposes useful controls, including:

- navigation items such as `转为Word`, `转为Excel`, `转为PPT`
- output settings
- a button containing the text `开始转换`
- completion dialog `操作完成`

UIA can invoke the `开始转换` button. The scanned PDF converted successfully to:

```text
C:\Code\test\pdfreadtest\工商底档-2023.10.26变更登记.docx
```

The output `.docx` contains OCR text, not only images. Extracted text begins with
the expected registration notice content:

```text
登记通知书
(武新市监)登字〔2023〕第200758号
武汉国科光领半导体科技有限公司：
```

Text extraction found roughly 290k text characters.

Reusable scripts:

- `codex-bridge-tools/ConvertPdfToWord.ps1`
- `codex-bridge-tools/DumpWpsPdfUi.ps1`
- `codex-bridge-tools/ClickStartConvert.ps1`
- `codex-bridge-tools/ExtractDocxText.ps1`
- `codex-bridge-tools/watch-verb.ps1`

## Practical Architecture

Recommended bridge stack:

```text
Codex / appserver / MCP
  -> tool router
  -> OfficeCLI backend for OOXML files
  -> WPS COM backend for application object model
  -> WPS JSAPI backend for add-in/sidebar context
  -> WPS UIA backend for product-only features
  -> WPS shell/native entrypoints
```

Use OfficeCLI for:

- headless `.docx/.xlsx/.pptx` operations
- CI and server-side batch processing
- structured JSON output
- MCP-friendly file tools

Use WPS COM for:

- real WPS open/save/export
- Word/Excel object model operations
- fields, review state, comments, calculations, and compatibility behavior

Use WPS JSAPI for:

- active document/window info
- selection/range reads and edits
- task pane UI and sidebar workflows
- add-in to external web notification through `WebNotify`

Use WPS UIA for:

- WPS PDF conversion/member tools
- product UI features not exposed by public APIs
- workflows where the user is already logged in and authorized

Use `OAAssist.ShellExecute` or NativeX only when needed:

- invoking existing WPS native feature entrypoints
- calling a controlled local helper
- accessing Windows-only APIs that JSAPI does not expose

## Open Questions

- Whether the `wpscloudsvr.exe /app_id=kpdf2wordv2 /app_params=<base64>` route
  accepts more parameters, such as output directory or auto-start.
- Whether UIA workflows remain stable across WPS versions and themes.
- Whether installed, visible WPS windows must be restarted to pick up newly
  enabled add-ins; `WpsClient` avoids this by starting a separate client.
- How much of WPS review/comment/accept-change behavior is better handled by COM
  versus JSAPI in practice.
