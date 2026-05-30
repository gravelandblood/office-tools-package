# office-tools-package Blueprint

Date: 2026-05-30

This blueprint turns the WPS bridge/UIA experiments into an open-source package
direction.

## Goal

`office-tools-package` should provide one agent-friendly tool layer over several
document automation engines:

- file-level OOXML tools, such as OfficeCLI
- WPS desktop automation through COM
- WPS add-in bridge through JSAPI/RPC
- WPS UI Automation for features that are not publicly exposed
- optional OCR/vision tools for scanned documents

The package should make these engines look like one CLI/MCP surface while still
reporting which backend was used.

## Backend Model

```text
office-tools-package
  -> officecli backend: headless OOXML operations
  -> wps-com backend: WPS application object model
  -> wps-jsapi backend: WPS add-in RPC bridge
  -> wps-uia backend: visible desktop automation
  -> native/shell backend: shell verbs and WPS product entrypoints
```

## Selection Rules

Use OfficeCLI by default when:

- the input is `.docx`, `.xlsx`, or `.pptx`
- the task is file-level and headless
- deterministic JSON output and CI usage matter
- the operation is structure-oriented: get/query/set/add/remove/merge/validate

Use WPS COM when:

- WPS must open, calculate, render, or save the file
- Word/Excel object-model fidelity matters
- review, comments, revisions, fields, printing, or export behavior depends on
  the WPS application

Use WPS JSAPI when:

- the operation belongs inside a WPS add-in/sidebar
- the active document/window/selection is important
- web-to-WPS callbacks or task pane UX are required

Use WPS UIA when:

- WPS exposes the feature only through product UI
- member-only or PDF tooling is needed
- the function can be safely represented as a click/choose/wait workflow

Use shell/native entrypoints when:

- WPS registers a shell verb or known runner command
- the runner can open the exact product workflow needed

## Verified Prototype: WPS Scanned PDF to Word

The experiment validated a Windows desktop workflow:

```text
PDF file
  -> Shell.Application verb for "convert to Word" (`转为Word`)
  -> wpscloudsvr.exe /app_id=kpdf2wordv2 /app_params=<base64>
  -> wps.exe Run /InstanceId=kpdf2wordv2 /action=ConvertToWord /file=<pdf>
  -> WPS PDF conversion window (`WPS PDF转换`)
  -> UI Automation InvokePattern on the start button (`开始转换`)
  -> .docx output in the source PDF directory
```

Observed output:

- Source PDF: `C:\Code\test\pdfreadtest\工商底档-2023.10.26变更登记.pdf`
- Output DOCX: `C:\Code\test\pdfreadtest\工商底档-2023.10.26变更登记.docx`
- The output contains OCR text, not only images.

Prototype command:

```powershell
powershell -ExecutionPolicy Bypass -File .\codex-bridge-tools\ConvertPdfToWord.ps1 `
  -InputPdf "C:\Code\test\pdfreadtest\工商底档-2023.10.26变更登记.pdf"
```

## Proposed CLI Shape

```bash
office-tools pdf to-word input.pdf --backend wps-uia
office-tools docx text input.docx --backend officecli
office-tools docx replace input.docx --find A --replace B --backend officecli
office-tools wps active-document --backend jsapi
office-tools wps export-pdf input.docx --backend com
office-tools mcp serve
```

Every command should return structured JSON by default:

```json
{
  "ok": true,
  "backend": "wps-uia",
  "input": "input.pdf",
  "output": "input.docx",
  "events": []
}
```

## Package Layout Proposal

```text
office-tools-package/
  README.md
  docs/
    architecture.md
    backend-matrix.md
    wps-pdf-to-word-uia.md
  packages/
    cli/
    mcp-server/
    backend-officecli/
    backend-wps-com/
    backend-wps-jsapi/
    backend-wps-uia/
  examples/
    pdf-to-word-wps-uia.ps1
    wps-addin-bridge/
```

## Open-Source Boundaries

- Do not bypass WPS membership or license gates.
- UIA workflows should automate the user's already-authorized desktop session.
- The project should clearly mark WPS UIA/COM as Windows desktop backends.
- Keep OfficeCLI integration optional and adapter-based, because it is a separate
  open-source project.
- Prefer explicit user-visible actions for paid/cloud features.

## Next Engineering Step

Turn the current PowerShell prototypes into a small `backend-wps-uia` package:

- `convertPdfToWord(inputPdf, options)`
- `findWindow(title)`
- `dumpTree(title)`
- `invokeButton(windowTitle, buttonText)`
- `waitForFile(path, timeout)`

Then wrap it with a Node CLI and return JSON.
