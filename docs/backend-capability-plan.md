# Backend Capability Plan

Date: 2026-05-30

This document is the implementation map for turning the current WPS and
OfficeCLI research into a stable package surface.

References checked while drafting this plan:

- OfficeCLI GitHub README: confirms the read/query/set/add/remove/raw,
  rendering, batch, resident, validation, and MCP-oriented command model.
- WPS Open Platform `Application` and `Document` object docs: confirms
  active-document/application-object operations such as `ActiveDocument`,
  `Documents`, activation, close, compare, revisions, print, and application
  metadata.
- `wpsjs-rpc-sdk-new` package docs and local prototype: confirms the external
  bridge pattern around `InvokeAsHttp` and `StartWpsInSilentMode`.

## Backend Summary

| Backend | Best for | Avoid for |
|---|---|---|
| OfficeCLI | Headless OOXML inspection and mutation, deterministic JSON, CI and MCP tools | Product-only WPS features, active desktop state, WPS membership workflows, PDF page operations |
| WPS JSAPI | Add-in/sidebar flows, active document/window/selection, web-to-WPS callbacks | Headless batch jobs, UI-only product features, arbitrary desktop automation |
| WPS COM | Real WPS application behavior, open/save/export, object-model fidelity | Non-Windows environments, unattended server usage without a desktop profile |
| WPS UIA | WPS features exposed only through desktop UI, especially PDF/member tools | Deterministic OOXML edits, CI, hidden/headless automation |

## OfficeCLI Capability Surface

OfficeCLI should be the default file backend for `.docx`, `.xlsx`, and `.pptx`
when the job can be expressed as OOXML operations.

Template note:

- Template Detective custom rendering should default to OfficeCLI-style OOXML
  patching, not COM/JSAPI/UIA.
- Office-native template compilation means wrapping existing DOCX ranges with
  content controls, adding custom XML parts/bindings, and preserving existing
  styles/numbering/table properties.
- `template compile-office` implements that route for safe scalar slots and
  uniquely located table loops, wrapping table body rows with Word repeating
  section controls while preserving the original table properties.
- Ambiguous compile targets should be returned as scored candidates. A human
  or LLM can review them and pass an accepted-candidates JSON file back to
  `template compile-office --accept-candidates` for fast trial rendering.
- `template render-office` completes the deterministic loop for compiled
  Office-native templates: scalar content controls are replaced from JSON,
  repeating section table rows are cloned from array data, and custom XML is
  updated in the package.
- Sidecar rules for conditions and complex loops should be applied by a
  deterministic OOXML patcher, then verified with `template compare-format`.
- Use WPS/Word COM or JSAPI only when real application behavior is required:
  field updates, pagination-sensitive output, active document interactions, or
  export/render verification.

Planned command groups:

| Area | Capabilities | Package command shape |
|---|---|---|
| Inspect | metadata, document structure, text, styles, comments, relationships, raw XML | `office-tools docx inspect`, `office-tools xlsx inspect`, `office-tools pptx inspect` |
| Query | JSON/path-like reads from document parts and semantic objects | `office-tools docx query input.docx --path ...` |
| Mutate | set/add/remove/replace paragraphs, runs, cells, slides, shapes, relationships | `office-tools docx replace`, `office-tools xlsx set`, `office-tools pptx add-shape` |
| Batch | apply a declarative JSON patch plan to files | `office-tools ooxml apply plan.json` |
| Validate | package integrity, relationship checks, schema-adjacent sanity checks | `office-tools ooxml validate input.docx` |
| Extract | text, tables, images, embedded media, notes | `office-tools docx text`, `office-tools pptx images` |
| Render bridge | delegate render/export to a configured app backend when OfficeCLI cannot render | router selects WPS COM or another renderer |
| MCP | expose the same deterministic operations to agents | `office-tools mcp serve` |

PDF note:

- Current OfficeCLI public capability descriptions are centered on Word,
  Excel, and PowerPoint file generation/inspection/editing, plus template and
  patch-style OOXML workflows.
- Do not route PDF split/merge/rotate through the OfficeCLI adapter unless
  OfficeCLI explicitly adds those commands later.
- For deterministic PDF page operations, prefer a dedicated PDF backend and
  keep WPS UIA only as a product-compatible fallback when WPS behavior is
  specifically required.

Implementation notes:

- Treat OfficeCLI as an optional adapter dependency, not vendored code.
- Normalize all OfficeCLI results into `{ ok, backend: "officecli", command,
  input, output?, data?, events }`.
- Probe availability with `officecli --version` and return actionable install
  diagnostics when missing.
- Prefer OfficeCLI before WPS UIA for any file-level OOXML task.

## WPS JSAPI Capability Surface

WPS JSAPI belongs inside a loaded WPS add-in and is reached externally through
the local WPS add-in service/RPC bridge.

Planned command groups:

| Area | Capabilities | Package command shape |
|---|---|---|
| Health | ping add-in, WPS version, loaded bridge status, OAAssist availability | `office-tools wps-jsapi ping` |
| Active context | active document, window, selection, cursor/range info | `office-tools wps active-document`, `office-tools wps selection` |
| Open/save | open document from path/URL, save, save-as, close | `office-tools wps open`, `office-tools wps save-as` |
| Document edits | selection/range text insertion, find/replace, comments, fields where exposed | `office-tools wps insert-text`, `office-tools wps replace` |
| Spreadsheet context | active workbook, worksheet, range values, selected range updates | `office-tools et active-workbook`, `office-tools et range` |
| Presentation context | active presentation, slides, shapes, selection edits | `office-tools wpp slides`, `office-tools wpp shapes` |
| Notifications | `WebNotify` events from add-in to host process | future event bridge |
| Native launch | controlled `OAAssist.ShellExecute` entrypoints such as WPS PDF converter | `office-tools wps launch` |

Implementation notes:

- Keep the add-in small: it should expose named bridge functions, not business
  logic.
- Use `WpsInvoke.InvokeAsHttp` for an already-running target client and
  `WpsClient.StartWpsInSilentMode` when isolation is more important than the
  visible active session.
- Return structured JSON from every bridge function.
- Explicitly distinguish active-session commands from silent-client commands.

## WPS UIA-Only Capability Surface

Use UIA only when there is no stable file-level, COM, or JSAPI API.

Initial UIA-only candidates:

| Capability | Why UIA | First implementation |
|---|---|---|
| PDF to Word through WPS member converter | Exposed as WPS product PDF UI and shell/native runner, not a documented JSAPI | Implemented as `office-tools pdf to-word --backend wps-uia` |
| PDF to Excel/PPT variants | Same PDF conversion window exposes conversion modes | Implemented as `office-tools pdf to-excel` and `office-tools pdf to-ppt` |
| PDF compression | WPS exposes a dedicated `PDF压缩` desktop tool with quality choices | Implemented as `office-tools pdf compress --out output.pdf --level standard` |
| PDF to image-PDF variant | WPS conversion window exposes `ConvertToImgPDF` and generates `_扫描版.pdf` | Implemented as `office-tools pdf to-image-pdf` |
| File slimming | WPS exposes a dedicated `文件瘦身` desktop tool with stable start/copy controls | Implemented as `office-tools file slim --out output` |
| PDF to image variants | Shell menu exposes entries, but output behavior still needs verification | Keep as research until action and output names are verified |
| Scanned PDF OCR conversion | Product converter performs OCR after user authorization | Extend PDF conversion backend with OCR diagnostics |
| PDF merge/split/encrypt tools | WPS PDF toolbox UI, member gated in many installs | Add per-flow UI tree snapshots before automation |
| UI-only export/settings dialogs | Some product settings are not surfaced through COM/JSAPI | Add `dump`, `find-window`, and dry-run diagnostics first |

UIA requirements:

- Windows desktop session must be visible/unlocked.
- The user must already be authorized for WPS paid/cloud features.
- Automation must not bypass license gates.
- Every UIA command needs diagnostics: window title, control text, WPS version,
  launch command, timeout, and UI dump support.

## PDF to Word Implementation Plan

Current command:

```bash
office-tools pdf to-word input.pdf --out output.docx --backend wps-uia
```

Implemented options:

| Option | Meaning |
|---|---|
| `--out <path>` | Exact output DOCX path |
| `--timeout <seconds>` | Wait for conversion output |
| `--cleanup <auto|always|never>` | Cleanup policy for WPS windows |
| `--overwrite` | Replace existing target output |
| `--verbose` | Include diagnostics such as staging paths |

Internal launcher options live under `office-tools wps-uia raw pdf-converter`.

Implemented follow-up:

- `pdf to-excel` and `pdf to-ppt` reuse the PDF conversion window with verified
  output extensions and cleanup.
- `pdf to-image-pdf` reuses the PDF conversion window with verified `_扫描版.pdf`
  output naming and cleanup.
- `pdf compress` uses the WPS `batchcompress` app, supports exact `--out`, and
  exposes only product-level quality values: `high`, `standard`, `medium`,
  and `low`.
- `file slim` uses the WPS `kdocumentslimming` app, supports exact `--out`,
  keeps the source safe by selecting copy output, and waits for WPS'
  `(已瘦身)<name>` default artifact.

Parameter discovery still needed:

- Observe WPS process command lines for PDF to image, text, and OCR flows.
- Check whether output directory, page range, language, OCR mode, and output
  format can be passed in app params or runner params.
- Add UIA setting controls only after the corresponding stable control names are
  dumped and versioned.
- For `file slim`, expose custom slimming settings only after Word, Excel,
  PowerPoint, and PDF setting panels are separately verified.

Milestones:

1. Stabilize `pdf to-word` with JSON output and failure codes.
2. Add `office-tools wps-uia dump --title ...` and `watch-verb` diagnostics.
3. Record WPS version/window/control snapshots for supported versions.
4. Add conversion variants only after command-line actions and output behavior
   are verified.
5. Expose the backend through MCP once CLI behavior is stable.
