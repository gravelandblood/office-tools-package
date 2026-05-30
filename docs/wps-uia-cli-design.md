# WPS UIA Capability and CLI Design

Date: 2026-05-31

This document rethinks the WPS UIA backend as a product surface instead of a
collection of launcher parameters.

## Why WPS UIA Exists

WPS UIA should be used only for product capabilities that are valuable but not
available through stable file APIs, COM, or JSAPI. It is a desktop automation
backend, not the default document engine.

Good UIA candidates have these traits:

- WPS exposes the feature through visible desktop UI.
- The feature depends on WPS account/member/cloud/product behavior.
- The feature is hard to reproduce with deterministic local libraries.
- The workflow can be represented as launch, configure, start, wait, collect,
  and clean up.

Poor UIA candidates:

- Plain OOXML reads/writes.
- Simple PDF merge/split/compress that can be done deterministically by local
  libraries.
- Workflows that require many ambiguous manual choices.
- Anything that tries to bypass WPS authorization.

## Current Local Observations

Observed PDF shell verbs include:

- `转换为Word`
- `转为Word`
- `PDF转office`
- `PDF编辑`
- `使用WPS PDF编辑`
- `WPS AI`

Observed WPS executables relevant to automation include:

- `wps.exe`, `et.exe`, `wpp.exe`
- `wpspdf.exe`
- `wpscloudsvr.exe`
- `wpscloudlaunch.exe`
- `wpsofd.exe`

Official WPS PDF surfaces also point to a broad PDF tool family: conversion,
OCR, compression, merge/split, editing, signing, watermarks, page organization,
protection/unlock, forms, batch PDF, and PDF AI.

## UIA Capability Matrix

| Capability | User Command | UIA Value | Priority | Notes |
|---|---|---:|---:|---|
| PDF to Word | `pdf to-word` | High | P0 | Already validated, including OCR output for scanned PDFs. |
| PDF to Excel | `pdf to-excel` | High | P1 | Same WPS PDF converter family; verify action/window/output extension. |
| PDF to PPT | `pdf to-ppt` | High | P1 | Same converter family; verify fidelity and output extension. |
| PDF to Image | `pdf to-image` | Medium | P1 | Likely WPS UI capability; local renderers may be better for deterministic image export. |
| Scanned PDF OCR | `pdf ocr` | High | P1 | Valuable when WPS OCR/member engine outperforms local OCR setup. |
| PDF to Text | `pdf to-text` | Medium | P2 | Prefer local extraction first; UIA only for WPS OCR output. |
| PDF compress | `pdf compress` | Medium | P2 | Useful if WPS compression presets are desired; otherwise use local PDF libs. |
| PDF merge | `pdf merge` | Low/Medium | P2 | Prefer deterministic local merge; UIA only for WPS-specific batch UI. |
| PDF split/extract pages | `pdf split` / `pdf extract-pages` | Low/Medium | P2 | Prefer local libs unless WPS UI has a needed feature. |
| PDF edit text/image | `pdf edit` | Low | P3 | Too interactive; better as diagnostics/prototype, not first stable CLI. |
| PDF watermark | `pdf watermark` | Medium | P3 | Could be scripted if UI controls are stable; local libs may be better. |
| PDF sign/fill form | `pdf sign` / `pdf fill` | Low | P3 | High ambiguity and user-auth semantics. |
| PDF protect/unlock | `pdf protect` / `pdf unlock` | Low | P3 | Security-sensitive; avoid until scope is explicit. |
| OFD conversion | `ofd ...` | Medium | P3 | `wpsofd.exe` suggests a possible route; needs separate research. |
| WPS AI on document | `wps ai ...` | Experimental | P4 | Product/account/UI churn likely high. |

## CLI Design Principles

1. User commands describe outcomes, not WPS internals.
2. `--out` and `--out-dir` are first-class; output location must be honored.
3. Internal WPS launcher details are hidden behind `--debug` or a raw command.
4. `--backend auto` is the default; `wps-uia` is explicit or selected only when
   no safer backend exists.
5. Every command returns JSON by default once used by agents/MCP.
6. UIA commands always clean up by default.
7. UIA commands expose diagnostics separately from the normal user path.

## Proposed Command Shape

Top-level:

```bash
office-tools capabilities
office-tools doctor
office-tools pdf <action> ...
office-tools docx <action> ...
office-tools xlsx <action> ...
office-tools pptx <action> ...
office-tools wps-uia <diagnostic> ...
```

PDF conversion:

```bash
office-tools pdf to-word input.pdf --out output.docx
office-tools pdf to-excel input.pdf --out output.xlsx
office-tools pdf to-ppt input.pdf --out output.pptx
office-tools pdf to-image input.pdf --out-dir pages --format png
office-tools pdf ocr input.pdf --out output.docx --format docx
```

PDF operations:

```bash
office-tools pdf compress input.pdf --out compressed.pdf --level medium
office-tools pdf merge a.pdf b.pdf --out merged.pdf
office-tools pdf split input.pdf --out-dir pages --pages 1-3,8
```

Diagnostics:

```bash
office-tools wps-uia env
office-tools wps-uia verbs input.pdf
office-tools wps-uia windows
office-tools wps-uia dump --title "WPS PDF转换"
office-tools wps-uia watch-launch input.pdf --verb "转为Word"
```

Raw/experimental escape hatch:

```bash
office-tools wps-uia raw pdf-converter input.pdf \
  --action ConvertToWord \
  --runner-param key=value
```

The raw command is intentionally not the main `pdf to-word` interface.

## Common Options

Common user-facing options:

| Option | Meaning |
|---|---|
| `--out <path>` | Exact output file path. |
| `--out-dir <dir>` | Output directory for multi-file commands. |
| `--backend <auto|officecli|wps-uia|wps-com|wps-jsapi>` | Backend selection. Default: `auto`. |
| `--overwrite` | Replace existing output. |
| `--timeout <seconds>` | Overall timeout. |
| `--cleanup <auto|always|never>` | UIA cleanup policy. Default: `auto`. |
| `--json` | Structured output. |
| `--verbose` | Include diagnostics in output. |
| `--dry-run` | Report planned backend/action without running it. |

Options that should move out of normal commands:

| Current Option | New Home |
|---|---|
| `--launch-mode` | `wps-uia raw ...` or `--debug-launch-mode` |
| `--preferred-verb` | `wps-uia raw` / diagnostics |
| `--wps-exe` | config/env/doctor, not common path |
| `--cloud-exe` | config/env/doctor |
| `--app-framework` | internal backend discovery |
| `--instance-id` | raw only |
| `--app-id` | raw only |
| `--app-name` | raw only |
| `--window-size` | raw only |
| `--src`, `--cloud-src` | raw only |
| `--cloud-app-params` | raw only |
| `--switchskin` | raw only |
| `--action` | raw only; normal command chooses action |
| `--runner-param`, `--runner-arg`, `--cloud-arg` | raw only |
| `--no-click` | diagnostics only |

## Output Path Semantics

The current implementation accepts `--output`, but the WPS converter itself
still writes to its default location first and then the wrapper moves the file.
That is a useful prototype behavior, but the stable CLI should be stricter:

- Rename `--output` to `--out`; keep `--output` as a deprecated alias.
- Always honor `--out` exactly.
- If WPS cannot directly write to that path, use a staging directory:
  1. Copy the input file into an isolated temp directory.
  2. Rename the staged input to match the requested output basename.
  3. Let WPS produce the default output in the staging directory.
  4. Move the result to `--out`.
  5. Clean up staging files and WPS windows.
- This avoids conflicts with an existing `<input>.docx` next to the source PDF.

## Proposed `pdf to-word` v2

User-facing:

```bash
office-tools pdf to-word input.pdf --out output.docx
```

Useful options:

```bash
office-tools pdf to-word input.pdf \
  --out output.docx \
  --backend wps-uia \
  --overwrite \
  --timeout 240 \
  --cleanup auto \
  --verbose
```

JSON result:

```json
{
  "ok": true,
  "command": "pdf.toWord",
  "backend": "wps-uia",
  "input": "input.pdf",
  "output": "output.docx",
  "events": [],
  "diagnostics": {
    "wpsVersion": "12.1.0.x",
    "launch": "shell-verb",
    "cleanup": []
  }
}
```

## Next Implementation Steps

1. Add `packages/core` with command schemas and common option parsing.
2. Replace `--output` with `--out` in CLI help; keep alias.
3. Move launcher internals to `office-tools wps-uia raw pdf-converter`.
4. Implement staging-dir output semantics for `pdf to-word`.
5. Add diagnostics commands: `env`, `verbs`, `windows`, `dump`, `watch-launch`.
6. Explore PDF converter variants in this order:
   - `to-excel`
   - `to-ppt`
   - `ocr`
   - `to-image`
   - `compress`
7. Only add merge/split after comparing local deterministic libraries against
   WPS UI behavior.

## References

- WPS Help: Convert PDF to Word with WPS Office
- WPS Help: Split-merge PDF Files with WPS Office
- WPS PDF tools pages: conversion, OCR, compression, merge/split, and batch PDF
