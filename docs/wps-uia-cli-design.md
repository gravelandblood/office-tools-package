# WPS UIA Capability and CLI Design

Date: 2026-05-31

This document keeps WPS UI Automation scoped as a product backend instead of a
bag of launcher flags.

## Why WPS UIA Exists

WPS UIA should be used only for product capabilities that are valuable but not
available through stable file APIs, COM, or JSAPI. It is a desktop automation
backend, not the default document engine.

Good UIA candidates:

- WPS exposes the feature through visible desktop UI.
- The feature depends on WPS account/member/cloud/product behavior.
- The feature is hard to reproduce with deterministic local libraries.
- The workflow can be represented as launch, configure, start, wait, collect,
  and clean up.

Poor UIA candidates:

- Plain OOXML reads/writes.
- Simple PDF merge/split/compress when local libraries can do the job.
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
| PDF to Word | `pdf to-word` | High | P0 | Implemented and validated, including OCR output for scanned PDFs. |
| PDF to Excel | `pdf to-excel` | High | P1 | Same converter family; verify action/window/output extension. |
| PDF to PPT | `pdf to-ppt` | High | P1 | Same converter family; verify fidelity and output extension. |
| Scanned PDF OCR | `pdf ocr` | High | P1 | Valuable when WPS OCR/member engine outperforms local OCR setup. |
| PDF to Image | `pdf to-image` | Medium | P1 | Local renderers may be better for deterministic image export. |
| PDF to Text | `pdf to-text` | Medium | P2 | Prefer local extraction first; UIA only for WPS OCR output. |
| PDF compress | `pdf compress` | Medium | P2 | Useful if WPS compression presets are desired. |
| PDF merge | `pdf merge` | Low/Medium | P2 | Prefer deterministic local merge unless WPS batch UI adds value. |
| PDF split/extract pages | `pdf split` / `pdf extract-pages` | Low/Medium | P2 | Prefer local libs unless WPS UI has a needed feature. |
| PDF edit text/image | `pdf edit` | Low | P3 | Too interactive for first stable CLI. |
| PDF watermark | `pdf watermark` | Medium | P3 | Could be scripted if UI controls are stable. |
| PDF sign/fill form | `pdf sign` / `pdf fill` | Low | P3 | High ambiguity and user-auth semantics. |
| PDF protect/unlock | `pdf protect` / `pdf unlock` | Low | P3 | Security-sensitive; avoid until scope is explicit. |
| OFD conversion | `ofd ...` | Medium | P3 | `wpsofd.exe` suggests a possible route; needs separate research. |
| WPS AI on document | `wps ai ...` | Experimental | P4 | Product/account/UI churn likely high. |

## CLI Design Principles

1. User commands describe outcomes, not WPS internals.
2. `--out` and `--out-dir` are first-class; output location must be honored.
3. Internal WPS launcher details live under raw/diagnostic commands.
4. `--backend auto` is the default; `wps-uia` is explicit or selected only when
   no safer backend exists.
5. Every command returns JSON by default once used by agents/MCP.
6. UIA commands clean up by default.
7. UIA commands expose diagnostics separately from the normal user path.

## Implemented Commands

Product:

```bash
office-tools pdf to-word input.pdf --out output.docx
```

Diagnostics:

```bash
office-tools wps-uia env
office-tools wps-uia verbs input.pdf
office-tools wps-uia windows
```

Raw/experimental escape hatch:

```bash
office-tools wps-uia raw pdf-converter input.pdf \
  --action ConvertToWord \
  --runner-param key=value
```

The raw command is intentionally not the main `pdf to-word` interface.

## Planned Command Shape

PDF conversion:

```bash
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

More diagnostics:

```bash
office-tools wps-uia dump --title "WPS PDF转换"
office-tools wps-uia watch-launch input.pdf --verb "转为Word"
```

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

Options moved out of normal commands:

| Raw Option | Home |
|---|---|
| `--launch-mode` | `wps-uia raw pdf-converter` |
| `--preferred-verb` | `wps-uia raw` / diagnostics |
| `--wps-exe` | config/env/doctor or raw |
| `--cloud-exe` | config/env/doctor or raw |
| `--app-framework` | internal backend discovery or raw |
| `--instance-id` | raw only |
| `--app-id` | raw only |
| `--app-name` | raw only |
| `--window-size` | raw only |
| `--src`, `--cloud-src` | raw only |
| `--cloud-app-params` | raw only |
| `--switchskin` | raw only |
| `--action` | raw only; normal command chooses action |
| `--runner-param`, `--runner-arg`, `--cloud-arg` | raw only |
| `--no-click` | diagnostics/raw only |

## Output Path Semantics

The product CLI accepts `--out` and honors it exactly. Since the WPS converter
writes beside the input file by default, the wrapper uses staging:

1. Copy the input file into an isolated temp directory.
2. Rename the staged input to match the requested output basename.
3. Let WPS produce the default output in the staging directory.
4. Clean up WPS windows/processes so the generated DOCX is unlocked.
5. Move the result to `--out`.
6. Clean up staging files.

`--output` remains as a deprecated alias.

## Current `pdf to-word`

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
  "events": []
}
```

With `--verbose`, launcher details move under `diagnostics`.

## Next Implementation Steps

1. Add `packages/core` with command schemas and common option parsing.
2. Add diagnostics commands: `dump`, `watch-launch`.
3. Explore PDF converter variants in this order:
   - `to-excel`
   - `to-ppt`
   - `ocr`
   - `to-image`
   - `compress`
4. Only add merge/split after comparing local deterministic libraries against
   WPS UI behavior.

## References

- WPS Help: Convert PDF to Word with WPS Office
- WPS Help: Split-merge PDF Files with WPS Office
- WPS PDF tools pages: conversion, OCR, compression, merge/split, and batch PDF
