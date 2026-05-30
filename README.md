# office-tools-package

This directory is the cleaned starting point for the future open-source office automation package.

## Directory Layout

- `docs/`  
  Research notes, architecture notes, and migration notes from the WPS/Codex/OfficeCLI investigation.

- `packages/backend-wps-uia/`  
  WPS UI Automation backend. This is currently the most validated path for automating WPS member-only UI flows such as scanned PDF to Word conversion.

- `packages/backend-officecli/`  
  Placeholder for the OfficeCLI adapter. Use this for headless OOXML file-level operations.

- `packages/backend-wps-jsapi/`  
  Placeholder for a WPS JSAPI bridge backend. Use this when a WPS add-in/sidebar needs active document context.

- `packages/backend-wps-com/`  
  Placeholder for WPS COM automation experiments and adapters.

- `packages/cli/`  
  Placeholder for the unified command-line interface.

- `packages/mcp-server/`  
  Placeholder for an MCP server exposing the package capabilities to agents.

- `examples/wps-addin-bridge/`  
  Preserved WPS add-in bridge prototype.

- `examples/wps-rpc-invoke-prototype/`  
  Preserved RPC invocation prototype for reference.

## Recommended First Milestone

Start with `packages/backend-wps-uia/scripts/ConvertPdfToWord.ps1` and wrap it as a stable CLI command that returns JSON. After that, add the OfficeCLI adapter for deterministic file-level document operations.

## Current CLI Prototype

Install workspace links:

```powershell
npm install
```

Show available backend capabilities:

```powershell
node packages/cli/bin/office-tools.js capabilities
```

Convert PDF to Word through the WPS UIA backend:

```powershell
node packages/cli/bin/office-tools.js pdf to-word C:\path\input.pdf --backend wps-uia
```

Useful PDF conversion options:

```powershell
node packages/cli/bin/office-tools.js pdf to-word C:\path\input.pdf `
  --output C:\path\output.docx `
  --launch-mode native `
  --overwrite `
  --timeout 240 `
  --runner-param output_dir=C:\path
```

By default, the WPS UIA backend closes WPS windows/processes opened by the
conversion after the output file is ready. Use `--no-cleanup` only for debugging
the WPS UI state.

The WPS UIA backend currently exposes three launch modes:

- `shell`: invoke the registered Windows shell verb for WPS PDF conversion.
- `native`: call the observed `wps.exe Run /InstanceId=kpdf2wordv2 ...` runner.
- `cloud`: call the observed `wpscloudsvr.exe /app_id=kpdf2wordv2 ...`
  entrypoint. This mode exposes `--cloud-app-params` and `--cloud-arg` because
  the full app parameter schema still needs version-specific discovery.

See `docs/backend-capability-plan.md` for the OfficeCLI, WPS JSAPI, and WPS UIA
capability split and implementation roadmap.
