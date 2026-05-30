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
