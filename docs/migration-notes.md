# Migration Notes for office-tools-package

Date: 2026-05-30

## Files to Carry Forward

Research:

- `wps-bridge-research.md`
- `OFFICE_TOOLS_PACKAGE_BLUEPRINT.md`

WPS add-in bridge prototype:

- `codex-bridge-addin/index.html`
- `codex-bridge-addin/main.js`
- `codex-bridge-addin/manifest.xml`
- `codex-bridge-addin/ribbon.xml`
- `codex-bridge-addin/js/bridge.js`

WPS RPC/browser validation prototype:

- `codex-bridge-tools/invoke.html`
- `codex-bridge-tools/invoke.ps1`
- `codex-bridge-tools/jsplugins.xml`
- `codex-bridge-tools/serve.js`

WPS UIA prototype:

- `codex-bridge-tools/ConvertPdfToWord.ps1`
- `codex-bridge-tools/DumpWpsPdfUi.ps1`
- `codex-bridge-tools/ClickStartConvert.ps1`
- `codex-bridge-tools/ExtractDocxText.ps1`
- `codex-bridge-tools/watch-verb.ps1`

## Do Not Carry Forward

- `node_modules/`
- `.playwright-cli/`
- `converted-copy.docx`
- downloaded tarballs unless needed for offline reference

## First Production Refactor

1. Create `packages/backend-wps-uia`.
2. Port `ConvertPdfToWord.ps1` behavior into a Node-facing command with JSON
   output and explicit error codes.
3. Keep the PowerShell implementation as the first backend, but isolate it behind
   a stable function:

```ts
convertPdfToWord(inputPdf: string, options: {
  timeoutMs?: number;
  clickStart?: boolean;
}): Promise<{
  ok: true;
  backend: "wps-uia";
  input: string;
  output: string;
  shellVerb: string;
}>;
```

4. Add diagnostic commands:

```bash
office-tools wps-uia windows
office-tools wps-uia dump --title "WPS PDF conversion window"
office-tools wps-uia watch-verb input.pdf --verb convert-to-word
```

5. Add an adapter interface so OfficeCLI, WPS COM, WPS JSAPI, and WPS UIA can be
   selected by a common router.

## Important Design Constraints

- WPS UIA is Windows desktop only.
- WPS UIA should not claim to be headless.
- WPS member features must require an already-authorized user session.
- All commands should return JSON and include `backend`.
- Prefer deterministic file backends such as OfficeCLI before UIA.
- UIA workflows need version/theme diagnostics because WPS UI text and layout may
  change.
