#!/usr/bin/env node
import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const isWindows = process.platform === "win32";

function run(command, args, options = {}) {
  return new Promise((resolve) => {
    const child = spawn(command, args, {
      cwd: root,
      windowsHide: true
    });

    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });
    child.on("close", (code) => {
      resolve({ code, stdout, stderr, label: options.label || [command, ...args].join(" ") });
    });
    child.on("error", (error) => {
      resolve({ code: 1, stdout, stderr: `${stderr}\n${error.message}`, label: options.label || [command, ...args].join(" ") });
    });
  });
}

function assertOk(result) {
  if (result.code !== 0) {
    throw new Error(`${result.label} failed\n${result.stdout}\n${result.stderr}`.trim());
  }
}

async function readJson(file) {
  return JSON.parse(await fs.readFile(path.join(root, file), "utf8"));
}

async function main() {
  const checks = [];

  for (const file of [
    "packages/cli/bin/office-tools.js",
    "packages/backend-template/src/index.js",
    "packages/backend-wps-uia/src/index.js",
    "packages/backend-officecli/src/index.js",
    "packages/backend-wps-jsapi/src/index.js"
  ]) {
    checks.push(run("node", ["--check", file], { label: `node --check ${file}` }));
  }

  if (isWindows) {
    const psScript = [
      "$ErrorActionPreference='Stop'",
      "$scripts=@('ConvertPdfToWord.ps1','CompressPdf.ps1','DumpWindow.ps1','SlimFile.ps1')",
      "foreach($s in $scripts){",
      "  $tokens=$null; $errs=$null",
      "  [System.Management.Automation.Language.Parser]::ParseFile(\"packages/backend-wps-uia/scripts/$s\", [ref]$tokens, [ref]$errs) > $null",
      "  if($errs.Count){ $errs | Format-List *; exit 1 }",
      "}"
    ].join("; ");
    checks.push(run("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", psScript], {
      label: "PowerShell script parse"
    }));
  }

  const results = await Promise.all(checks);
  for (const result of results) {
    assertOk(result);
  }

  const rootPackage = await readJson("package.json");
  const cliPackage = await readJson("packages/cli/package.json");
  if (!rootPackage.scripts?.verify || !rootPackage.scripts?.["smoke:wps-uia"]) {
    throw new Error("Root package scripts must include verify and smoke:wps-uia.");
  }
  if (cliPackage.bin?.["office-tools"] !== "bin/office-tools.js") {
    throw new Error("@office-tools/cli must expose the office-tools bin.");
  }

  const capabilities = await run("node", ["packages/cli/bin/office-tools.js", "capabilities"], {
    label: "office-tools capabilities"
  });
  assertOk(capabilities);
  const parsed = JSON.parse(capabilities.stdout);
  const wpsUia = parsed.backends.find((backend) => backend.backend === "wps-uia");
  const template = parsed.backends.find((backend) => backend.backend === "template");
  const ids = new Set(wpsUia.commands.map((command) => command.id));
  for (const id of [
    "pdf.toWord",
    "pdf.toExcel",
    "pdf.toPpt",
    "pdf.toImagePdf",
    "pdf.compress",
    "file.slim",
    "wpsUia.dumpWindow"
  ]) {
    if (!ids.has(id)) {
      throw new Error(`Missing capability: ${id}`);
    }
  }
  const templateIds = new Set(template.commands.map((command) => command.id));
  for (const id of [
    "template.inferFormat",
    "template.render",
    "template.compareFormat",
    "template.inspectFormat",
    "template.profile",
    "template.analyze",
    "template.planOffice",
    "template.compileOffice",
    "template.renderOffice"
  ]) {
    if (!templateIds.has(id)) {
      throw new Error(`Missing capability: ${id}`);
    }
  }

  const help = await run("node", ["packages/cli/bin/office-tools.js", "--help"], {
    label: "office-tools --help"
  });
  assertOk(help);
  for (const text of ["pdf to-image-pdf", "file slim", "template infer-format", "template profile", "template analyze", "template plan-office", "template compile-office", "template render-office", "wps-uia dump-window"]) {
    if (!help.stdout.includes(text)) {
      throw new Error(`Help output missing: ${text}`);
    }
  }

  console.log("verify OK");
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
