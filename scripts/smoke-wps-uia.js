#!/usr/bin/env node
import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const cli = path.join(root, "packages", "cli", "bin", "office-tools.js");

function parseArgs(argv) {
  const options = {
    commands: ["to-image-pdf", "compress", "file-slim"],
    keepOutputs: false,
    timeoutSeconds: 240
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case "--sample":
        options.sample = argv[++i];
        break;
      case "--out-dir":
        options.outDir = argv[++i];
        break;
      case "--commands":
        options.commands = argv[++i].split(",").map((item) => item.trim()).filter(Boolean);
        break;
      case "--timeout":
        options.timeoutSeconds = Number(argv[++i]);
        break;
      case "--keep-outputs":
        options.keepOutputs = true;
        break;
      default:
        throw new Error(`Unknown option: ${arg}`);
    }
  }

  options.sample ||= process.env.OFFICE_TOOLS_SAMPLE_PDF;
  if (!options.sample) {
    throw new Error("Pass --sample <input.pdf> or set OFFICE_TOOLS_SAMPLE_PDF.");
  }
  return options;
}

function runOfficeTools(args) {
  return new Promise((resolve) => {
    const child = spawn("node", [cli, ...args], {
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
      resolve({ code, stdout, stderr, args });
    });
    child.on("error", (error) => {
      resolve({ code: 1, stdout, stderr: `${stderr}\n${error.message}`, args });
    });
  });
}

function parseJson(result) {
  if (result.code !== 0) {
    throw new Error(`Command failed: office-tools ${result.args.join(" ")}\n${result.stdout}\n${result.stderr}`.trim());
  }
  const parsed = JSON.parse(result.stdout);
  if (!parsed.ok) {
    throw new Error(`Command returned ok=false: office-tools ${result.args.join(" ")}\n${result.stdout}`);
  }
  return parsed;
}

async function assertNoResiduals() {
  const windows = parseJson(await runOfficeTools(["wps-uia", "windows"]));
  const badWindows = windows.windows.filter((win) => [
    "WPS PDF转换",
    "PDF压缩",
    "文件瘦身",
    "wpsoffice"
  ].some((text) => String(win.name || "").includes(text)));
  if (badWindows.length) {
    throw new Error(`Residual WPS windows found: ${JSON.stringify(badWindows, null, 2)}`);
  }

  const temp = os.tmpdir();
  const entries = await fs.readdir(temp, { withFileTypes: true });
  const badPrefixes = [
    "office-tools-wps-uia.lock",
    "office-tools-pdf2imagepdf-",
    "office-tools-pdfcompress-",
    "office-tools-fileslim-"
  ];
  const leftovers = entries
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .filter((name) => badPrefixes.some((prefix) => name.startsWith(prefix)));
  if (leftovers.length) {
    throw new Error(`Residual temp directories found: ${leftovers.join(", ")}`);
  }
}

async function main() {
  if (process.platform !== "win32") {
    throw new Error("WPS UIA smoke tests require Windows.");
  }

  const options = parseArgs(process.argv.slice(2));
  const sample = path.resolve(options.sample);
  await fs.access(sample);
  if (path.extname(sample).toLowerCase() !== ".pdf") {
    throw new Error(`Sample must be a PDF: ${sample}`);
  }

  const outDir = path.resolve(options.outDir || await fs.mkdtemp(path.join(os.tmpdir(), "office-tools-smoke-")));
  await fs.mkdir(outDir, { recursive: true });
  const outputs = [];

  const commands = {
    "to-image-pdf": ["pdf", "to-image-pdf", sample, "--out", path.join(outDir, "smoke-image-pdf.pdf"), "--timeout", String(options.timeoutSeconds), "--overwrite"],
    compress: ["pdf", "compress", sample, "--out", path.join(outDir, "smoke-compress.pdf"), "--timeout", String(options.timeoutSeconds), "--overwrite"],
    "file-slim": ["file", "slim", sample, "--out", path.join(outDir, "smoke-file-slim.pdf"), "--timeout", String(options.timeoutSeconds), "--overwrite"]
  };

  for (const commandName of options.commands) {
    const args = commands[commandName];
    if (!args) {
      throw new Error(`Unsupported smoke command: ${commandName}`);
    }
    const result = parseJson(await runOfficeTools(args));
    outputs.push(result.output);
    console.log(`${commandName} OK -> ${result.output}`);
  }

  await assertNoResiduals();

  if (!options.keepOutputs) {
    for (const output of outputs) {
      await fs.rm(output, { force: true });
    }
    if (!options.outDir) {
      await fs.rm(outDir, { recursive: true, force: true });
    }
  }

  console.log("wps-uia smoke OK");
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
