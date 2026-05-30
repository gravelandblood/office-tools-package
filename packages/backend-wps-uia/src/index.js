import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import { fileURLToPath } from "node:url";
import path from "node:path";

const scriptDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
  "scripts"
);

function pushSwitch(args, name, value) {
  if (value === undefined || value === null || value === "") return;
  args.push(name, String(value));
}

function pushFlag(args, name, enabled) {
  if (enabled) args.push(name);
}

function parseJsonLine(stdout) {
  const lines = stdout.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  for (let i = lines.length - 1; i >= 0; i -= 1) {
    try {
      return JSON.parse(lines[i]);
    } catch {}
  }
  throw new Error(`PowerShell did not return JSON. Output: ${stdout}`);
}

function normalizePowerShellError(stderr, stdout, fallbackCode) {
  const text = (stderr || stdout || "").trim();
  const firstLine = text.split(/\r?\n/).find(Boolean);
  const error = new Error(firstLine || `PowerShell exited with code ${fallbackCode}`);
  error.code = fallbackCode;
  error.stdout = stdout;
  error.stderr = stderr;
  return error;
}

function runPowerShell(args, { windowsHide = false } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn("powershell.exe", args, { windowsHide });

    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code !== 0) {
        reject(normalizePowerShellError(stderr, stdout, code));
        return;
      }

      resolve({ stdout, stderr });
    });
  });
}

async function runConvertScript(inputPdf, options = {}) {
  if (!inputPdf) {
    throw new Error("inputPdf is required");
  }

  const script = path.join(scriptDir, "ConvertPdfToWord.ps1");
  const args = [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    script,
    "-InputPdf",
    inputPdf
  ];

  pushSwitch(args, "-OutputPath", options.outputPath);
  pushSwitch(args, "-TimeoutSeconds", options.timeoutSeconds);
  pushSwitch(args, "-LaunchMode", options.launchMode);
  pushSwitch(args, "-WpsExe", options.wpsExe);
  pushSwitch(args, "-CloudExe", options.cloudExe);
  pushSwitch(args, "-AppFramework", options.appFramework);
  pushSwitch(args, "-InstanceId", options.instanceId);
  pushSwitch(args, "-AppId", options.appId);
  pushSwitch(args, "-AppName", options.appName);
  pushSwitch(args, "-WindowSize", options.windowSize);
  pushSwitch(args, "-Src", options.src);
  pushSwitch(args, "-CloudSrc", options.cloudSrc);
  pushSwitch(args, "-CloudAppParams", options.cloudAppParams);
  pushSwitch(args, "-SwitchSkin", options.switchSkin);
  pushSwitch(args, "-Action", options.action);
  pushSwitch(args, "-CleanupSeconds", options.cleanupSeconds);

  if (options.runnerParams && options.runnerParams.length) {
    args.push("-RunnerParam", ...options.runnerParams.map(String));
  }

  if (options.runnerArgs && options.runnerArgs.length) {
    args.push("-RunnerArg", ...options.runnerArgs.map(String));
  }

  if (options.cloudArgs && options.cloudArgs.length) {
    args.push("-CloudArg", ...options.cloudArgs.map(String));
  }

  if (options.preferredVerbs && options.preferredVerbs.length) {
    args.push("-PreferredVerb", ...options.preferredVerbs.map(String));
  }

  pushFlag(args, "-NoClick", options.noClick);
  pushFlag(args, "-Overwrite", options.overwrite);
  pushFlag(args, "-NoCleanup", options.noCleanup);

  const { stdout, stderr } = await runPowerShell(args, { windowsHide: false });
  try {
    return parseJsonLine(stdout);
  } catch (error) {
    error.stdout = stdout;
    error.stderr = stderr;
    throw error;
  }
}

function normalizeCleanupMode(mode) {
  if (!mode || mode === "auto" || mode === "always") return false;
  if (mode === "never") return true;
  throw new Error(`Unsupported cleanup mode: ${mode}`);
}

const pdfConverterTargets = {
  word: {
    command: "pdf.toWord",
    extension: ".docx",
    action: "ConvertToWord",
    tempPrefix: "office-tools-pdf2word-",
    launchMode: "shell"
  },
  excel: {
    command: "pdf.toExcel",
    extension: ".xlsx",
    action: "ConvertToExcel",
    tempPrefix: "office-tools-pdf2excel-",
    launchMode: "native"
  },
  ppt: {
    command: "pdf.toPpt",
    extension: ".pptx",
    action: "ConvertToPowerPoint",
    tempPrefix: "office-tools-pdf2ppt-",
    launchMode: "native"
  }
};

function toProductPdfConverterResult(target, result, options = {}, overrides = {}) {
  const product = {
    ok: result.ok,
    command: target.command,
    backend: result.backend,
    input: overrides.input || result.input,
    output: overrides.output || result.output,
    length: result.length,
    lastWriteTime: result.lastWriteTime,
    events: result.cleanup || []
  };

  if (options.verbose) {
    product.diagnostics = {
      launchMode: result.launchMode,
      action: result.action,
      shellVerb: result.shellVerb,
      runner: result.runner || undefined,
      staging: overrides.staging
    };
  }

  return product;
}

async function convertPdfWithWps(targetName, inputPdf, options = {}) {
  const target = pdfConverterTargets[targetName];
  if (!target) {
    throw new Error(`Unsupported PDF converter target: ${targetName}`);
  }

  const outPath = options.outPath || options.outputPath;
  const cleanupNever = options.noCleanup || normalizeCleanupMode(options.cleanup);
  const baseOptions = {
    timeoutSeconds: options.timeoutSeconds,
    overwrite: options.overwrite,
    noCleanup: cleanupNever,
    cleanupSeconds: options.cleanupSeconds,
    launchMode: options.launchMode || target.launchMode,
    preferredVerbs: options.preferredVerbs,
    action: target.action
  };

  if (!outPath) {
    const result = await runConvertScript(inputPdf, baseOptions);
    return toProductPdfConverterResult(target, result, options);
  }

  const resolvedInput = path.resolve(inputPdf);
  const resolvedOut = path.resolve(outPath);
  if (path.extname(resolvedOut).toLowerCase() !== target.extension) {
    throw new Error(`--out must end with ${target.extension} for ${target.command}: ${resolvedOut}`);
  }

  try {
    await fs.access(resolvedOut);
    if (!options.overwrite) {
      throw new Error(`Output already exists. Pass --overwrite to replace it: ${resolvedOut}`);
    }
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
  }

  const stagingDir = await fs.mkdtemp(path.join(os.tmpdir(), target.tempPrefix));
  const stagedPdf = path.join(stagingDir, `${path.basename(resolvedOut, target.extension)}.pdf`);
  try {
    await fs.copyFile(resolvedInput, stagedPdf);
    await fs.mkdir(path.dirname(resolvedOut), { recursive: true });
    const result = await runConvertScript(stagedPdf, {
      ...baseOptions,
      outputPath: resolvedOut,
      overwrite: options.overwrite
    });
    return toProductPdfConverterResult(target, result, options, {
      input: resolvedInput,
      output: resolvedOut,
      staging: {
        dir: stagingDir,
        input: stagedPdf
      }
    });
  } finally {
    try {
      await fs.rm(stagingDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 500 });
    } catch (error) {
      if (options.verbose) {
        process.stderr.write(`warning: unable to remove staging directory ${stagingDir}: ${error.message}\n`);
      }
    }
  }
}

export function convertPdfToWord(inputPdf, options = {}) {
  return convertPdfWithWps("word", inputPdf, options);
}

export function convertPdfToExcel(inputPdf, options = {}) {
  return convertPdfWithWps("excel", inputPdf, options);
}

export function convertPdfToPpt(inputPdf, options = {}) {
  return convertPdfWithWps("ppt", inputPdf, options);
}

export function convertPdfToWordRaw(inputPdf, options = {}) {
  return runConvertScript(inputPdf, options);
}

async function runJsonCommand(command) {
  const utf8Command = `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
${command}
`;
  const { stdout } = await runPowerShell([
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    utf8Command
  ], { windowsHide: true });
  return parseJsonLine(stdout);
}

export function getEnvironment() {
  const command = `
$root = Join-Path $env:LOCALAPPDATA 'Kingsoft\\WPS Office'
$versions = @()
if (Test-Path -LiteralPath $root) {
  $versions = @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue |
    Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'office6\\wps.exe') } |
    Sort-Object Name -Descending |
    ForEach-Object {
      $office6 = Join-Path $_.FullName 'office6'
      [pscustomobject]@{
        version = $_.Name
        office6 = $office6
        wps = Join-Path $office6 'wps.exe'
        wpspdf = Join-Path $office6 'wpspdf.exe'
        wpscloudsvr = Join-Path $office6 'wpscloudsvr.exe'
      }
    })
}
[pscustomobject]@{
  ok = $true
  backend = 'wps-uia'
  platform = $env:OS
  localAppData = $env:LOCALAPPDATA
  versions = $versions
} | ConvertTo-Json -Compress -Depth 6
`;
  return runJsonCommand(command);
}

export function listPdfVerbs(inputPdf) {
  const escaped = String(inputPdf).replace(/'/g, "''");
  const command = `
$pdf = (Resolve-Path -LiteralPath '${escaped}').Path
$shell = New-Object -ComObject Shell.Application
$folder = $shell.Namespace((Split-Path $pdf -Parent))
$item = $folder.ParseName((Split-Path $pdf -Leaf))
$verbs = @($item.Verbs()) | ForEach-Object { ($_.Name -replace '&','').Trim() } | Where-Object { $_ }
[pscustomobject]@{
  ok = $true
  backend = 'wps-uia'
  input = $pdf
  verbs = @($verbs)
} | ConvertTo-Json -Compress -Depth 4
`;
  return runJsonCommand(command);
}

export function listWindows() {
  const command = `
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes
$root=[System.Windows.Automation.AutomationElement]::RootElement
$wins=$root.FindAll([System.Windows.Automation.TreeScope]::Children,[System.Windows.Automation.Condition]::TrueCondition)
$items = @($wins | ForEach-Object {
  try {
    [pscustomobject]@{
      name=$_.Current.Name
      className=$_.Current.ClassName
      processId=$_.Current.ProcessId
      controlType=$_.Current.ControlType.ProgrammaticName
    }
  } catch {}
})
[pscustomobject]@{
  ok = $true
  backend = 'wps-uia'
  windows = $items
} | ConvertTo-Json -Compress -Depth 5
`;
  return runJsonCommand(command);
}

export const capabilities = {
  backend: "wps-uia",
  commands: [
    {
      id: "pdf.toWord",
      description: "Convert a PDF to DOCX through the WPS PDF conversion desktop UI.",
      options: [
        "outPath",
        "timeoutSeconds",
        "cleanup",
        "cleanupSeconds",
        "overwrite",
        "verbose"
      ]
    },
    {
      id: "pdf.toExcel",
      description: "Convert a PDF to XLSX through the WPS PDF conversion desktop UI.",
      options: [
        "outPath",
        "timeoutSeconds",
        "cleanup",
        "cleanupSeconds",
        "overwrite",
        "verbose"
      ]
    },
    {
      id: "pdf.toPpt",
      description: "Convert a PDF to PPTX through the WPS PDF conversion desktop UI.",
      options: [
        "outPath",
        "timeoutSeconds",
        "cleanup",
        "cleanupSeconds",
        "overwrite",
        "verbose"
      ]
    },
    {
      id: "wpsUia.env",
      description: "Inspect local WPS installation paths and versions."
    },
    {
      id: "wpsUia.verbs",
      description: "List Windows shell verbs for a PDF."
    },
    {
      id: "wpsUia.windows",
      description: "List top-level UI Automation windows."
    },
    {
      id: "wpsUia.raw.pdfConverter",
      description: "Experimental raw WPS PDF converter launcher with internal parameters."
    }
  ]
};
