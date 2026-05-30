import { spawn } from "node:child_process";
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

export function convertPdfToWord(inputPdf, options = {}) {
  if (!inputPdf) {
    return Promise.reject(new Error("inputPdf is required"));
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

  return new Promise((resolve, reject) => {
    const child = spawn("powershell.exe", args, {
      windowsHide: false
    });

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

      try {
        resolve(parseJsonLine(stdout));
      } catch (error) {
        error.stdout = stdout;
        error.stderr = stderr;
        reject(error);
      }
    });
  });
}

export const capabilities = {
  backend: "wps-uia",
  commands: [
    {
      id: "pdf.toWord",
      description: "Convert a PDF to DOCX through the WPS PDF conversion desktop UI.",
      launchModes: ["shell", "native", "cloud"],
      options: [
        "outputPath",
        "timeoutSeconds",
        "noClick",
        "overwrite",
        "preferredVerbs",
        "wpsExe",
        "cloudExe",
        "appFramework",
        "instanceId",
        "appId",
        "appName",
        "windowSize",
        "src",
        "cloudSrc",
        "cloudAppParams",
        "switchSkin",
        "action",
        "runnerParams",
        "runnerArgs",
        "cloudArgs"
      ]
    }
  ]
};
