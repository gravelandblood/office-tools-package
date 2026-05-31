#!/usr/bin/env node
import { capabilities as officecliCapabilities } from "@office-tools/backend-officecli";
import { capabilities as wpsJsapiCapabilities } from "@office-tools/backend-wps-jsapi";
import {
  capabilities as wpsUiaCapabilities,
  compressPdf,
  convertPdfToExcel,
  convertPdfToImagePdf,
  convertPdfToPpt,
  convertPdfToWord,
  convertPdfToWordRaw,
  dumpWindow,
  getEnvironment,
  listPdfVerbs,
  listWindows,
  slimFile
} from "@office-tools/backend-wps-uia";

function printJson(value) {
  process.stdout.write(`${JSON.stringify(value, null, 2)}\n`);
}

function usage() {
  process.stdout.write(`office-tools

Usage:
  office-tools capabilities
  office-tools pdf to-word <input.pdf> --out <output.docx> [options]
  office-tools pdf to-excel <input.pdf> --out <output.xlsx> [options]
  office-tools pdf to-ppt <input.pdf> --out <output.pptx> [options]
  office-tools pdf to-image-pdf <input.pdf> --out <output.pdf> [options]
  office-tools pdf compress <input.pdf> --out <output.pdf> [options]
  office-tools file slim <input> --out <output> [options]
  office-tools wps-uia env
  office-tools wps-uia verbs <input.pdf>
  office-tools wps-uia windows
  office-tools wps-uia dump-window [options]
  office-tools wps-uia raw pdf-converter <input.pdf> [options]

PDF options:
  --out <path>                   Output path.
  --output <path>                Deprecated alias for --out.
  --backend <auto|wps-uia>       Backend to use. Default: auto.
  --timeout <seconds>            Wait time for the output file. Default: 180.
  --cleanup <auto|always|never>  Cleanup policy for UIA windows. Default: auto.
  --cleanup-seconds <seconds>    Wait time for cleanup. Default: 20.
  --level <high|standard|medium|low>
                                  Compression quality. Default: standard.
  --overwrite                    Replace an existing output file.
  --verbose                      Include staging diagnostics.

File options:
  --out <path>                   Output path.
  --output <path>                Deprecated alias for --out.
  --backend <auto|wps-uia>       Backend to use. Default: auto.
  --timeout <seconds>            Wait time for the output file. Default: 180.
  --cleanup <auto|always|never>  Cleanup policy for UIA windows. Default: auto.
  --cleanup-seconds <seconds>    Wait time for cleanup. Default: 20.
  --overwrite                    Replace an existing output file.
  --verbose                      Include staging diagnostics.

Raw WPS UIA options:
  --launch-mode <mode>           shell, native, or cloud. Default: shell.
  --no-click                     Launch the window but do not click start.
  --no-cleanup                   Leave WPS windows open after conversion.
  --preferred-verb <text>        Shell verb text to prefer. Can repeat.
  --wps-exe <path>               Path to wps.exe for native launch mode.
  --cloud-exe <path>             Path to wpscloudsvr.exe for cloud launch mode.
  --app-framework <path>         Path to kappessframework.dll for native launch mode.
  --instance-id <value>          WPS runner InstanceId. Default: kpdf2wordv2.
  --app-id <value>               WPS runner appId. Default: kpdf2wordv2.
  --app-name <value>             WPS runner appname. Default: WPS PDF Convert.
  --window-size <value>          WPS runner size. Default: 960&670.
  --src <value>                  WPS runner src marker.
  --cloud-src <value>            WPS cloud runner src marker.
  --cloud-app-params <value>     WPS cloud runner app_params payload.
  --switchskin <number>          WPS runner switchskin. Default: 0.
  --action <value>               WPS runner action. Default: ConvertToWord.
  --runner-param <key=value>     Extra WPS runner /key=value parameter. Can repeat.
  --runner-arg <value>           Raw extra WPS runner argument. Can repeat.
  --cloud-arg <value>            Raw extra WPS cloud runner argument. Can repeat.

WPS UIA diagnostics:
  --title <text>                 Dump windows whose title contains text.
  --process-id <pid>             Dump a specific process window.
  --max-depth <number>           Maximum UIA tree depth. Default: 6.
  --all                          Dump all top-level windows.
`);
}

function readOption(args, index) {
  const value = args[index + 1];
  if (!value || value.startsWith("--")) {
    throw new Error(`${args[index]} requires a value`);
  }
  return value;
}

function parsePdfConversion(command, argv) {
  const input = argv[0];
  if (!input || input.startsWith("--")) {
    throw new Error("input PDF is required");
  }

  const options = {};
  let backend = "auto";

  for (let i = 1; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case "--backend":
        backend = readOption(argv, i);
        i += 1;
        break;
      case "--out":
      case "--output":
        options.outPath = readOption(argv, i);
        i += 1;
        break;
      case "--timeout":
        options.timeoutSeconds = Number(readOption(argv, i));
        i += 1;
        break;
      case "--cleanup":
        options.cleanup = readOption(argv, i);
        i += 1;
        break;
      case "--cleanup-seconds":
        options.cleanupSeconds = Number(readOption(argv, i));
        i += 1;
        break;
      case "--overwrite":
        options.overwrite = true;
        break;
      case "--verbose":
        options.verbose = true;
        break;
      default:
        throw new Error(`Unknown pdf to-word option: ${arg}`);
    }
  }

  if (backend !== "auto" && backend !== "wps-uia") {
    throw new Error(`Unsupported backend for pdf ${command}: ${backend}`);
  }

  if (!options.outPath) {
    throw new Error(`pdf ${command} requires --out <output>`);
  }

  return { input, options };
}

function parsePdfCompress(argv) {
  const input = argv[0];
  if (!input || input.startsWith("--")) {
    throw new Error("input PDF is required");
  }

  const options = {};
  let backend = "auto";

  for (let i = 1; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case "--backend":
        backend = readOption(argv, i);
        i += 1;
        break;
      case "--out":
      case "--output":
        options.outPath = readOption(argv, i);
        i += 1;
        break;
      case "--level":
        options.level = readOption(argv, i);
        i += 1;
        break;
      case "--timeout":
        options.timeoutSeconds = Number(readOption(argv, i));
        i += 1;
        break;
      case "--cleanup":
        options.cleanup = readOption(argv, i);
        i += 1;
        break;
      case "--cleanup-seconds":
        options.cleanupSeconds = Number(readOption(argv, i));
        i += 1;
        break;
      case "--overwrite":
        options.overwrite = true;
        break;
      case "--verbose":
        options.verbose = true;
        break;
      default:
        throw new Error(`Unknown pdf compress option: ${arg}`);
    }
  }

  if (backend !== "auto" && backend !== "wps-uia") {
    throw new Error(`Unsupported backend for pdf compress: ${backend}`);
  }

  if (options.level && !["high", "standard", "medium", "low"].includes(options.level)) {
    throw new Error(`Unsupported pdf compress level: ${options.level}`);
  }

  if (!options.outPath) {
    throw new Error("pdf compress requires --out <output.pdf>");
  }

  return { input, options };
}

function parseFileSlim(argv) {
  const input = argv[0];
  if (!input || input.startsWith("--")) {
    throw new Error("input file is required");
  }

  const options = {};
  let backend = "auto";

  for (let i = 1; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case "--backend":
        backend = readOption(argv, i);
        i += 1;
        break;
      case "--out":
      case "--output":
        options.outPath = readOption(argv, i);
        i += 1;
        break;
      case "--timeout":
        options.timeoutSeconds = Number(readOption(argv, i));
        i += 1;
        break;
      case "--cleanup":
        options.cleanup = readOption(argv, i);
        i += 1;
        break;
      case "--cleanup-seconds":
        options.cleanupSeconds = Number(readOption(argv, i));
        i += 1;
        break;
      case "--overwrite":
        options.overwrite = true;
        break;
      case "--verbose":
        options.verbose = true;
        break;
      default:
        throw new Error(`Unknown file slim option: ${arg}`);
    }
  }

  if (backend !== "auto" && backend !== "wps-uia") {
    throw new Error(`Unsupported backend for file slim: ${backend}`);
  }

  if (!options.outPath) {
    throw new Error("file slim requires --out <output>");
  }

  return { input, options };
}

function parseRawPdfConverter(argv) {
  const input = argv[0];
  if (!input || input.startsWith("--")) {
    throw new Error("input PDF is required");
  }

  const options = {
    preferredVerbs: [],
    runnerParams: [],
    runnerArgs: [],
    cloudArgs: []
  };

  for (let i = 1; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case "--out":
      case "--output":
        options.outputPath = readOption(argv, i);
        i += 1;
        break;
      case "--timeout":
        options.timeoutSeconds = Number(readOption(argv, i));
        i += 1;
        break;
      case "--launch-mode":
        options.launchMode = readOption(argv, i);
        i += 1;
        break;
      case "--no-click":
        options.noClick = true;
        break;
      case "--no-cleanup":
        options.noCleanup = true;
        break;
      case "--cleanup-seconds":
        options.cleanupSeconds = Number(readOption(argv, i));
        i += 1;
        break;
      case "--overwrite":
        options.overwrite = true;
        break;
      case "--preferred-verb":
        options.preferredVerbs.push(readOption(argv, i));
        i += 1;
        break;
      case "--wps-exe":
        options.wpsExe = readOption(argv, i);
        i += 1;
        break;
      case "--cloud-exe":
        options.cloudExe = readOption(argv, i);
        i += 1;
        break;
      case "--app-framework":
        options.appFramework = readOption(argv, i);
        i += 1;
        break;
      case "--instance-id":
        options.instanceId = readOption(argv, i);
        i += 1;
        break;
      case "--app-id":
        options.appId = readOption(argv, i);
        i += 1;
        break;
      case "--app-name":
        options.appName = readOption(argv, i);
        i += 1;
        break;
      case "--window-size":
        options.windowSize = readOption(argv, i);
        i += 1;
        break;
      case "--src":
        options.src = readOption(argv, i);
        i += 1;
        break;
      case "--cloud-src":
        options.cloudSrc = readOption(argv, i);
        i += 1;
        break;
      case "--cloud-app-params":
        options.cloudAppParams = readOption(argv, i);
        i += 1;
        break;
      case "--switchskin":
        options.switchSkin = Number(readOption(argv, i));
        i += 1;
        break;
      case "--action":
        options.action = readOption(argv, i);
        i += 1;
        break;
      case "--runner-param":
        options.runnerParams.push(readOption(argv, i));
        i += 1;
        break;
      case "--runner-arg":
        options.runnerArgs.push(readOption(argv, i));
        i += 1;
        break;
      case "--cloud-arg":
        options.cloudArgs.push(readOption(argv, i));
        i += 1;
        break;
      default:
        throw new Error(`Unknown raw pdf-converter option: ${arg}`);
    }
  }

  return { input, options };
}

function parseDumpWindow(argv) {
  const options = {};

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case "--title":
        options.title = readOption(argv, i);
        i += 1;
        break;
      case "--process-id":
        options.processId = Number(readOption(argv, i));
        i += 1;
        break;
      case "--max-depth":
        options.maxDepth = Number(readOption(argv, i));
        i += 1;
        break;
      case "--timeout":
        options.timeoutSeconds = Number(readOption(argv, i));
        i += 1;
        break;
      case "--all":
        options.all = true;
        break;
      default:
        throw new Error(`Unknown wps-uia dump-window option: ${arg}`);
    }
  }

  if (!options.all && !options.title && !options.processId) {
    throw new Error("wps-uia dump-window requires --title <text>, --process-id <pid>, or --all");
  }

  return options;
}

async function main() {
  const [domain, command, subcommand, ...rest] = process.argv.slice(2);
  if (!domain || domain === "--help" || domain === "-h") {
    usage();
    return;
  }

  if (domain === "capabilities") {
    printJson({
      ok: true,
      backends: [
        officecliCapabilities,
        wpsJsapiCapabilities,
        wpsUiaCapabilities
      ]
    });
    return;
  }

  if (domain === "pdf" && command === "to-word") {
    const { input, options } = parsePdfConversion(command, [subcommand, ...rest]);
    printJson(await convertPdfToWord(input, options));
    return;
  }

  if (domain === "pdf" && command === "to-excel") {
    const { input, options } = parsePdfConversion(command, [subcommand, ...rest]);
    printJson(await convertPdfToExcel(input, options));
    return;
  }

  if (domain === "pdf" && command === "to-ppt") {
    const { input, options } = parsePdfConversion(command, [subcommand, ...rest]);
    printJson(await convertPdfToPpt(input, options));
    return;
  }

  if (domain === "pdf" && command === "to-image-pdf") {
    const { input, options } = parsePdfConversion(command, [subcommand, ...rest]);
    printJson(await convertPdfToImagePdf(input, options));
    return;
  }

  if (domain === "pdf" && command === "compress") {
    const { input, options } = parsePdfCompress([subcommand, ...rest]);
    printJson(await compressPdf(input, options));
    return;
  }

  if (domain === "file" && command === "slim") {
    const { input, options } = parseFileSlim([subcommand, ...rest]);
    printJson(await slimFile(input, options));
    return;
  }

  if (domain === "wps-uia" && command === "env") {
    printJson(await getEnvironment());
    return;
  }

  if (domain === "wps-uia" && command === "verbs") {
    if (!subcommand) throw new Error("wps-uia verbs requires <input.pdf>");
    printJson(await listPdfVerbs(subcommand));
    return;
  }

  if (domain === "wps-uia" && command === "windows") {
    printJson(await listWindows());
    return;
  }

  if (domain === "wps-uia" && command === "dump-window") {
    printJson(await dumpWindow(parseDumpWindow([subcommand, ...rest].filter(Boolean))));
    return;
  }

  if (domain === "wps-uia" && command === "raw" && subcommand === "pdf-converter") {
    const { input, options } = parseRawPdfConverter(rest);
    printJson(await convertPdfToWordRaw(input, options));
    return;
  }

  throw new Error(`Unknown command: ${[domain, command, subcommand].filter(Boolean).join(" ")}`);
}

main().catch((error) => {
  printJson({
    ok: false,
    error: error.message,
    code: error.code || 1
  });
  process.exitCode = typeof error.code === "number" ? error.code : 1;
});
