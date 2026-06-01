#!/usr/bin/env node
import { capabilities as officecliCapabilities } from "@office-tools/backend-officecli";
import {
  analyzeTemplates,
  capabilities as templateCapabilities,
  compareFormat,
  inferFormatTemplate,
  inspectFormat,
  planOfficeTemplate,
  profileDocx,
  renderTemplate
} from "@office-tools/backend-template";
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
  office-tools template inspect-format <input.docx> [options]
  office-tools template profile <input.docx> [options]
  office-tools template analyze <input.docx...> [options]
  office-tools template plan-office <template-ir.json> [options]
  office-tools template infer-format <input.docx> --out-template <template.docx> --out-data <data.json> --out-profile <profile.json>
  office-tools template render <template.docx> --data <data.json> --out <output.docx>
  office-tools template compare-format <left.docx> <right.docx> [options]
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

Template options:
  --out <path>                   Output file path.
  --out-template <path>          Generated DOCX template path.
  --out-data <path>              Generated JSON data path.
  --out-profile <path>           Generated JSON format profile path.
  --data <path>                  JSON data for rendering.
  --summary                      Print only profile summary and signals.
  --out-ir <path>                Generated Template Detective IR path.
  --out-plan <path>              Generated Office-native template patch plan path.
  --include-text                 Include extracted text in compare output.

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

function parseTemplateInfer(argv) {
  const input = argv[0];
  if (!input || input.startsWith("--")) {
    throw new Error("input DOCX is required");
  }

  const options = {};
  for (let i = 1; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case "--out-template":
        options.templatePath = readOption(argv, i);
        i += 1;
        break;
      case "--out-data":
        options.dataPath = readOption(argv, i);
        i += 1;
        break;
      case "--out-profile":
        options.profilePath = readOption(argv, i);
        i += 1;
        break;
      default:
        throw new Error(`Unknown template infer-format option: ${arg}`);
    }
  }

  return { input, options };
}

function parseTemplateProfile(argv) {
  const input = argv[0];
  if (!input || input.startsWith("--")) {
    throw new Error("template profile requires <input.docx>");
  }

  const options = {};
  for (let i = 1; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case "--out":
      case "--output":
        options.outputPath = readOption(argv, i);
        i += 1;
        break;
      case "--summary":
        options.summary = true;
        break;
      default:
        throw new Error(`Unknown template profile option: ${arg}`);
    }
  }

  return { input, options };
}

function parseTemplateAnalyze(argv) {
  const inputs = [];
  const options = {};

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg) continue;
    switch (arg) {
      case "--out":
      case "--output":
      case "--out-ir":
        options.outputPath = readOption(argv, i);
        i += 1;
        break;
      case "--summary":
        options.summary = true;
        break;
      default:
        if (arg.startsWith("--")) {
          throw new Error(`Unknown template analyze option: ${arg}`);
        }
        inputs.push(arg);
        break;
    }
  }

  if (!inputs.length) {
    throw new Error("template analyze requires <input.docx...>");
  }

  return { inputs, options };
}

function parseTemplatePlanOffice(argv) {
  const input = argv[0];
  if (!input || input.startsWith("--")) {
    throw new Error("template plan-office requires <template-ir.json>");
  }

  const options = {};
  for (let i = 1; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case "--out":
      case "--output":
      case "--out-plan":
        options.outputPath = readOption(argv, i);
        i += 1;
        break;
      case "--summary":
        options.summary = true;
        break;
      default:
        throw new Error(`Unknown template plan-office option: ${arg}`);
    }
  }

  return { input, options };
}

function parseTemplateRender(argv) {
  const input = argv[0];
  if (!input || input.startsWith("--")) {
    throw new Error("template DOCX is required");
  }

  const options = {};
  for (let i = 1; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case "--data":
        options.dataPath = readOption(argv, i);
        i += 1;
        break;
      case "--out":
      case "--output":
        options.outputPath = readOption(argv, i);
        i += 1;
        break;
      default:
        throw new Error(`Unknown template render option: ${arg}`);
    }
  }

  return { input, options };
}

function parseTemplateCompare(argv) {
  const left = argv[0];
  const right = argv[1];
  if (!left || left.startsWith("--") || !right || right.startsWith("--")) {
    throw new Error("template compare-format requires <left.docx> <right.docx>");
  }

  const options = {};
  for (let i = 2; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case "--include-text":
        options.includeText = true;
        break;
      default:
        throw new Error(`Unknown template compare-format option: ${arg}`);
    }
  }

  return { left, right, options };
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
        templateCapabilities,
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

  if (domain === "template" && command === "inspect-format") {
    if (!subcommand) throw new Error("template inspect-format requires <input.docx>");
    printJson(await inspectFormat(subcommand));
    return;
  }

  if (domain === "template" && command === "profile") {
    const { input, options } = parseTemplateProfile([subcommand, ...rest]);
    printJson(await profileDocx(input, options));
    return;
  }

  if (domain === "template" && command === "analyze") {
    const { inputs, options } = parseTemplateAnalyze([subcommand, ...rest]);
    printJson(await analyzeTemplates(inputs, options));
    return;
  }

  if (domain === "template" && command === "plan-office") {
    const { input, options } = parseTemplatePlanOffice([subcommand, ...rest]);
    printJson(await planOfficeTemplate(input, options));
    return;
  }

  if (domain === "template" && command === "infer-format") {
    const { input, options } = parseTemplateInfer([subcommand, ...rest]);
    printJson(await inferFormatTemplate(input, options));
    return;
  }

  if (domain === "template" && command === "render") {
    const { input, options } = parseTemplateRender([subcommand, ...rest]);
    printJson(await renderTemplate(input, options));
    return;
  }

  if (domain === "template" && command === "compare-format") {
    const { left, right, options } = parseTemplateCompare([subcommand, ...rest]);
    printJson(await compareFormat(left, right, options));
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
