#!/usr/bin/env node
import { capabilities as officecliCapabilities } from "@office-tools/backend-officecli";
import { capabilities as wpsJsapiCapabilities } from "@office-tools/backend-wps-jsapi";
import { capabilities as wpsUiaCapabilities, convertPdfToWord } from "@office-tools/backend-wps-uia";

function printJson(value) {
  process.stdout.write(`${JSON.stringify(value, null, 2)}\n`);
}

function usage() {
  process.stdout.write(`office-tools

Usage:
  office-tools pdf to-word <input.pdf> [options]
  office-tools capabilities

Options:
  --backend wps-uia              Backend to use. Only wps-uia is implemented.
  --output <path>                Move the generated DOCX to this path.
  --timeout <seconds>            Wait time for the output file. Default: 180.
  --launch-mode <mode>           shell, native, or cloud. Default: shell.
  --no-click                     Launch the window but do not click start.
  --no-cleanup                   Leave WPS windows open after conversion.
  --cleanup-seconds <seconds>    Wait time for cleanup. Default: 20.
  --overwrite                    Replace an existing output file.
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
`);
}

function readOption(args, index) {
  const value = args[index + 1];
  if (!value || value.startsWith("--")) {
    throw new Error(`${args[index]} requires a value`);
  }
  return value;
}

function parsePdfToWord(argv) {
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
  let backend = "wps-uia";

  for (let i = 1; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case "--backend":
        backend = readOption(argv, i);
        i += 1;
        break;
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
        throw new Error(`Unknown option: ${arg}`);
    }
  }

  if (backend !== "wps-uia") {
    throw new Error(`Unsupported backend for pdf to-word: ${backend}`);
  }

  return { input, options };
}

async function main() {
  const [domain, command, ...rest] = process.argv.slice(2);
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
    const { input, options } = parsePdfToWord(rest);
    printJson(await convertPdfToWord(input, options));
    return;
  }

  throw new Error(`Unknown command: ${[domain, command].filter(Boolean).join(" ")}`);
}

main().catch((error) => {
  printJson({
    ok: false,
    error: error.message,
    code: error.code || 1
  });
  process.exitCode = typeof error.code === "number" ? error.code : 1;
});
