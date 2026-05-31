#!/usr/bin/env node
import { profileDocx } from "../../../packages/backend-template/src/index.js";

async function main() {
  const [input, ...rest] = process.argv.slice(2);
  if (!input) {
    throw new Error("Usage: node profile-docx.mjs <input.docx> [--out profile.json] [--summary]");
  }

  const options = {};
  for (let i = 0; i < rest.length; i += 1) {
    const arg = rest[i];
    if (arg === "--out" || arg === "--output") {
      const value = rest[i + 1];
      if (!value || value.startsWith("--")) throw new Error(`${arg} requires a value`);
      options.outputPath = value;
      i += 1;
    } else if (arg === "--summary") {
      options.summary = true;
    } else {
      throw new Error(`Unknown option: ${arg}`);
    }
  }

  process.stdout.write(`${JSON.stringify(await profileDocx(input, options), null, 2)}\n`);
}

main().catch((error) => {
  process.stderr.write(`${error.message}\n`);
  process.exitCode = 1;
});
