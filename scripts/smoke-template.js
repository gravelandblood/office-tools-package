#!/usr/bin/env node
import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const cli = path.join(root, "packages", "cli", "bin", "office-tools.js");

function run(args) {
  return new Promise((resolve) => {
    const child = spawn("node", [cli, ...args], { cwd: root, windowsHide: true });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("close", (code) => resolve({ code, stdout, stderr, args }));
    child.on("error", (error) => resolve({ code: 1, stdout, stderr: `${stderr}\n${error.message}`, args }));
  });
}

function parseJson(result) {
  if (result.code !== 0) {
    throw new Error(`Command failed: ${result.args.join(" ")}\n${result.stdout}\n${result.stderr}`);
  }
  const parsed = JSON.parse(result.stdout);
  if (!parsed.ok) {
    throw new Error(`Command returned ok=false: ${result.args.join(" ")}\n${result.stdout}`);
  }
  return parsed;
}

async function writeSampleDocx(outPath) {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType } = await import("docx");
  const border = { style: BorderStyle.SINGLE, size: 4, color: "4472C4" };
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: "Arial", size: 22 },
          paragraph: { spacing: { after: 120 } }
        }
      }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1260, bottom: 1440, left: 1260 }
        }
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "项目尽调报告", bold: true, size: 32, color: "1F4E79" })]
        }),
        new Paragraph({
          children: [
            new TextRun({ text: "公司名称：", bold: true }),
            new TextRun("北京样例科技有限公司")
          ]
        }),
        new Paragraph({
          children: [
            new TextRun({ text: "报告日期：", bold: true }),
            new TextRun("2026年06月01日")
          ]
        }),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [3120, 3120, 3120],
          rows: [
            new TableRow({
              children: ["项目", "负责人", "状态"].map((text) => new TableCell({
                width: { size: 3120, type: WidthType.DXA },
                borders: { top: border, bottom: border, left: border, right: border },
                shading: { fill: "D9EAF7", type: ShadingType.CLEAR },
                children: [new Paragraph({ children: [new TextRun({ text, bold: true })] })]
              }))
            }),
            new TableRow({
              children: ["合同审查", "张三", "已完成"].map((text) => new TableCell({
                width: { size: 3120, type: WidthType.DXA },
                borders: { top: border, bottom: border, left: border, right: border },
                children: [new Paragraph(text)]
              }))
            })
          ]
        }),
        new Paragraph({
          children: [new TextRun({ text: "结论：", bold: true }), new TextRun("样例文档格式保持稳定。")]
        })
      ]
    }]
  });
  await fs.mkdir(path.dirname(outPath), { recursive: true });
  await fs.writeFile(outPath, await Packer.toBuffer(doc));
}

async function main() {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "office-tools-template-smoke-"));
  const sample = path.join(tempDir, "sample.docx");
  const template = path.join(tempDir, "template.docx");
  const data = path.join(tempDir, "data.json");
  const profile = path.join(tempDir, "profile.json");
  const detectiveProfile = path.join(tempDir, "detective-profile.json");
  const rendered = path.join(tempDir, "rendered.docx");

  await writeSampleDocx(sample);
  parseJson(await run(["template", "inspect-format", sample]));
  const profiled = parseJson(await run(["template", "profile", sample, "--out", detectiveProfile, "--summary"]));
  if (profiled.profile.summary.counts.paragraphs < 4 || profiled.profile.summary.counts.tables !== 1) {
    throw new Error(`Unexpected template profile summary: ${JSON.stringify(profiled, null, 2)}`);
  }
  const fullProfile = JSON.parse(await fs.readFile(detectiveProfile, "utf8"));
  const detailedPart = fullProfile.parts?.find((part) => part.paragraphs?.some((paragraph) => paragraph.runs?.length));
  if (!detailedPart || !Array.isArray(fullProfile.signals)) {
    throw new Error("Template profile did not include detailed paragraphs, runs, and signals.");
  }
  const inferred = parseJson(await run([
    "template", "infer-format", sample,
    "--out-template", template,
    "--out-data", data,
    "--out-profile", profile
  ]));
  if (inferred.fields < 8) {
    throw new Error(`Expected at least 8 text fields, got ${inferred.fields}`);
  }
  parseJson(await run(["template", "render", template, "--data", data, "--out", rendered]));
  const compared = parseJson(await run(["template", "compare-format", sample, rendered]));
  if (!compared.textEqual || !compared.profileEqual) {
    throw new Error(`Rendered DOCX does not match sample: ${JSON.stringify(compared, null, 2)}`);
  }

  const mutatedData = path.join(tempDir, "data-mutated.json");
  const mutatedRendered = path.join(tempDir, "rendered-mutated.docx");
  const payload = JSON.parse(await fs.readFile(data, "utf8"));
  const firstKey = Object.keys(payload.fields)[0];
  payload.fields[firstKey] = "替换后的动态文本";
  await fs.writeFile(mutatedData, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
  parseJson(await run(["template", "render", template, "--data", mutatedData, "--out", mutatedRendered]));
  const mutatedCompared = parseJson(await run(["template", "compare-format", sample, mutatedRendered]));
  if (mutatedCompared.textEqual || !mutatedCompared.profileEqual) {
    throw new Error(`Mutated render should preserve format but change text: ${JSON.stringify(mutatedCompared, null, 2)}`);
  }

  await fs.rm(tempDir, { recursive: true, force: true });
  console.log("template smoke OK");
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
