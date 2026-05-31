import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

const WORD_TEXT_PART = /^word\/(document|footnotes|endnotes|comments)\.xml$|^word\/(header|footer)\d+\.xml$/;
const XML_DECLARATION = /^<\?xml[^>]*>\s*/u;
const TEXT_NODE_RE = /<w:t\b([^>]*)>([\s\S]*?)<\/w:t>/gu;
const TAG_RE = /<[^>]+>/gu;

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "",
  preserveOrder: false,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false
});

export const capabilities = {
  backend: "template",
  commands: [
    {
      id: "template.inferFormat",
      description: "Create a DOCX format-preserving text template plus data/profile JSON from a sample DOCX."
    },
    {
      id: "template.render",
      description: "Render a DOCX template produced by template.inferFormat with JSON data."
    },
    {
      id: "template.compareFormat",
      description: "Compare DOCX text and OOXML formatting fingerprints."
    },
    {
      id: "template.inspectFormat",
      description: "Inspect DOCX structure and formatting fingerprints."
    }
  ]
};

function escapeXml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function decodeXmlText(text) {
  return String(text)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function normalizeXml(xml) {
  return xml.replace(XML_DECLARATION, "").trim();
}

function asArray(value) {
  if (value === undefined || value === null) return [];
  return Array.isArray(value) ? value : [value];
}

function stableStringify(value) {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  const keys = Object.keys(value).sort();
  return `{${keys.map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`).join(",")}}`;
}

async function loadDocx(filePath) {
  const buffer = await fs.readFile(filePath);
  return JSZip.loadAsync(buffer);
}

async function writeDocx(zip, outputPath) {
  await fs.mkdir(path.dirname(path.resolve(outputPath)), { recursive: true });
  const buffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE"
  });
  await fs.writeFile(outputPath, buffer);
}

function getWordXmlParts(zip) {
  return Object.values(zip.files)
    .filter((file) => !file.dir && WORD_TEXT_PART.test(file.name))
    .sort((a, b) => a.name.localeCompare(b.name));
}

function collectTextNodes(xml, partName) {
  const nodes = [];
  let index = 0;
  for (const match of xml.matchAll(TEXT_NODE_RE)) {
    nodes.push({
      id: `t${String(index + 1).padStart(4, "0")}`,
      part: partName,
      index,
      attrs: match[1],
      raw: match[2],
      text: decodeXmlText(match[2])
    });
    index += 1;
  }
  return nodes;
}

function templateXml(xml, partName, entries) {
  let index = 0;
  return xml.replace(TEXT_NODE_RE, (full, attrs) => {
    const entry = entries.find((item) => item.part === partName && item.index === index);
    index += 1;
    if (!entry) return full;
    const placeholder = `{{${entry.key}}}`;
    const nextAttrs = attrs.includes("xml:space=") ? attrs : `${attrs} xml:space="preserve"`;
    return `<w:t${nextAttrs}>${escapeXml(placeholder)}</w:t>`;
  });
}

function renderXml(xml, data) {
  return xml.replace(TEXT_NODE_RE, (full, attrs, raw) => {
    const text = decodeXmlText(raw);
    const match = text.match(/^\{\{([A-Za-z0-9_.-]+)\}\}$/u);
    if (!match) return full;
    const value = data[match[1]];
    const rendered = value === undefined || value === null ? "" : String(value);
    const nextAttrs = attrs.includes("xml:space=") ? attrs : `${attrs} xml:space="preserve"`;
    return `<w:t${nextAttrs}>${escapeXml(rendered)}</w:t>`;
  });
}

function textFromXml(xml) {
  return collectTextNodes(xml, "").map((node) => node.text).join("");
}

function textFromZipParts(parts) {
  return parts.map((part) => textFromXml(part.xml)).join("\n");
}

function parseXmlPart(xml) {
  return parser.parse(xml);
}

function hashFingerprint(value) {
  let hash = 2166136261;
  const input = stableStringify(value);
  for (let i = 0; i < input.length; i += 1) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function stripText(obj) {
  if (obj === null || obj === undefined) return obj;
  if (typeof obj !== "object") return typeof obj === "string" ? "" : obj;
  if (Array.isArray(obj)) return obj.map(stripText);
  const out = {};
  for (const [key, value] of Object.entries(obj)) {
    if (key === "w:t" || key === "#text") continue;
    out[key] = stripText(value);
  }
  return out;
}

function collectParagraphs(node, result = []) {
  if (!node || typeof node !== "object") return result;
  if (Array.isArray(node)) {
    for (const item of node) collectParagraphs(item, result);
    return result;
  }
  if (node["w:p"]) {
    for (const p of asArray(node["w:p"])) {
      result.push({
        style: p["w:pPr"]?.["w:pStyle"]?.val || null,
        fingerprint: hashFingerprint(stripText(p)),
        text: extractTextFromParsed(p)
      });
    }
  }
  for (const value of Object.values(node)) {
    collectParagraphs(value, result);
  }
  return result;
}

function collectTables(node, result = []) {
  if (!node || typeof node !== "object") return result;
  if (Array.isArray(node)) {
    for (const item of node) collectTables(item, result);
    return result;
  }
  if (node["w:tbl"]) {
    for (const tbl of asArray(node["w:tbl"])) {
      const rows = asArray(tbl["w:tr"]);
      result.push({
        rows: rows.length,
        columns: Math.max(0, ...rows.map((row) => asArray(row["w:tc"]).length)),
        fingerprint: hashFingerprint(stripText(tbl))
      });
    }
  }
  for (const value of Object.values(node)) {
    collectTables(value, result);
  }
  return result;
}

function extractTextFromParsed(node) {
  if (node === null || node === undefined) return "";
  if (typeof node === "string") return node;
  if (typeof node !== "object") return "";
  if (Array.isArray(node)) return node.map(extractTextFromParsed).join("");
  let out = "";
  if (typeof node["w:t"] === "string") out += node["w:t"];
  for (const value of Object.values(node)) {
    out += extractTextFromParsed(value);
  }
  return out;
}

async function readWordParts(zip) {
  const parts = [];
  for (const file of getWordXmlParts(zip)) {
    const xml = await file.async("string");
    parts.push({ name: file.name, xml });
  }
  return parts;
}

async function inspectZip(zip) {
  const parts = await readWordParts(zip);
  const textEntries = parts.flatMap((part) => collectTextNodes(part.xml, part.name));
  const styleXml = zip.file("word/styles.xml") ? await zip.file("word/styles.xml").async("string") : "";
  const numberingXml = zip.file("word/numbering.xml") ? await zip.file("word/numbering.xml").async("string") : "";
  const documentPart = parts.find((part) => part.name === "word/document.xml");
  const parsedDocument = documentPart ? parseXmlPart(documentPart.xml) : {};

  const paragraphs = collectParagraphs(parsedDocument);
  const tables = collectTables(parsedDocument);
  const runLikeCount = (documentPart?.xml.match(/<w:r\b/gu) || []).length;

  return {
    parts: parts.map((part) => ({
      name: part.name,
      textNodes: collectTextNodes(part.xml, part.name).length,
      textLength: textFromXml(part.xml).length
    })),
    counts: {
      textNodes: textEntries.length,
      paragraphs: paragraphs.length,
      runs: runLikeCount,
      tables: tables.length,
      stylesBytes: styleXml.length,
      numberingBytes: numberingXml.length
    },
    paragraphs: paragraphs.map((paragraph, index) => ({
      index,
      style: paragraph.style,
      textLength: paragraph.text.length,
      fingerprint: paragraph.fingerprint
    })),
    tables: tables.map((table, index) => ({
      index,
      rows: table.rows,
      columns: table.columns,
      fingerprint: table.fingerprint
    })),
    packageFingerprint: hashFingerprint({
      parts: parts.map((part) => ({
        name: part.name,
        structure: normalizeXml(part.xml).replace(TEXT_NODE_RE, "<w:t/>")
      })),
      styles: normalizeXml(styleXml),
      numbering: normalizeXml(numberingXml)
    })
  };
}

export async function inspectFormat(inputDocx) {
  const zip = await loadDocx(inputDocx);
  const profile = await inspectZip(zip);
  return {
    ok: true,
    command: "template.inspectFormat",
    input: path.resolve(inputDocx),
    profile
  };
}

export async function inferFormatTemplate(inputDocx, options = {}) {
  const templatePath = options.templatePath || options.outTemplate;
  const dataPath = options.dataPath || options.outData;
  const profilePath = options.profilePath || options.outProfile;
  if (!templatePath || !dataPath || !profilePath) {
    throw new Error("template infer-format requires --out-template, --out-data, and --out-profile");
  }

  const zip = await loadDocx(inputDocx);
  const parts = await readWordParts(zip);
  const entries = [];
  const data = {};

  for (const part of parts) {
    for (const node of collectTextNodes(part.xml, part.name)) {
      const key = `text.${String(entries.length + 1).padStart(4, "0")}`;
      const entry = {
        key,
        part: node.part,
        index: node.index,
        text: node.text,
        textLength: node.text.length
      };
      entries.push(entry);
      data[key] = node.text;
    }
  }

  for (const part of parts) {
    const nextXml = templateXml(part.xml, part.name, entries);
    zip.file(part.name, nextXml);
  }

  await writeDocx(zip, templatePath);

  const profile = await inspectZip(zip);
  const outData = {
    version: 1,
    source: path.resolve(inputDocx),
    template: path.resolve(templatePath),
    fields: data
  };
  const outProfile = {
    version: 1,
    source: path.resolve(inputDocx),
    template: path.resolve(templatePath),
    mode: "format-preserving-text-placeholders",
    placeholders: entries,
    profile
  };

  await fs.mkdir(path.dirname(path.resolve(dataPath)), { recursive: true });
  await fs.writeFile(dataPath, `${JSON.stringify(outData, null, 2)}\n`, "utf8");
  await fs.mkdir(path.dirname(path.resolve(profilePath)), { recursive: true });
  await fs.writeFile(profilePath, `${JSON.stringify(outProfile, null, 2)}\n`, "utf8");

  return {
    ok: true,
    command: "template.inferFormat",
    input: path.resolve(inputDocx),
    template: path.resolve(templatePath),
    data: path.resolve(dataPath),
    profile: path.resolve(profilePath),
    fields: entries.length,
    profileFingerprint: profile.packageFingerprint
  };
}

export async function renderTemplate(templateDocx, options = {}) {
  const dataPath = options.dataPath || options.data;
  const outputPath = options.outputPath || options.out;
  if (!dataPath || !outputPath) {
    throw new Error("template render requires --data and --out");
  }

  const payload = JSON.parse(await fs.readFile(dataPath, "utf8"));
  const data = payload.fields || payload;
  const zip = await loadDocx(templateDocx);
  const parts = await readWordParts(zip);
  for (const part of parts) {
    zip.file(part.name, renderXml(part.xml, data));
  }
  await writeDocx(zip, outputPath);
  const profile = await inspectZip(zip);

  return {
    ok: true,
    command: "template.render",
    template: path.resolve(templateDocx),
    data: path.resolve(dataPath),
    output: path.resolve(outputPath),
    profileFingerprint: profile.packageFingerprint
  };
}

export async function compareFormat(leftDocx, rightDocx, options = {}) {
  const leftZip = await loadDocx(leftDocx);
  const rightZip = await loadDocx(rightDocx);
  const leftParts = await readWordParts(leftZip);
  const rightParts = await readWordParts(rightZip);
  const leftProfile = await inspectZip(leftZip);
  const rightProfile = await inspectZip(rightZip);
  const leftText = textFromZipParts(leftParts);
  const rightText = textFromZipParts(rightParts);
  const textEqual = leftText === rightText;
  const profileEqual = leftProfile.packageFingerprint === rightProfile.packageFingerprint;
  const paragraphMismatches = [];
  const paragraphCount = Math.max(leftProfile.paragraphs.length, rightProfile.paragraphs.length);
  for (let index = 0; index < paragraphCount; index += 1) {
    const leftParagraph = leftProfile.paragraphs[index];
    const rightParagraph = rightProfile.paragraphs[index];
    const leftFormat = leftParagraph ? { style: leftParagraph.style, fingerprint: leftParagraph.fingerprint } : null;
    const rightFormat = rightParagraph ? { style: rightParagraph.style, fingerprint: rightParagraph.fingerprint } : null;
    if (stableStringify(leftFormat) !== stableStringify(rightFormat)) {
      paragraphMismatches.push({ index, left: leftParagraph || null, right: rightParagraph || null });
      if (paragraphMismatches.length >= 5) break;
    }
  }

  const tableMismatches = [];
  const tableCount = Math.max(leftProfile.tables.length, rightProfile.tables.length);
  for (let index = 0; index < tableCount; index += 1) {
    const leftTable = leftProfile.tables[index];
    const rightTable = rightProfile.tables[index];
    if (stableStringify(leftTable) !== stableStringify(rightTable)) {
      tableMismatches.push({ index, left: leftTable || null, right: rightTable || null });
      if (tableMismatches.length >= 5) break;
    }
  }

  const result = {
    ok: true,
    command: "template.compareFormat",
    left: path.resolve(leftDocx),
    right: path.resolve(rightDocx),
    textEqual,
    profileEqual,
    leftTextLength: leftText.length,
    rightTextLength: rightText.length,
    leftFingerprint: leftProfile.packageFingerprint,
    rightFingerprint: rightProfile.packageFingerprint,
    counts: {
      left: leftProfile.counts,
      right: rightProfile.counts
    },
    mismatches: {
      paragraphs: paragraphMismatches,
      tables: tableMismatches
    }
  };

  if (options.includeText) {
    result.leftText = leftText;
    result.rightText = rightText;
  }

  return result;
}
