import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

const WORD_TEXT_PART = /^word\/(document|footnotes|endnotes|comments)\.xml$|^word\/(header|footer)\d+\.xml$/;
const XML_DECLARATION = /^<\?xml[^>]*>\s*/u;
const TEXT_NODE_RE = /<w:t\b([^>]*)>([\s\S]*?)<\/w:t>/gu;
const TAG_RE = /<[^>]+>/gu;
const PARSED_TEXT_KEYS = new Set(["w:t", "w:delText", "w:instrText", "w:delInstrText"]);
const OFFICE_TEMPLATE_STORE_ITEM_ID = "{B5F04A55-6F5D-4F32-9E3F-0B6720E8E1D2}";

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
    },
    {
      id: "template.profile",
      description: "Extract a normalized DOCX template-detection profile with evidence, formatting atoms, and conflict signals."
    },
    {
      id: "template.analyze",
      description: "Analyze one or more DOCX examples into a Template Detective IR with rule candidates and conflicts."
    },
    {
      id: "template.planOffice",
      description: "Map Template Detective IR to an Office-native template patch plan using content controls, custom XML, repeating sections, and sidecar gaps."
    },
    {
      id: "template.compileOffice",
      description: "Compile safe Office-native template patches into a DOCX using content controls and custom XML parts."
    }
  ]
};

function escapeXml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeXmlAttr(text) {
  return escapeXml(text)
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
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

function paragraphBlocksFromXml(xml) {
  const blocks = [];
  const re = /<w:p\b[\s\S]*?<\/w:p>/gu;
  let index = 0;
  for (const match of xml.matchAll(re)) {
    blocks.push({
      index,
      start: match.index,
      end: match.index + match[0].length,
      xml: match[0],
      text: textFromXml(match[0])
    });
    index += 1;
  }
  return blocks;
}

function tableBlocksFromXml(xml) {
  const blocks = [];
  const re = /<w:tbl\b[\s\S]*?<\/w:tbl>/gu;
  let index = 0;
  for (const match of xml.matchAll(re)) {
    blocks.push({
      index,
      start: match.index,
      end: match.index + match[0].length,
      xml: match[0],
      text: textFromXml(match[0])
    });
    index += 1;
  }
  return blocks;
}

function paragraphFingerprintFromXml(paragraphXml) {
  try {
    const parsed = parseXmlPart(paragraphXml);
    const paragraph = parsed["w:p"];
    return paragraph ? hashFingerprint(stripText(paragraph)) : null;
  } catch {
    return null;
  }
}

function tableFingerprintFromXml(tableXml) {
  try {
    const parsed = parseXmlPart(tableXml);
    const table = parsed["w:tbl"];
    return table ? hashFingerprint(stripText(table)) : null;
  } catch {
    return null;
  }
}

function scoreTextLocation(block, evidence, kind) {
  const preview = evidence.textPreview || "";
  const expectedFingerprint = evidence.fingerprint || null;
  const hintedIndex = kind === "table" ? evidence.tableIndex : evidence.paragraphIndex;
  const text = kind === "table" ? normalizePreviewText(block.text) : block.text;
  const normalizedPreview = kind === "table" ? normalizePreviewText(preview) : preview;
  let score = 0;
  const reasons = [];

  if (Number.isInteger(hintedIndex)) {
    const distance = Math.abs(block.index - hintedIndex);
    const indexScore = distance === 0 ? 0.35 : Math.max(0, 0.2 - distance * 0.02);
    if (indexScore > 0) {
      score += indexScore;
      reasons.push(distance === 0 ? "index-exact" : `index-near:${distance}`);
    }
  }

  if (normalizedPreview && text === normalizedPreview) {
    score += 0.35;
    reasons.push("text-exact");
  } else if (normalizedPreview && text.includes(normalizedPreview)) {
    score += 0.28;
    reasons.push("text-preview");
  }

  if (expectedFingerprint && block.fingerprint === expectedFingerprint) {
    score += 0.3;
    reasons.push("fingerprint");
  }

  return {
    id: `${kind}.${String(block.index).padStart(4, "0")}`,
    kind,
    index: block.index,
    confidence: Number(Math.min(score, 1).toFixed(3)),
    reasons,
    textPreview: block.text.slice(0, 120),
    fingerprint: block.fingerprint || null,
    range: {
      start: block.start,
      end: block.end
    }
  };
}

function locateParagraphBlock(xml, evidence) {
  const blocks = paragraphBlocksFromXml(xml).map((block) => ({
    ...block,
    fingerprint: paragraphFingerprintFromXml(block.xml)
  }));
  const candidates = blocks
    .map((block) => ({ block, evidence: scoreTextLocation(block, evidence, "paragraph") }))
    .filter((item) => item.evidence.confidence > 0)
    .sort((left, right) => right.evidence.confidence - left.evidence.confidence || left.block.index - right.block.index);
  const best = candidates[0] || null;
  const runnerUp = candidates[1] || null;
  const unique = best && (!runnerUp || best.evidence.confidence - runnerUp.evidence.confidence >= 0.2);
  return {
    block: unique && best.evidence.confidence >= 0.7 ? best.block : null,
    candidates: candidates.slice(0, 5).map((item) => item.evidence),
    confidence: best?.evidence.confidence || 0,
    ambiguous: Boolean(best && !unique)
  };
}

function blockAtIndex(xml, kind, index) {
  if (!Number.isInteger(index)) return null;
  if (kind === "table") {
    return tableBlocksFromXml(xml).map((block) => ({
      ...block,
      fingerprint: tableFingerprintFromXml(block.xml)
    }))[index] || null;
  }
  return paragraphBlocksFromXml(xml).map((block) => ({
    ...block,
    fingerprint: paragraphFingerprintFromXml(block.xml)
  }))[index] || null;
}

function normalizePreviewText(text) {
  return String(text || "").replace(/[\s|]+/gu, "");
}

function locateTableBlock(xml, evidence) {
  const blocks = tableBlocksFromXml(xml).map((block) => ({
    ...block,
    fingerprint: tableFingerprintFromXml(block.xml)
  }));
  const candidates = blocks
    .map((block) => ({ block, evidence: scoreTextLocation(block, evidence, "table") }))
    .filter((item) => item.evidence.confidence > 0)
    .sort((left, right) => right.evidence.confidence - left.evidence.confidence || left.block.index - right.block.index);
  const best = candidates[0] || null;
  const runnerUp = candidates[1] || null;
  const unique = best && (!runnerUp || best.evidence.confidence - runnerUp.evidence.confidence >= 0.2);
  return {
    block: unique && best.evidence.confidence >= 0.55 ? best.block : null,
    candidates: candidates.slice(0, 5).map((item) => item.evidence),
    confidence: best?.evidence.confidence || 0,
    ambiguous: Boolean(best && !unique)
  };
}

function hasUnsupportedRangeMarkup(xml) {
  return /<w:(sdt|fldChar|instrText|drawing|pict|bookmarkStart|bookmarkEnd)\b/u.test(xml);
}

function wrapRunWithContentControl(runXml, patch) {
  const title = escapeXmlAttr(patch.contentControl?.title || patch.ruleId || patch.id);
  const tag = escapeXmlAttr(patch.contentControl?.tag || patch.ruleId || patch.id);
  const binding = patch.customXmlBinding;
  const dataBinding = binding
    ? `<w:dataBinding w:storeItemID="${escapeXmlAttr(binding.storeItemId || OFFICE_TEMPLATE_STORE_ITEM_ID)}" w:xpath="${escapeXmlAttr(binding.xpath)}" w:prefixMappings="${escapeXmlAttr(binding.prefixMappings || "")}"/>`
    : "";
  return [
    "<w:sdt>",
    "<w:sdtPr>",
    `<w:alias w:val="${title}"/>`,
    `<w:tag w:val="${tag}"/>`,
    "<w:text/>",
    dataBinding,
    "</w:sdtPr>",
    "<w:sdtContent>",
    runXml,
    "</w:sdtContent>",
    "</w:sdt>"
  ].join("");
}

function wrapRowsWithRepeatingSection(rowsXml, patch) {
  const title = escapeXmlAttr(patch.contentControl?.title || patch.ruleId || patch.id);
  const tag = escapeXmlAttr(patch.contentControl?.tag || patch.ruleId || patch.id);
  return [
    "<w:sdt>",
    "<w:sdtPr>",
    `<w:alias w:val="${title}"/>`,
    `<w:tag w:val="${tag}"/>`,
    "<w15:repeatingSection/>",
    "</w:sdtPr>",
    "<w:sdtContent>",
    "<w:sdt>",
    "<w:sdtPr><w15:repeatingSectionItem/></w:sdtPr>",
    "<w:sdtContent>",
    rowsXml,
    "</w:sdtContent>",
    "</w:sdt>",
    "</w:sdtContent>",
    "</w:sdt>"
  ].join("");
}

function ensureW15Namespace(xml) {
  if (xml.includes("xmlns:w15=")) return xml;
  return xml.replace(/<w:document\b([^>]*)>/u, (full, attrs) => (
    `<w:document${attrs} xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">`
  ));
}

function rowBlocksFromTableXml(tableXml) {
  const rows = [];
  const re = /<w:tr\b[\s\S]*?<\/w:tr>/gu;
  let index = 0;
  for (const match of tableXml.matchAll(re)) {
    rows.push({
      index,
      start: match.index,
      end: match.index + match[0].length,
      xml: match[0],
      text: textFromXml(match[0])
    });
    index += 1;
  }
  return rows;
}

function findSingleRunForLabelValue(paragraphXml, label, value) {
  if (!label || !value || hasUnsupportedRangeMarkup(paragraphXml)) return null;
  const paragraphText = textFromXml(paragraphXml);
  if (!paragraphText.includes(label) || !paragraphText.endsWith(value)) return null;

  const runRe = /<w:r\b[\s\S]*?<\/w:r>/gu;
  const runs = Array.from(paragraphXml.matchAll(runRe)).map((match) => ({
    start: match.index,
    end: match.index + match[0].length,
    xml: match[0],
    text: textFromXml(match[0])
  }));
  const exactRuns = runs.filter((run) => run.text === value);
  if (exactRuns.length === 1) return exactRuns[0];

  const suffixRuns = runs.filter((run) => run.text.endsWith(value) && !run.text.includes(label));
  if (suffixRuns.length === 1) return suffixRuns[0];

  return null;
}

function findSingleRunForWholeParagraphValue(paragraphXml, value) {
  if (!value || hasUnsupportedRangeMarkup(paragraphXml)) return null;
  if (textFromXml(paragraphXml) !== value) return null;

  const runRe = /<w:r\b[\s\S]*?<\/w:r>/gu;
  const runs = Array.from(paragraphXml.matchAll(runRe)).map((match) => ({
    start: match.index,
    end: match.index + match[0].length,
    xml: match[0],
    text: textFromXml(match[0])
  })).filter((run) => run.text.length > 0);
  if (runs.length === 1 && runs[0].text === value) return runs[0];
  return null;
}

function compileSlotPatchIntoPart(xml, patch, override = null) {
  const evidence = Array.isArray(patch.evidence) ? patch.evidence.find((item) => item.part && Number.isInteger(item.paragraphIndex)) : null;
  if (!evidence || patch.ruleKind !== "slot" || !patch.contentControl) {
    return { xml, applied: false, reason: "patch is not a paragraph slot with evidence" };
  }

  const location = locateParagraphBlock(xml, evidence);
  const overrideBlock = override?.kind === "paragraph" ? blockAtIndex(xml, "paragraph", override.index) : null;
  const paragraph = overrideBlock || location.block;
  if (!paragraph) {
    return {
      xml,
      applied: false,
      reason: location.ambiguous ? "target paragraph is ambiguous" : "target paragraph was not found by text and fingerprint",
      confidence: overrideBlock ? override.confidence ?? location.confidence : location.confidence,
      candidates: location.candidates
    };
  }

  const label = patch.label || null;
  const preview = evidence.textPreview || "";
  const value = label && preview.startsWith(label) ? preview.slice(label.length).replace(/^[:\uFF1A]\s*/u, "") : null;
  const run = label
    ? findSingleRunForLabelValue(paragraph.xml, label, value)
    : findSingleRunForWholeParagraphValue(paragraph.xml, preview);
  if (!run) {
    return {
      xml,
      applied: false,
      reason: "slot value is not isolated in one safe run",
      confidence: overrideBlock ? override.confidence ?? location.confidence : location.confidence,
      candidates: location.candidates
    };
  }

  const wrappedParagraph = `${paragraph.xml.slice(0, run.start)}${wrapRunWithContentControl(run.xml, patch)}${paragraph.xml.slice(run.end)}`;
  return {
    xml: `${xml.slice(0, paragraph.start)}${wrappedParagraph}${xml.slice(paragraph.end)}`,
    applied: true,
    confidence: overrideBlock ? override.confidence ?? location.confidence : location.confidence,
    candidates: location.candidates,
    overrideApplied: Boolean(overrideBlock),
    reason: "wrapped value run with plain-text content control"
  };
}

function compileLoopPatchIntoPart(xml, patch, override = null) {
  const evidence = Array.isArray(patch.evidence) ? patch.evidence.find((item) => item.part && Number.isInteger(item.tableIndex)) : null;
  if (!evidence || patch.ruleKind !== "loop" || !patch.contentControl) {
    return { xml, applied: false, reason: "patch is not a table loop with evidence" };
  }

  const location = locateTableBlock(xml, evidence);
  const overrideBlock = override?.kind === "table" ? blockAtIndex(xml, "table", override.index) : null;
  const table = overrideBlock || location.block;
  if (!table) {
    return {
      xml,
      applied: false,
      reason: location.ambiguous ? "target table is ambiguous" : "target table was not found by text and fingerprint",
      confidence: overrideBlock ? override.confidence ?? location.confidence : location.confidence,
      candidates: location.candidates
    };
  }
  if (hasUnsupportedRangeMarkup(table.xml)) {
    return { xml, applied: false, reason: "target table already contains unsupported range markup", confidence: overrideBlock ? override.confidence ?? location.confidence : location.confidence, candidates: location.candidates };
  }

  const rows = rowBlocksFromTableXml(table.xml);
  if (rows.length < 2) {
    return { xml, applied: false, reason: "table does not have a header row and repeatable body rows", confidence: overrideBlock ? override.confidence ?? location.confidence : location.confidence, candidates: location.candidates };
  }

  const bodyRows = rows.slice(1);
  const bodyStart = bodyRows[0].start;
  const bodyEnd = bodyRows[bodyRows.length - 1].end;
  const bodyXml = table.xml.slice(bodyStart, bodyEnd);
  const wrappedBody = wrapRowsWithRepeatingSection(bodyXml, patch);
  const wrappedTable = `${table.xml.slice(0, bodyStart)}${wrappedBody}${table.xml.slice(bodyEnd)}`;
  return {
    xml: ensureW15Namespace(`${xml.slice(0, table.start)}${wrappedTable}${xml.slice(table.end)}`),
    applied: true,
    confidence: overrideBlock ? override.confidence ?? location.confidence : location.confidence,
    candidates: location.candidates,
    overrideApplied: Boolean(overrideBlock),
    reason: "wrapped table body rows with repeating section content control"
  };
}

function uniqueSortedBindablePatches(patches) {
  const byPath = new Map();
  for (const patch of patches) {
    if (!["slot", "loop"].includes(patch.ruleKind) || !patch.customXmlBinding?.xpath) continue;
    if (!byPath.has(patch.customXmlBinding.xpath)) byPath.set(patch.customXmlBinding.xpath, patch);
  }
  return Array.from(byPath.values()).sort((left, right) => (
    left.customXmlBinding.xpath.localeCompare(right.customXmlBinding.xpath)
  ));
}

function setDeepXmlValue(root, segments, value) {
  let cursor = root;
  for (const segment of segments) {
    if (!segment) continue;
    if (!cursor.children.has(segment)) cursor.children.set(segment, { children: new Map(), value: "" });
    cursor = cursor.children.get(segment);
  }
  cursor.value = value;
}

function xmlTreeToString(name, node) {
  const children = Array.from(node.children.entries())
    .map(([childName, child]) => xmlTreeToString(childName, child))
    .join("");
  if (children) return `<${name}>${children}</${name}>`;
  return `<${name}>${escapeXml(node.value || "")}</${name}>`;
}

function patchInitialValue(patch) {
  const value = Array.isArray(patch.evidence) ? patch.evidence[0]?.textPreview || "" : "";
  if (patch.ruleKind === "loop") {
    return "";
  }
  const label = patch.label ? `${patch.label}${value.includes("：") ? "：" : ":"}` : "";
  return patch.label && value.startsWith(patch.label)
    ? value.slice(label.length).trim()
    : value;
}

function buildCustomXmlData(patches) {
  const root = { children: new Map(), value: "" };
  for (const patch of uniqueSortedBindablePatches(patches)) {
    const segments = patch.customXmlBinding.xpath
      .replace(/^\/template\/data\/?/u, "")
      .split("/")
      .filter(Boolean);
    setDeepXmlValue(root, segments, patchInitialValue(patch));
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<template><data>${Array.from(root.children.entries()).map(([name, node]) => xmlTreeToString(name, node)).join("")}</data></template>`;
}

function nextCustomXmlIndex(zip) {
  const indexes = Object.keys(zip.files)
    .map((name) => name.match(/^customXml\/item(\d+)\.xml$/u))
    .filter(Boolean)
    .map((match) => Number(match[1]));
  return indexes.length ? Math.max(...indexes) + 1 : 1;
}

function relationshipXml(id, type, target) {
  return `<Relationship Id="${escapeXmlAttr(id)}" Type="${escapeXmlAttr(type)}" Target="${escapeXmlAttr(target)}"/>`;
}

function upsertRelationship(relsXml, relationship) {
  if (relsXml.includes(`Target="${escapeXmlAttr(relationship.target)}"`)) return relsXml;
  const entry = relationshipXml(relationship.id, relationship.type, relationship.target);
  if (relsXml.includes("</Relationships>")) {
    return relsXml.replace("</Relationships>", `${entry}</Relationships>`);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entry}</Relationships>`;
}

function contentTypeOverride(partName, contentType) {
  return `<Override PartName="${escapeXmlAttr(partName)}" ContentType="${escapeXmlAttr(contentType)}"/>`;
}

function upsertContentTypeOverride(contentTypesXml, partName, contentType) {
  if (contentTypesXml.includes(`PartName="${escapeXmlAttr(partName)}"`)) return contentTypesXml;
  const entry = contentTypeOverride(partName, contentType);
  if (contentTypesXml.includes("</Types>")) {
    return contentTypesXml.replace("</Types>", `${entry}</Types>`);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${entry}</Types>`;
}

async function addCustomXmlPart(zip, patches) {
  const bindablePatches = patches.filter((patch) => ["slot", "loop"].includes(patch.ruleKind) && patch.customXmlBinding);
  if (!bindablePatches.length) return null;

  const itemIndex = nextCustomXmlIndex(zip);
  const itemPath = `customXml/item${itemIndex}.xml`;
  const propsPath = `customXml/itemProps${itemIndex}.xml`;
  const itemRelsPath = `customXml/_rels/item${itemIndex}.xml.rels`;
  const data = buildCustomXmlData(bindablePatches);
  const props = [
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    `<ds:datastoreItem ds:itemID="${OFFICE_TEMPLATE_STORE_ITEM_ID}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">`,
    "<ds:schemaRefs/>",
    "</ds:datastoreItem>"
  ].join("");
  const itemRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps\" Target=\"itemProps" + itemIndex + ".xml\"/></Relationships>";

  zip.file(itemPath, data);
  zip.file(propsPath, props);
  zip.file(itemRelsPath, itemRels);

  const relsPath = "_rels/.rels";
  const relsXml = zip.file(relsPath) ? await zip.file(relsPath).async("string") : "";
  const relId = `rIdOfficeToolsTemplate${crypto.createHash("sha1").update(itemPath).digest("hex").slice(0, 8)}`;
  zip.file(relsPath, upsertRelationship(relsXml, {
    id: relId,
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
    target: itemPath
  }));

  const contentTypesPath = "[Content_Types].xml";
  const contentTypesXml = zip.file(contentTypesPath) ? await zip.file(contentTypesPath).async("string") : "";
  let nextContentTypes = upsertContentTypeOverride(contentTypesXml, `/${itemPath}`, "application/xml");
  nextContentTypes = upsertContentTypeOverride(nextContentTypes, `/${propsPath}`, "application/vnd.openxmlformats-officedocument.customXmlProperties+xml");
  zip.file(contentTypesPath, nextContentTypes);

  return {
    item: itemPath,
    properties: propsPath,
    relationships: itemRelsPath,
    fields: bindablePatches.filter((patch) => patch.ruleKind === "slot").length,
    arrays: bindablePatches.filter((patch) => patch.ruleKind === "loop").length
  };
}

async function loadCandidateOverrides(options) {
  const overridePath = options.acceptCandidatesPath || options.acceptCandidates;
  if (!overridePath) return new Map();
  const payload = JSON.parse((await fs.readFile(overridePath, "utf8")).replace(/^\uFEFF/u, ""));
  const items = Array.isArray(payload) ? payload : payload.acceptedCandidates || payload.overrides || [];
  const overrides = new Map();
  for (const item of items) {
    const patchId = item.patchId || item.id;
    if (!patchId) continue;
    overrides.set(patchId, {
      kind: item.kind,
      index: item.index,
      confidence: item.confidence,
      reason: item.reason || "accepted-candidate"
    });
  }
  return overrides;
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

function stripEmpty(value) {
  if (value === undefined || value === null || value === "") return undefined;
  if (Array.isArray(value)) {
    const items = value.map(stripEmpty).filter((item) => item !== undefined);
    return items.length ? items : undefined;
  }
  if (typeof value === "object") {
    const out = {};
    for (const [key, item] of Object.entries(value)) {
      const next = stripEmpty(item);
      if (next !== undefined) out[key] = next;
    }
    return Object.keys(out).length ? out : undefined;
  }
  return value;
}

function cloneJson(value) {
  if (value === undefined) return undefined;
  return JSON.parse(JSON.stringify(value));
}

function firstVal(node, key, attr = "w:val") {
  const item = asArray(node?.[key])[0];
  if (item === undefined || item === null) return null;
  if (typeof item === "string") return item;
  if (typeof item === "object" && attr in item) return item[attr];
  return null;
}

function profileProps(value) {
  return stripEmpty(cloneJson(value)) || null;
}

function hasKeyDeep(node, key) {
  if (!node || typeof node !== "object") return false;
  if (Array.isArray(node)) return node.some((item) => hasKeyDeep(item, key));
  if (Object.prototype.hasOwnProperty.call(node, key)) return true;
  return Object.values(node).some((value) => hasKeyDeep(value, key));
}

function collectNodesByKey(node, key, result = []) {
  if (!node || typeof node !== "object") return result;
  if (Array.isArray(node)) {
    for (const item of node) collectNodesByKey(item, key, result);
    return result;
  }
  if (node[key]) {
    for (const item of asArray(node[key])) result.push(item);
  }
  for (const value of Object.values(node)) collectNodesByKey(value, key, result);
  return result;
}

function collectRunsFromParagraph(paragraph, partName, paragraphIndex) {
  return asArray(paragraph["w:r"]).map((run, index) => {
    const text = extractTextFromParsed(run);
    const rPr = profileProps(run["w:rPr"]);
    return {
      id: `r${String(paragraphIndex + 1).padStart(4, "0")}.${String(index + 1).padStart(3, "0")}`,
      index,
      text,
      textLength: text.length,
      rPr,
      styleId: firstVal(run["w:rPr"], "w:rStyle"),
      formatFingerprint: hashFingerprint(rPr || {}),
      structureFingerprint: hashFingerprint(stripText(run)),
      flags: {
        hasDrawing: hasKeyDeep(run, "w:drawing"),
        hasPicture: hasKeyDeep(run, "w:pict"),
        hasField: hasKeyDeep(run, "w:fldChar") || hasKeyDeep(run, "w:instrText"),
        hasTab: hasKeyDeep(run, "w:tab"),
        hasBreak: hasKeyDeep(run, "w:br")
      },
      evidence: {
        part: partName,
        path: `paragraphs[${paragraphIndex}].runs[${index}]`
      }
    };
  });
}

function paragraphProfile(paragraph, partName, index) {
  const pPr = profileProps(paragraph["w:pPr"]);
  const text = extractTextFromParsed(paragraph);
  const runs = collectRunsFromParagraph(paragraph, partName, index);
  return {
    id: `p${String(index + 1).padStart(4, "0")}`,
    part: partName,
    index,
    text,
    textLength: text.length,
    styleId: firstVal(paragraph["w:pPr"], "w:pStyle"),
    numbering: profileProps(paragraph["w:pPr"]?.["w:numPr"]),
    alignment: firstVal(paragraph["w:pPr"], "w:jc"),
    spacing: profileProps(paragraph["w:pPr"]?.["w:spacing"]),
    indent: profileProps(paragraph["w:pPr"]?.["w:ind"]),
    pPr,
    runCount: runs.length,
    runs,
    formatFingerprint: hashFingerprint(stripText(paragraph)),
    paragraphPropertiesFingerprint: hashFingerprint(pPr || {}),
    contentFingerprint: hashFingerprint(text),
    flags: {
      hasDrawing: hasKeyDeep(paragraph, "w:drawing"),
      hasPicture: hasKeyDeep(paragraph, "w:pict"),
      hasField: hasKeyDeep(paragraph, "w:fldChar") || hasKeyDeep(paragraph, "w:instrText"),
      hasBookmark: hasKeyDeep(paragraph, "w:bookmarkStart") || hasKeyDeep(paragraph, "w:bookmarkEnd")
    },
    evidence: {
      part: partName,
      path: `paragraphs[${index}]`
    }
  };
}

function tableProfile(table, partName, index) {
  const rows = asArray(table["w:tr"]).map((row, rowIndex) => {
    const cells = asArray(row["w:tc"]).map((cell, cellIndex) => {
      const paragraphs = collectNodesByKey(cell, "w:p", []);
      const tcPr = profileProps(cell["w:tcPr"]);
      return {
        index: cellIndex,
        text: extractTextFromParsed(cell),
        paragraphCount: paragraphs.length,
        tcPr,
        width: profileProps(cell["w:tcPr"]?.["w:tcW"]),
        gridSpan: firstVal(cell["w:tcPr"], "w:gridSpan"),
        vMerge: firstVal(cell["w:tcPr"], "w:vMerge"),
        shading: profileProps(cell["w:tcPr"]?.["w:shd"]),
        borders: profileProps(cell["w:tcPr"]?.["w:tcBorders"]),
        formatFingerprint: hashFingerprint(stripText(cell)),
        cellPropertiesFingerprint: hashFingerprint(tcPr || {})
      };
    });
    const trPr = profileProps(row["w:trPr"]);
    return {
      index: rowIndex,
      cellCount: cells.length,
      text: extractTextFromParsed(row),
      trPr,
      cells,
      formatFingerprint: hashFingerprint(stripText(row)),
      rowPropertiesFingerprint: hashFingerprint(trPr || {})
    };
  });
  const tblPr = profileProps(table["w:tblPr"]);
  const tblGrid = profileProps(table["w:tblGrid"]);
  return {
    id: `tbl${String(index + 1).padStart(4, "0")}`,
    part: partName,
    index,
    rowCount: rows.length,
    columnCount: Math.max(0, ...rows.map((row) => row.cellCount)),
    textLength: extractTextFromParsed(table).length,
    tblPr,
    tblGrid,
    rows,
    formatFingerprint: hashFingerprint(stripText(table)),
    tablePropertiesFingerprint: hashFingerprint({ tblPr, tblGrid }),
    evidence: {
      part: partName,
      path: `tables[${index}]`
    }
  };
}

function collectDetailedParagraphs(parsedPart, partName) {
  return collectNodesByKey(parsedPart, "w:p", []).map((paragraph, index) => (
    paragraphProfile(paragraph, partName, index)
  ));
}

function collectDetailedTables(parsedPart, partName) {
  return collectNodesByKey(parsedPart, "w:tbl", []).map((table, index) => (
    tableProfile(table, partName, index)
  ));
}

function collectStyles(stylesXml) {
  if (!stylesXml) return [];
  const parsed = parseXmlPart(stylesXml);
  return asArray(parsed["w:styles"]?.["w:style"]).map((style, index) => {
    const pPr = profileProps(style["w:pPr"]);
    const rPr = profileProps(style["w:rPr"]);
    const tblPr = profileProps(style["w:tblPr"]);
    return {
      index,
      styleId: style["w:styleId"] || null,
      type: style["w:type"] || null,
      name: firstVal(style, "w:name"),
      basedOn: firstVal(style, "w:basedOn"),
      next: firstVal(style, "w:next"),
      linked: firstVal(style, "w:link"),
      aliases: firstVal(style, "w:aliases"),
      isDefault: style["w:default"] === "1" || style["w:default"] === "true",
      pPr,
      rPr,
      tblPr,
      fingerprint: hashFingerprint(stripText(style))
    };
  });
}

function collectNumbering(numberingXml) {
  if (!numberingXml) return { abstractNums: [], nums: [] };
  const parsed = parseXmlPart(numberingXml);
  const numbering = parsed["w:numbering"] || {};
  return {
    abstractNums: asArray(numbering["w:abstractNum"]).map((item, index) => ({
      index,
      abstractNumId: item["w:abstractNumId"] || null,
      levelCount: asArray(item["w:lvl"]).length,
      levels: asArray(item["w:lvl"]).map((level) => ({
        ilvl: level["w:ilvl"] || null,
        numFmt: firstVal(level, "w:numFmt"),
        lvlText: firstVal(level, "w:lvlText"),
        start: firstVal(level, "w:start"),
        pPr: profileProps(level["w:pPr"]),
        rPr: profileProps(level["w:rPr"]),
        fingerprint: hashFingerprint(stripText(level))
      })),
      fingerprint: hashFingerprint(stripText(item))
    })),
    nums: asArray(numbering["w:num"]).map((item, index) => ({
      index,
      numId: item["w:numId"] || null,
      abstractNumId: firstVal(item, "w:abstractNumId"),
      fingerprint: hashFingerprint(stripText(item))
    }))
  };
}

function collectSections(parsedPart, partName) {
  return collectNodesByKey(parsedPart, "w:sectPr", []).map((section, index) => ({
    id: `sect${String(index + 1).padStart(4, "0")}`,
    part: partName,
    index,
    pageSize: profileProps(section["w:pgSz"]),
    pageMargins: profileProps(section["w:pgMar"]),
    columns: profileProps(section["w:cols"]),
    headers: asArray(section["w:headerReference"]).map(profileProps),
    footers: asArray(section["w:footerReference"]).map(profileProps),
    fingerprint: hashFingerprint(stripText(section)),
    evidence: {
      part: partName,
      path: `sections[${index}]`
    }
  }));
}

function summarizeProfilePart(partProfile) {
  return {
    name: partProfile.name,
    textLength: partProfile.textLength,
    textNodeCount: partProfile.textNodeCount,
    paragraphCount: partProfile.paragraphs.length,
    tableCount: partProfile.tables.length,
    sectionCount: partProfile.sections.length,
    fingerprint: partProfile.fingerprint
  };
}

function groupBy(items, keyFn) {
  const groups = new Map();
  for (const item of items) {
    const key = keyFn(item);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(item);
  }
  return groups;
}

function detectProfileSignals(profile) {
  const signals = [];
  const paragraphs = profile.parts.flatMap((part) => part.paragraphs.map((paragraph) => ({
    ...paragraph,
    partName: part.name
  })));
  const styleGroups = groupBy(paragraphs.filter((paragraph) => paragraph.styleId), (paragraph) => paragraph.styleId);
  for (const [styleId, items] of styleGroups) {
    const variants = groupBy(items, (paragraph) => paragraph.paragraphPropertiesFingerprint);
    if (variants.size > 1) {
      signals.push({
        type: "paragraph-style-direct-format-variants",
        severity: variants.size > 3 ? "medium" : "low",
        message: `Paragraph style ${styleId} appears with ${variants.size} direct-format variants.`,
        styleId,
        variantCount: variants.size,
        evidence: Array.from(variants.values()).slice(0, 5).map((group) => ({
          count: group.length,
          sample: {
            part: group[0].partName,
            paragraphIndex: group[0].index,
            textPreview: group[0].text.slice(0, 80),
            fingerprint: group[0].paragraphPropertiesFingerprint
          }
        }))
      });
    }
  }

  const tableProfiles = profile.parts.flatMap((part) => part.tables.map((table) => ({
    ...table,
    partName: part.name
  })));
  for (const table of tableProfiles) {
    const rowFingerprints = new Set(table.rows.slice(1).map((row) => row.formatFingerprint));
    if (table.rows.length >= 4 && rowFingerprints.size > 1) {
      signals.push({
        type: "table-row-format-variants",
        severity: "low",
        message: `Table ${table.id} has ${rowFingerprints.size} body-row format variants.`,
        tableId: table.id,
        evidence: {
          part: table.partName,
          tableIndex: table.index,
          rowCount: table.rowCount
        }
      });
    }
  }

  const directRunFormattingCount = paragraphs.reduce((count, paragraph) => (
    count + paragraph.runs.filter((run) => run.rPr).length
  ), 0);
  if (directRunFormattingCount > Math.max(20, paragraphs.length)) {
    signals.push({
      type: "heavy-direct-run-formatting",
      severity: "medium",
      message: "The document relies heavily on direct run formatting; generalized templates should normalize repeated atoms before inducing rules.",
      directRunFormattingCount
    });
  }

  return signals;
}

function makeEvidence(sourceId, details) {
  return {
    sourceId,
    ...details
  };
}

function addFormatAtom(atomMap, kind, fingerprint, properties, evidence) {
  const key = `${kind}:${fingerprint}`;
  if (!atomMap.has(key)) {
    atomMap.set(key, {
      id: `fmt.${kind}.${String(atomMap.size + 1).padStart(4, "0")}`,
      kind,
      fingerprint,
      properties: properties || null,
      evidence: [],
      stability: "unknown"
    });
  }
  atomMap.get(key).evidence.push(evidence);
  return atomMap.get(key).id;
}

function blockSignature(node) {
  return `${node.kind}:${node.part || ""}:${node.index ?? ""}`;
}

function textClass(text) {
  const value = String(text || "").trim();
  if (!value) return "empty";
  if (/^\d{4}[-/.\u5e74]\d{1,2}([-/.\u6708]\d{1,2}\u65e5?)?$/u.test(value)) return "date";
  if (/^[\d,]+(\.\d+)?%?$/u.test(value)) return "number";
  if (valueHasLabel(value)) return "labelValue";
  if (value.length <= 12) return "shortText";
  return "longText";
}

function valueHasLabel(value) {
  return /[:\uFF1A]/u.test(value) && value.replace(/[:\uFF1A].*$/u, "").trim().length <= 20;
}

function splitLabelValue(value) {
  const match = String(value || "").match(/^(.{1,30}?)[\uFF1A:]\s*(.+)$/u);
  if (!match) return null;
  return {
    label: match[1].trim(),
    value: match[2].trim()
  };
}

function normalizeAnchorText(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/\s+/gu, "")
    .replace(/[^\p{L}\p{N}]+/gu, "");
}

function textBigrams(text) {
  const normalized = normalizeAnchorText(text);
  if (!normalized) return new Set();
  if (normalized.length === 1) return new Set([normalized]);
  const grams = new Set();
  for (let index = 0; index < normalized.length - 1; index += 1) {
    grams.add(normalized.slice(index, index + 2));
  }
  return grams;
}

function jaccard(left, right) {
  if (!left.size && !right.size) return 0;
  let intersection = 0;
  for (const item of left) {
    if (right.has(item)) intersection += 1;
  }
  const union = new Set([...left, ...right]).size;
  return union ? intersection / union : 0;
}

function textSimilarity(left, right) {
  const leftText = normalizeAnchorText(left);
  const rightText = normalizeAnchorText(right);
  if (!leftText || !rightText) return 0;
  if (leftText === rightText) return 1;
  return jaccard(textBigrams(leftText), textBigrams(rightText));
}

function tableHeaderText(table) {
  const firstRow = table.rows[0];
  if (!firstRow) return "";
  return firstRow.cells.map((cell) => cell.text.trim()).join("|");
}

function indexProximity(left, right) {
  const distance = Math.abs((left.index ?? 0) - (right.index ?? 0));
  return Math.max(0, 1 - distance / 20);
}

function blockAnchorKey(block) {
  if (block.kind === "table") {
    const header = normalizeAnchorText(block.headerText);
    return header ? `table:${block.part}:${header}` : `table:${block.part}:${block.columnCount}`;
  }
  const labelValue = splitLabelValue(block.text);
  if (labelValue) return `paragraph:${block.part}:label:${normalizeAnchorText(labelValue.label)}`;
  const normalized = normalizeAnchorText(block.text);
  if (normalized && normalized.length <= 40) return `paragraph:${block.part}:text:${normalized}`;
  if (block.styleId) return `paragraph:${block.part}:style:${block.styleId}`;
  return `paragraph:${block.part}:index:${block.index}`;
}

function blockSimilarity(left, right) {
  if (left.kind !== right.kind) return { score: 0, reason: "kind-mismatch" };
  let score = left.part === right.part ? 0.08 : 0;
  const reasons = [];

  if (left.kind === "table") {
    const headerScore = textSimilarity(left.headerText, right.headerText);
    if (headerScore > 0) {
      score += headerScore * 0.58;
      reasons.push(`header:${headerScore.toFixed(2)}`);
    }
    if (left.columnCount && left.columnCount === right.columnCount) {
      score += 0.16;
      reasons.push("columns");
    }
    if (left.formatFingerprint === right.formatFingerprint) {
      score += 0.12;
      reasons.push("format");
    }
    score += indexProximity(left, right) * 0.08;
    if (Math.abs((left.rowCount || 0) - (right.rowCount || 0)) <= 1) {
      score += 0.06;
      reasons.push("row-count");
    }
    return { score: Math.min(score, 1), reason: reasons.join(",") || "weak-table" };
  }

  const leftLabel = splitLabelValue(left.text);
  const rightLabel = splitLabelValue(right.text);
  if (leftLabel && rightLabel) {
    const labelScore = textSimilarity(leftLabel.label, rightLabel.label);
    score += labelScore * 0.62;
    if (labelScore > 0) reasons.push(`label:${labelScore.toFixed(2)}`);
  } else {
    const bodyScore = textSimilarity(left.text, right.text);
    score += bodyScore * 0.42;
    if (bodyScore > 0) reasons.push(`text:${bodyScore.toFixed(2)}`);
  }
  if (left.styleId && left.styleId === right.styleId) {
    score += 0.14;
    reasons.push("style");
  }
  if (left.formatFingerprint === right.formatFingerprint) {
    score += 0.14;
    reasons.push("format");
  }
  if (inferRoleCandidates(left).some((role) => inferRoleCandidates(right).includes(role))) {
    score += 0.08;
    reasons.push("role");
  }
  score += indexProximity(left, right) * 0.08;

  return { score: Math.min(score, 1), reason: reasons.join(",") || "weak-paragraph" };
}

function alignmentThreshold(block) {
  if (block.kind === "table") return 0.48;
  if (splitLabelValue(block.text)) return 0.5;
  const normalized = normalizeAnchorText(block.text);
  if (!normalized) return 0.72;
  if (normalized.length <= 20) return 0.56;
  return 0.62;
}

function collectDocumentBlocks(profile, sourceId) {
  const blocks = [];
  for (const part of profile.parts) {
    for (const paragraph of part.paragraphs) {
      blocks.push({
        id: `${sourceId}.p.${String(blocks.length + 1).padStart(4, "0")}`,
        kind: "paragraph",
        sourceId,
        part: part.name,
        index: paragraph.index,
        text: paragraph.text,
        textPreview: paragraph.text.slice(0, 120),
        structureFingerprint: paragraph.formatFingerprint,
        formatFingerprint: paragraph.paragraphPropertiesFingerprint,
        runFormatFingerprints: paragraph.runs.map((run) => run.formatFingerprint),
        styleId: paragraph.styleId,
        anchorKey: null,
        evidence: makeEvidence(sourceId, {
          part: part.name,
          paragraphIndex: paragraph.index,
          path: paragraph.evidence.path,
          textPreview: paragraph.text.slice(0, 120),
          fingerprint: paragraph.formatFingerprint
        }),
        original: paragraph
      });
    }
    for (const table of part.tables) {
      blocks.push({
        id: `${sourceId}.tbl.${String(blocks.length + 1).padStart(4, "0")}`,
        kind: "table",
        sourceId,
        part: part.name,
        index: table.index,
        text: table.rows.map((row) => row.text).join("\n"),
        textPreview: table.rows.slice(0, 3).map((row) => row.text).join(" | ").slice(0, 120),
        headerText: tableHeaderText(table),
        structureFingerprint: table.formatFingerprint,
        formatFingerprint: table.tablePropertiesFingerprint,
        rowCount: table.rowCount,
        columnCount: table.columnCount,
        anchorKey: null,
        evidence: makeEvidence(sourceId, {
          part: part.name,
          tableIndex: table.index,
          path: table.evidence.path,
          textPreview: table.rows.slice(0, 3).map((row) => row.text).join(" | ").slice(0, 120),
          fingerprint: table.formatFingerprint
        }),
        original: table
      });
    }
  }
  return blocks.sort((left, right) => {
    if (left.part !== right.part) return left.part.localeCompare(right.part);
    return left.index - right.index;
  }).map((block) => ({
    ...block,
    anchorKey: blockAnchorKey(block)
  }));
}

function buildProfileSource(inputDocx, profile, index) {
  return {
    id: `src${String(index + 1).padStart(3, "0")}`,
    docxPath: path.resolve(inputDocx),
    role: index === 0 ? "baseline" : "example",
    fingerprint: profile.package.fingerprint,
    profile
  };
}

function inferStructureAndAtoms(sources, atomMap) {
  const structure = [];
  for (const source of sources) {
    const blocks = collectDocumentBlocks(source.profile, source.id);
    for (const block of blocks) {
      const formatAtomRefs = [];
      if (block.kind === "paragraph") {
        formatAtomRefs.push(addFormatAtom(
          atomMap,
          "paragraph",
          block.original.paragraphPropertiesFingerprint,
          block.original.pPr,
          block.evidence
        ));
        for (const run of block.original.runs) {
          formatAtomRefs.push(addFormatAtom(
            atomMap,
            "run",
            run.formatFingerprint,
            run.rPr,
            makeEvidence(source.id, {
              part: block.part,
              paragraphIndex: block.index,
              runIndex: run.index,
              path: run.evidence.path,
              textPreview: run.text.slice(0, 80),
              fingerprint: run.formatFingerprint
            })
          ));
        }
      } else if (block.kind === "table") {
        formatAtomRefs.push(addFormatAtom(
          atomMap,
          "table",
          block.original.tablePropertiesFingerprint,
          { tblPr: block.original.tblPr, tblGrid: block.original.tblGrid },
          block.evidence
        ));
        for (const row of block.original.rows) {
          formatAtomRefs.push(addFormatAtom(
            atomMap,
            "row",
            row.rowPropertiesFingerprint,
            row.trPr,
            makeEvidence(source.id, {
              part: block.part,
              tableIndex: block.index,
              rowIndex: row.index,
              textPreview: row.text.slice(0, 80),
              fingerprint: row.rowPropertiesFingerprint
            })
          ));
          for (const cell of row.cells) {
            formatAtomRefs.push(addFormatAtom(
              atomMap,
              "cell",
              cell.cellPropertiesFingerprint,
              cell.tcPr,
              makeEvidence(source.id, {
                part: block.part,
                tableIndex: block.index,
                rowIndex: row.index,
                cellIndex: cell.index,
                textPreview: cell.text.slice(0, 80),
                fingerprint: cell.cellPropertiesFingerprint
              })
            ));
          }
        }
      }

      structure.push({
        id: `node.${String(structure.length + 1).padStart(4, "0")}`,
        kind: block.kind,
        sourceRefs: [block.evidence],
        formatAtomRefs: Array.from(new Set(formatAtomRefs)),
        textPreview: block.textPreview,
        roleCandidates: inferRoleCandidates(block),
        signature: blockSignature(block),
        children: []
      });
    }
  }
  return structure;
}

function inferRoleCandidates(block) {
  if (block.kind === "table") {
    const candidates = ["table"];
    if (block.rowCount > 2) candidates.push("loopCandidate");
    return candidates;
  }
  const text = String(block.text || "").trim();
  if (!text) return ["empty"];
  const candidates = [];
  if (valueHasLabel(text)) candidates.push("labelValue");
  if (/^[\u4E00\u4E8C\u4E09\u56DB\u4E94\u516D\u4E03\u516B\u4E5D\u5341]+[\u3001.．]/u.test(text) || /^\d+[.．、]/u.test(text)) candidates.push("numberedHeading");
  if (text.length <= 30 && block.original.runs.some((run) => run.rPr?.["w:b"] !== undefined)) candidates.push("heading");
  if (text.length > 80) candidates.push("body");
  if (!candidates.length) candidates.push(textClass(text));
  return candidates;
}

function groupBlocksForRules(sources) {
  if (sources.length <= 1) {
    const byPosition = new Map();
    for (const source of sources) {
      for (const block of collectDocumentBlocks(source.profile, source.id)) {
        const key = blockSignature(block);
        if (!byPosition.has(key)) {
          byPosition.set(key, {
            signature: key,
            anchorKey: block.anchorKey,
            blocks: [],
            alignment: { strategy: "single-source", score: 1, reason: "single-source", pairs: [] }
          });
        }
        byPosition.get(key).blocks.push(block);
      }
    }
    return Array.from(byPosition.values());
  }

  return alignBlocksAcrossSources(sources);
}

function alignBlocksAcrossSources(sources) {
  const sourceEntries = sources.map((source) => ({
    source,
    blocks: collectDocumentBlocks(source.profile, source.id)
  }));
  const baseline = sourceEntries[0];
  const remainingBySource = new Map(sourceEntries.slice(1).map((entry) => [entry.source.id, new Set(entry.blocks)]));
  const groups = [];

  for (const baseBlock of baseline.blocks) {
    const group = {
      signature: `align:${String(groups.length + 1).padStart(4, "0")}`,
      anchorKey: baseBlock.anchorKey,
      blocks: [baseBlock],
      alignment: {
        strategy: "anchor-similarity",
        score: 1,
        reason: "baseline",
        pairs: []
      }
    };

    for (const entry of sourceEntries.slice(1)) {
      const remaining = remainingBySource.get(entry.source.id);
      let best = null;
      for (const candidate of remaining) {
        if (candidate.kind !== baseBlock.kind) continue;
        const similarity = blockSimilarity(baseBlock, candidate);
        const sameAnchor = candidate.anchorKey === baseBlock.anchorKey;
        const score = sameAnchor ? Math.min(1, similarity.score + 0.25) : similarity.score;
        if (!best || score > best.score) {
          best = {
            block: candidate,
            score,
            reason: `${sameAnchor ? "anchor," : ""}${similarity.reason}`
          };
        }
      }

      const threshold = alignmentThreshold(baseBlock);
      if (best && best.score >= threshold) {
        remaining.delete(best.block);
        group.blocks.push(best.block);
        group.alignment.pairs.push({
          sourceId: entry.source.id,
          score: Number(best.score.toFixed(3)),
          threshold,
          reason: best.reason
        });
      }
    }

    if (group.alignment.pairs.length) {
      const total = group.alignment.pairs.reduce((sum, pair) => sum + pair.score, 0);
      group.alignment.score = Number((total / group.alignment.pairs.length).toFixed(3));
      group.alignment.reason = group.alignment.pairs.map((pair) => `${pair.sourceId}:${pair.reason}`).join("; ");
    } else {
      group.alignment.score = 0;
      group.alignment.reason = "no-match";
    }
    groups.push(group);
  }

  for (const entry of sourceEntries.slice(1)) {
    for (const block of remainingBySource.get(entry.source.id)) {
      groups.push({
        signature: `align:${String(groups.length + 1).padStart(4, "0")}`,
        anchorKey: block.anchorKey,
        blocks: [block],
        alignment: {
          strategy: "unmatched",
          score: 0,
          reason: "no-anchor-match",
          pairs: []
        }
      });
    }
  }

  return groups;
}

function inferRulesAndConflicts(sources, alignedGroups = null) {
  const rules = [];
  const conflicts = [];
  const dataFields = {};
  const arrays = {};
  const groups = alignedGroups || groupBlocksForRules(sources);
  let fieldIndex = 1;
  let arrayIndex = 1;

  for (const group of groups) {
    const { signature, blocks, alignment } = group;
    const sourceCount = new Set(blocks.map((block) => block.sourceId)).size;
    const texts = new Set(blocks.map((block) => block.text));
    const formats = new Set(blocks.map((block) => block.formatFingerprint));
    const first = blocks[0];
    const evidences = blocks.map((block) => block.evidence);
    const alignedAllSources = sourceCount === sources.length;
    const trustedAlignment = sources.length <= 1 || (alignedAllSources && alignment.score >= alignmentThreshold(first));

    if (sourceCount < sources.length) {
      conflicts.push({
        id: `conflict.${String(conflicts.length + 1).padStart(4, "0")}`,
        type: "optionalBlockAmbiguity",
        severity: "medium",
        classification: "unresolvedConflict",
        evidence: evidences,
        recommendedAction: "Confirm whether this block is optional or missing because samples are not structurally aligned.",
        signature,
        anchorKey: group.anchorKey,
        alignment,
        presentIn: sourceCount,
        expectedSources: sources.length
      });
    }

    if (first.kind === "paragraph") {
      const labelValue = blocks.map((block) => splitLabelValue(block.text));
      const hasConsistentLabel = labelValue.every(Boolean)
        && new Set(labelValue.map((item) => item.label)).size === 1;
      const valueSet = new Set(labelValue.filter(Boolean).map((item) => item.value));

      if (trustedAlignment && hasConsistentLabel && valueSet.size > 1) {
        const fieldName = semanticFieldName(labelValue[0].label, fieldIndex);
        fieldIndex += 1;
        dataFields[fieldName] = {
          type: inferFieldType(Array.from(valueSet)),
          description: `Value after label ${labelValue[0].label}`,
          examples: Array.from(valueSet).slice(0, 5),
          confidence: 0.72
        };
        rules.push({
          id: `rule.${String(rules.length + 1).padStart(4, "0")}`,
          kind: "slot",
          target: signature,
          expression: fieldName,
          confidence: 0.72,
          evidence: evidences,
          alternatives: [{ kind: "staticText", reason: "Could be fixed if samples refer to different reports rather than one template." }],
          validationStatus: "unverified",
          label: labelValue[0].label,
          alignment
        });
      } else if (texts.size === 1 && first.text.trim()) {
        rules.push({
          id: `rule.${String(rules.length + 1).padStart(4, "0")}`,
          kind: "staticText",
          target: signature,
          value: first.text,
          confidence: sourceCount === sources.length ? 0.86 : 0.55,
          evidence: evidences,
          alternatives: [],
          validationStatus: "unverified",
          alignment
        });
      } else if (trustedAlignment && texts.size > 1 && sourceCount > 1) {
        const fieldName = `field.${String(fieldIndex).padStart(3, "0")}`;
        fieldIndex += 1;
        dataFields[fieldName] = {
          type: inferFieldType(Array.from(texts)),
          description: `Variable paragraph at ${signature}`,
          examples: Array.from(texts).slice(0, 5),
          confidence: 0.48
        };
        rules.push({
          id: `rule.${String(rules.length + 1).padStart(4, "0")}`,
          kind: "slot",
          target: signature,
          expression: fieldName,
          confidence: 0.48,
          evidence: evidences,
          alternatives: [{ kind: "conditional", reason: "Paragraph may represent different optional content rather than one scalar field." }],
          validationStatus: "unverified",
          alignment
        });
      } else if (texts.size > 1 && sourceCount > 1) {
        conflicts.push({
          id: `conflict.${String(conflicts.length + 1).padStart(4, "0")}`,
          type: "optionalBlockAmbiguity",
          severity: "medium",
          classification: "unresolvedConflict",
          evidence: evidences,
          recommendedAction: "Text differs but alignment confidence is too low to infer a scalar slot.",
          signature,
          anchorKey: group.anchorKey,
          alignment
        });
      }
    }

    if (first.kind === "table") {
      const rowCounts = new Set(blocks.map((block) => block.rowCount));
      if ((trustedAlignment && rowCounts.size > 1) || (sourceCount === 1 && first.rowCount > 2)) {
        const arrayName = `table${String(arrayIndex).padStart(3, "0")}.rows`;
        arrayIndex += 1;
        arrays[arrayName] = {
          type: "array",
          description: `Rows inferred from table at ${signature}`,
          examples: blocks.map((block) => ({
            rowCount: block.rowCount,
            columnCount: block.columnCount,
            preview: block.textPreview
          })),
          confidence: rowCounts.size > 1 ? 0.68 : 0.42
        };
        rules.push({
          id: `rule.${String(rules.length + 1).padStart(4, "0")}`,
          kind: "loop",
          target: signature,
          expression: arrayName,
          confidence: rowCounts.size > 1 ? 0.68 : 0.42,
          evidence: evidences,
          alternatives: [{ kind: "staticTable", reason: "Repeated rows may be fixed report layout if row count does not change across samples." }],
          validationStatus: "unverified",
          alignment
        });
      } else if (rowCounts.size > 1 && !trustedAlignment) {
        conflicts.push({
          id: `conflict.${String(conflicts.length + 1).padStart(4, "0")}`,
          type: "tableShapeAmbiguity",
          severity: "medium",
          classification: "unresolvedConflict",
          evidence: evidences,
          recommendedAction: "Table row counts differ but table headers/anchors are not similar enough to infer a loop.",
          signature,
          anchorKey: group.anchorKey,
          alignment
        });
      }
    }

    if (formats.size > 1) {
      conflicts.push({
        id: `conflict.${String(conflicts.length + 1).padStart(4, "0")}`,
        type: first.kind === "table" ? "tableShapeAmbiguity" : "sameRoleDifferentFormat",
        severity: "medium",
        classification: "unresolvedConflict",
        evidence: evidences,
        recommendedAction: "Review whether the differing format is a condition, a sample anomaly, or direct-format drift.",
        signature,
        anchorKey: group.anchorKey,
        alignment,
        formatVariantCount: formats.size
      });
    }
  }

  return { rules, conflicts, dataFields, arrays };
}

function semanticFieldName(label, fallbackIndex) {
  const normalized = String(label || "")
    .replace(/\s+/gu, "")
    .replace(/[^\p{L}\p{N}]+/gu, "");
  const dictionary = [
    [/(company|\u516C\u53F8|\u4F01\u4E1A|\u5355\u4F4D)/iu, "company.name"],
    [/(date|\u65E5\u671F|\u65F6\u95F4)/iu, "report.date"],
    [/(name|owner|person|\u59D3\u540D|\u8D1F\u8D23\u4EBA|\u8054\u7CFB\u4EBA)/iu, "person.name"],
    [/(amount|\u91D1\u989D|\u4EF7\u6B3E|\u603B\u989D)/iu, "amount"],
    [/(address|\u5730\u5740|\u4F4F\u6240)/iu, "address"]
  ];
  for (const [pattern, name] of dictionary) {
    if (pattern.test(normalized)) return name;
  }
  return `field.${String(fallbackIndex).padStart(3, "0")}`;
}

function inferFieldType(values) {
  if (values.every((value) => textClass(value) === "date")) return "date";
  if (values.every((value) => textClass(value) === "number")) return "number";
  return "string";
}

function finalizeFormatAtomStability(formatAtoms, sourcesLength) {
  return formatAtoms.map((atom) => {
    const sourceIds = new Set(atom.evidence.map((item) => item.sourceId));
    return {
      ...atom,
      stability: sourceIds.size === sourcesLength && atom.evidence.length >= sourcesLength ? "stable" : "variant"
    };
  });
}

function summarizeAlignment(groups, sourcesLength) {
  if (sourcesLength <= 1) {
    return {
      strategy: "single-source",
      groups: groups.length,
      matchedGroups: groups.length,
      unmatchedGroups: 0,
      averageScore: 1,
      lowConfidenceGroups: 0
    };
  }

  const matched = groups.filter((group) => new Set(group.blocks.map((block) => block.sourceId)).size === sourcesLength);
  const scores = matched.map((group) => group.alignment?.score || 0);
  const averageScore = scores.length
    ? scores.reduce((sum, score) => sum + score, 0) / scores.length
    : 0;
  return {
    strategy: "anchor-similarity",
    groups: groups.length,
    matchedGroups: matched.length,
    unmatchedGroups: groups.length - matched.length,
    averageScore: Number(averageScore.toFixed(3)),
    lowConfidenceGroups: matched.filter((group) => (group.alignment?.score || 0) < alignmentThreshold(group.blocks[0])).length
  };
}

async function analyzeDocxInputs(inputDocxPaths) {
  const sources = [];
  for (let index = 0; index < inputDocxPaths.length; index += 1) {
    const zip = await loadDocx(inputDocxPaths[index]);
    const profile = await profileZip(zip);
    sources.push(buildProfileSource(inputDocxPaths[index], profile, index));
  }

  const atomMap = new Map();
  const structure = inferStructureAndAtoms(sources, atomMap);
  const groups = groupBlocksForRules(sources);
  const { rules, conflicts, dataFields, arrays } = inferRulesAndConflicts(sources, groups);
  const profileSignals = sources.flatMap((source) => source.profile.signals.map((signal) => ({
    ...signal,
    sourceId: source.id
  })));

  const ir = {
    version: 1,
    mode: inputDocxPaths.length > 1 ? "generalize" : "audit",
    generatedAt: new Date().toISOString(),
    sources: sources.map((source) => ({
      id: source.id,
      docxPath: source.docxPath,
      role: source.role,
      fingerprint: source.fingerprint
    })),
    formatAtoms: finalizeFormatAtomStability(Array.from(atomMap.values()), sources.length),
    structure,
    dataSchema: {
      fields: dataFields,
      arrays,
      objects: {},
      required: []
    },
    rules,
    conflicts,
    validation: [],
    profileSignals,
    summary: {
      sourceCount: sources.length,
      structureNodes: structure.length,
      formatAtoms: atomMap.size,
      rules: rules.length,
      conflicts: conflicts.length,
      fields: Object.keys(dataFields).length,
      arrays: Object.keys(arrays).length,
      profileSignals: profileSignals.length
    },
    alignment: summarizeAlignment(groups, sources.length)
  };

  return ir;
}

function officeCapabilityForRule(rule) {
  if (rule.kind === "slot") {
    return {
      native: "contentControl",
      binding: "customXmlPart",
      patchKind: "wrap-range-with-sdt",
      confidence: Math.min(rule.confidence ?? 0.5, 0.9),
      rationale: "Scalar dynamic text maps to a Word content control bound to a custom XML field."
    };
  }
  if (rule.kind === "loop") {
    return {
      native: "repeatingSectionContentControl",
      binding: "customXmlPart",
      patchKind: "wrap-row-or-block-with-repeating-sdt",
      confidence: Math.min(rule.confidence ?? 0.4, 0.78),
      rationale: "Repeated rows or block groups should use Word repeating section content controls when the range is structurally stable."
    };
  }
  if (rule.kind === "staticText") {
    return {
      native: "plainWordContent",
      binding: null,
      patchKind: "preserve-existing-ooxml",
      confidence: rule.confidence ?? 0.8,
      rationale: "Static text should be left as native Word content, preserving existing OOXML and styles."
    };
  }
  if (rule.kind === "conditional") {
    return {
      native: "contentControlGroup",
      binding: "sidecarCondition",
      patchKind: "annotate-sdt-with-sidecar-condition",
      confidence: rule.confidence ?? 0.45,
      rationale: "Word has no first-class conditional template primitive; use a native content control as the range marker and a sidecar expression for rendering."
    };
  }
  return {
    native: "sidecarRule",
    binding: "sidecar",
    patchKind: "manual-or-renderer-rule",
    confidence: rule.confidence ?? 0.3,
    rationale: "No direct Office-native mapping is known for this rule kind."
  };
}

function officeFieldPath(rule, index) {
  if (rule.expression && /^[A-Za-z_][A-Za-z0-9_.[\]-]*$/u.test(rule.expression)) {
    return rule.expression;
  }
  return `field.${String(index).padStart(3, "0")}`;
}

function contentControlTag(prefix, fieldPath) {
  return `${prefix}:${fieldPath}`.replace(/\s+/gu, "_");
}

function compileReadinessForRule(rule) {
  if (rule.kind === "slot") {
    const requires = rule.label ? "value-run-range" : "paragraph-or-run-range";
    return {
      status: "needs-exact-range",
      requiredTarget: requires,
      deterministicBackend: "officecli-ooxml-patcher",
      reason: rule.label
        ? "The current IR identifies the label/value paragraph, but a native content control should wrap only the dynamic value after the label."
        : "The current IR identifies the paragraph block; direct DOCX patching needs exact run/text-node boundaries before wrapping content."
    };
  }
  if (rule.kind === "loop") {
    return {
      status: "needs-exact-range",
      requiredTarget: "table-row-or-block-range",
      deterministicBackend: "officecli-ooxml-patcher",
      reason: "Repeating section content controls must wrap whole paragraphs or table rows with validated boundaries."
    };
  }
  if (rule.kind === "conditional") {
    return {
      status: "needs-rule-and-range",
      requiredTarget: "paragraph-or-block-range",
      deterministicBackend: "officecli-ooxml-patcher-with-sidecar",
      reason: "Word has no native conditional expression; the native control can mark the range, while sidecar metadata decides keep/delete during rendering."
    };
  }
  return {
    status: "not-patchable",
    requiredTarget: null,
    deterministicBackend: "manual-review",
    reason: "This rule kind is not mapped to an automatic Office-native patch."
  };
}

function buildOfficePatch(rule, index) {
  const capability = officeCapabilityForRule(rule);
  const fieldPath = officeFieldPath(rule, index);
  const base = {
    id: `patch.${String(index).padStart(4, "0")}`,
    ruleId: rule.id,
    ruleKind: rule.kind,
    target: rule.target,
    label: rule.label || null,
    expression: rule.expression || null,
    evidence: rule.evidence,
    confidence: capability.confidence,
    officeNative: capability.native,
    binding: capability.binding,
    patchKind: capability.patchKind,
    compileReadiness: compileReadinessForRule(rule),
    rationale: capability.rationale
  };

  if (rule.kind === "slot") {
    return {
      ...base,
      contentControl: {
        type: "plainText",
        title: fieldPath,
        tag: contentControlTag("ot-field", fieldPath),
        lockContentControl: false,
        lockContents: false
      },
      customXmlBinding: {
        storeItemId: OFFICE_TEMPLATE_STORE_ITEM_ID,
        xpath: `/template/data/${fieldPath.replace(/[.[\]-]+/gu, "/")}`,
        prefixMappings: ""
      },
      sidecar: null
    };
  }

  if (rule.kind === "loop") {
    return {
      ...base,
      contentControl: {
        type: "repeatingSection",
        title: fieldPath,
        tag: contentControlTag("ot-repeat", fieldPath),
        lockContentControl: false,
        lockContents: false
      },
      customXmlBinding: {
        storeItemId: OFFICE_TEMPLATE_STORE_ITEM_ID,
        xpath: `/template/data/${fieldPath.replace(/[.[\]-]+/gu, "/")}`,
        prefixMappings: ""
      },
      sidecar: {
        rendererRequired: true,
        reason: "Repeating section creation and row cloning require range-safe compilation before direct DOCX patching."
      }
    };
  }

  if (rule.kind === "conditional") {
    return {
      ...base,
      contentControl: {
        type: "richText",
        title: fieldPath,
        tag: contentControlTag("ot-condition", fieldPath),
        lockContentControl: false,
        lockContents: false
      },
      customXmlBinding: null,
      sidecar: {
        condition: rule.expression || null,
        rendererRequired: true
      }
    };
  }

  return {
    ...base,
    contentControl: null,
    customXmlBinding: null,
    sidecar: { rendererRequired: true }
  };
}

function buildOfficePreserveItem(rule, index) {
  const capability = officeCapabilityForRule(rule);
  return {
    id: `preserve.${String(index).padStart(4, "0")}`,
    ruleId: rule.id,
    ruleKind: rule.kind,
    target: rule.target,
    value: rule.value,
    evidence: rule.evidence,
    confidence: capability.confidence,
    officeNative: capability.native,
    patchKind: capability.patchKind,
    rationale: capability.rationale
  };
}

function conflictToOfficeGap(conflict, index) {
  return {
    id: `gap.${String(index).padStart(4, "0")}`,
    conflictId: conflict.id,
    type: conflict.type,
    severity: conflict.severity,
    classification: conflict.classification,
    evidence: conflict.evidence,
    recommendedAction: conflict.recommendedAction,
    officeNativeFallback: conflict.type === "optionalBlockAmbiguity"
      ? "contentControlGroup + sidecar condition"
      : "preserve existing OOXML until conflict is classified",
    rationale: "Office-native templates require an explicit range and stable semantics; unresolved conflicts stay out of automatic patching."
  };
}

function buildOfficeTemplatePlan(ir) {
  const rules = Array.isArray(ir.rules) ? ir.rules : [];
  const conflicts = Array.isArray(ir.conflicts) ? ir.conflicts : [];
  const patchableRules = rules.filter((rule) => ["slot", "loop", "conditional"].includes(rule.kind));
  const preserveRules = rules.filter((rule) => rule.kind === "staticText");
  const patches = patchableRules
    .map((rule, index) => buildOfficePatch(rule, index + 1));
  const preserves = preserveRules
    .map((rule, index) => buildOfficePreserveItem(rule, index + 1));
  const gaps = conflicts.map((conflict, index) => conflictToOfficeGap(conflict, index + 1));
  const officeNativePatchCount = patches.filter((patch) => patch.officeNative !== "sidecarRule").length;
  const sidecarRendererRuleCount = patches.filter((patch) => patch.sidecar?.rendererRequired).length;
  const sidecarGapCount = gaps.length;
  const sidecarTotalCount = sidecarRendererRuleCount + sidecarGapCount;

  return {
    version: 1,
    kind: "office-native-template-plan",
    generatedAt: new Date().toISOString(),
    sourceIr: {
      mode: ir.mode,
      sources: ir.sources || [],
      summary: ir.summary || null,
      alignment: ir.alignment || null
    },
    principles: [
      "Preserve existing DOCX OOXML, styles, numbering, sections, headers, footers, and table formatting unless a patch explicitly wraps a range.",
      "Prefer Word content controls and custom XML bindings for scalar slots.",
      "Prefer repeating section content controls for stable row or block loops.",
      "Represent conditions and unresolved ambiguity as sidecar metadata attached to native content-control ranges.",
      "Never rewrite formatting rules into a private styling system when Word styles or direct OOXML can be preserved."
    ],
    nativeTargets: {
      scalarSlots: "w:sdt plain-text/rich-text content controls with w:tag and optional custom XML binding",
      loops: "Word repeating section content controls where a stable table row or block range is known",
      staticContent: "Unmodified WordprocessingML",
      styles: "Existing styles.xml, numbering.xml, table properties, paragraph/run properties",
      gaps: "Sidecar JSON metadata plus conflict report"
    },
    backendDecision: {
      compile: "officecli-ooxml-patcher",
      render: "officecli-ooxml-patcher + Office-native content controls/custom XML where possible",
      verify: "template.compare-format first; COM/JSAPI only for live Office/WPS layout, field refresh, save-as, export, or active document workflows",
      notForTemplateRendering: ["wps-uia"],
      rationale: "Template compilation and rendering are deterministic package edits, not visible desktop UI automation. UIA remains reserved for product UI capabilities such as PDF conversion."
    },
    patches,
    preserves,
    gaps,
    summary: {
      patches: patches.length,
      actualPatchCount: patches.length,
      officeNativePatchCount,
      nativePatchCount: officeNativePatchCount,
      sidecarRendererRuleCount,
      sidecarGapCount,
      sidecarTotalCount,
      sidecarPatchCount: sidecarTotalCount,
      gaps: gaps.length,
      slotPatches: patches.filter((patch) => patch.ruleKind === "slot").length,
      loopPatches: patches.filter((patch) => patch.ruleKind === "loop").length,
      conditionalPatches: patches.filter((patch) => patch.ruleKind === "conditional").length,
      staticPreserveItems: preserves.length,
      staticPreservePatches: preserves.length
    }
  };
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
  if (typeof node === "string") return "";
  if (typeof node !== "object") return "";
  if (Array.isArray(node)) return node.map(extractTextFromParsed).join("");
  let out = "";
  for (const [key, value] of Object.entries(node)) {
    if (PARSED_TEXT_KEYS.has(key)) {
      if (typeof value === "string") {
        out += value;
      } else if (Array.isArray(value)) {
        out += value.map((item) => (typeof item === "string" ? item : item?.["#text"] || "")).join("");
      } else if (value && typeof value === "object") {
        out += value["#text"] || "";
      }
    } else {
      out += extractTextFromParsed(value);
    }
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

async function profileZip(zip) {
  const parts = await readWordParts(zip);
  const styleXml = zip.file("word/styles.xml") ? await zip.file("word/styles.xml").async("string") : "";
  const numberingXml = zip.file("word/numbering.xml") ? await zip.file("word/numbering.xml").async("string") : "";
  const profileParts = [];

  for (const part of parts) {
    const parsed = parseXmlPart(part.xml);
    const paragraphs = collectDetailedParagraphs(parsed, part.name);
    const tables = collectDetailedTables(parsed, part.name);
    const sections = collectSections(parsed, part.name);
    profileParts.push({
      name: part.name,
      textLength: textFromXml(part.xml).length,
      textNodeCount: collectTextNodes(part.xml, part.name).length,
      paragraphs,
      tables,
      sections,
      fingerprint: hashFingerprint({
        name: part.name,
        structure: normalizeXml(part.xml).replace(TEXT_NODE_RE, "<w:t/>")
      })
    });
  }

  const profile = {
    version: 1,
    kind: "docx-template-profile",
    generatedAt: new Date().toISOString(),
    package: {
      partCount: Object.values(zip.files).filter((file) => !file.dir).length,
      wordTextParts: parts.map((part) => part.name),
      fingerprint: hashFingerprint({
        parts: parts.map((part) => ({
          name: part.name,
          structure: normalizeXml(part.xml).replace(TEXT_NODE_RE, "<w:t/>")
        })),
        styles: normalizeXml(styleXml),
        numbering: normalizeXml(numberingXml)
      })
    },
    styles: collectStyles(styleXml),
    numbering: collectNumbering(numberingXml),
    parts: profileParts
  };

  profile.summary = {
    parts: profileParts.map(summarizeProfilePart),
    counts: {
      styles: profile.styles.length,
      abstractNums: profile.numbering.abstractNums.length,
      nums: profile.numbering.nums.length,
      paragraphs: profileParts.reduce((sum, part) => sum + part.paragraphs.length, 0),
      runs: profileParts.reduce((sum, part) => (
        sum + part.paragraphs.reduce((partSum, paragraph) => partSum + paragraph.runCount, 0)
      ), 0),
      tables: profileParts.reduce((sum, part) => sum + part.tables.length, 0),
      sections: profileParts.reduce((sum, part) => sum + part.sections.length, 0),
      textNodes: profileParts.reduce((sum, part) => sum + part.textNodeCount, 0)
    }
  };
  profile.signals = detectProfileSignals(profile);

  return profile;
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

export async function profileDocx(inputDocx, options = {}) {
  const zip = await loadDocx(inputDocx);
  const profile = await profileZip(zip);
  const result = {
    ok: true,
    command: "template.profile",
    input: path.resolve(inputDocx),
    profile
  };

  const outputPath = options.outputPath || options.out;
  if (outputPath) {
    await fs.mkdir(path.dirname(path.resolve(outputPath)), { recursive: true });
    await fs.writeFile(outputPath, `${JSON.stringify(profile, null, 2)}\n`, "utf8");
    result.output = path.resolve(outputPath);
  }

  if (options.summary) {
    result.profile = {
      version: profile.version,
      kind: profile.kind,
      generatedAt: profile.generatedAt,
      package: profile.package,
      summary: profile.summary,
      signals: profile.signals
    };
  }

  return result;
}

export async function analyzeTemplates(inputDocxPaths, options = {}) {
  const inputs = asArray(inputDocxPaths).filter(Boolean);
  if (!inputs.length) {
    throw new Error("template analyze requires at least one input DOCX");
  }

  const ir = await analyzeDocxInputs(inputs);
  const result = {
    ok: true,
    command: "template.analyze",
    inputs: inputs.map((input) => path.resolve(input)),
    ir
  };

  const outputPath = options.outputPath || options.out;
  if (outputPath) {
    await fs.mkdir(path.dirname(path.resolve(outputPath)), { recursive: true });
    await fs.writeFile(outputPath, `${JSON.stringify(ir, null, 2)}\n`, "utf8");
    result.output = path.resolve(outputPath);
  }

  if (options.summary) {
    result.ir = {
      version: ir.version,
      mode: ir.mode,
      generatedAt: ir.generatedAt,
      sources: ir.sources,
      summary: ir.summary,
      alignment: ir.alignment,
      conflicts: ir.conflicts.slice(0, 20),
      profileSignals: ir.profileSignals.slice(0, 20)
    };
  }

  return result;
}

export async function planOfficeTemplate(irPath, options = {}) {
  if (!irPath) {
    throw new Error("template plan-office requires <template-ir.json>");
  }
  const ir = JSON.parse(await fs.readFile(irPath, "utf8"));
  const plan = buildOfficeTemplatePlan(ir);
  const result = {
    ok: true,
    command: "template.planOffice",
    input: path.resolve(irPath),
    plan
  };

  const outputPath = options.outputPath || options.out;
  if (outputPath) {
    await fs.mkdir(path.dirname(path.resolve(outputPath)), { recursive: true });
    await fs.writeFile(outputPath, `${JSON.stringify(plan, null, 2)}\n`, "utf8");
    result.output = path.resolve(outputPath);
  }

  if (options.summary) {
    result.plan = {
      version: plan.version,
      kind: plan.kind,
      generatedAt: plan.generatedAt,
      sourceIr: plan.sourceIr,
      nativeTargets: plan.nativeTargets,
      backendDecision: plan.backendDecision,
      summary: plan.summary,
      gaps: plan.gaps.slice(0, 20)
    };
  }

  return result;
}

export async function compileOfficeTemplate(inputDocx, options = {}) {
  const planPath = options.planPath || options.plan;
  const outputPath = options.outputPath || options.out;
  if (!inputDocx || !planPath || !outputPath) {
    throw new Error("template compile-office requires <input.docx> --plan <office-plan.json> --out <template.docx>");
  }

  const plan = JSON.parse(await fs.readFile(planPath, "utf8"));
  const zip = await loadDocx(inputDocx);
  const parts = await readWordParts(zip);
  const byPart = new Map(parts.map((part) => [part.name, part.xml]));
  const candidateOverrides = await loadCandidateOverrides(options);
  const applied = [];
  const skipped = [];

  for (const patch of plan.patches || []) {
    if (!["slot", "loop"].includes(patch.ruleKind)) {
      skipped.push({ patchId: patch.id, ruleId: patch.ruleId, reason: "only safe scalar slots and table loops are implemented in this compiler pass" });
      continue;
    }

    const evidence = Array.isArray(patch.evidence) ? patch.evidence.find((item) => item.part) : null;
    const partName = evidence?.part;
    if (!partName || !byPart.has(partName)) {
      skipped.push({ patchId: patch.id, ruleId: patch.ruleId, reason: "target part was not found" });
      continue;
    }

    const currentXml = byPart.get(partName);
    const override = candidateOverrides.get(patch.id) || null;
    const result = patch.ruleKind === "loop"
      ? compileLoopPatchIntoPart(currentXml, patch, override)
      : compileSlotPatchIntoPart(currentXml, patch, override);
    if (result.applied) {
      byPart.set(partName, result.xml);
      applied.push({
        patchId: patch.id,
        ruleId: patch.ruleId,
        part: partName,
        confidence: result.confidence,
        overrideApplied: result.overrideApplied || false,
        reason: result.reason
      });
    } else {
      skipped.push({
        patchId: patch.id,
        ruleId: patch.ruleId,
        part: partName,
        confidence: result.confidence || 0,
        candidates: result.candidates || [],
        reason: result.reason
      });
    }
  }

  for (const [partName, xml] of byPart.entries()) {
    zip.file(partName, xml);
  }
  const customXml = await addCustomXmlPart(zip, (plan.patches || []).filter((patch) => applied.some((item) => item.patchId === patch.id)));
  await writeDocx(zip, outputPath);
  const profile = await inspectZip(zip);

  return {
    ok: true,
    command: "template.compileOffice",
    input: path.resolve(inputDocx),
    plan: path.resolve(planPath),
    output: path.resolve(outputPath),
    appliedPatches: applied.length,
    skippedPatches: skipped.length,
    customXml,
    applied,
    skipped,
    profileFingerprint: profile.packageFingerprint
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
