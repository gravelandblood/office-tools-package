import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

const WORD_TEXT_PART = /^word\/(document|footnotes|endnotes|comments)\.xml$|^word\/(header|footer)\d+\.xml$/;
const XML_DECLARATION = /^<\?xml[^>]*>\s*/u;
const TEXT_NODE_RE = /<w:t\b([^>]*)>([\s\S]*?)<\/w:t>/gu;
const TAG_RE = /<[^>]+>/gu;
const PARSED_TEXT_KEYS = new Set(["w:t", "w:delText", "w:instrText", "w:delInstrText"]);

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
  if (node.kind === "paragraph") {
    return `p:${node.part}:${node.index}`;
  }
  if (node.kind === "table") {
    return `tbl:${node.part}:${node.index}`;
  }
  return `${node.kind}:${node.part || ""}:${node.index ?? ""}`;
}

function textClass(text) {
  const value = String(text || "").trim();
  if (!value) return "empty";
  if (/^\d{4}[-/.年]\d{1,2}([-/.月]\d{1,2}日?)?$/u.test(value)) return "date";
  if (/^[\d,]+(\.\d+)?%?$/u.test(value)) return "number";
  if (valueHasLabel(value)) return "labelValue";
  if (value.length <= 12) return "shortText";
  return "longText";
}

function valueHasLabel(value) {
  return /[:：]/u.test(value) && value.replace(/[:：].*$/u, "").trim().length <= 20;
}

function splitLabelValue(value) {
  const match = String(value || "").match(/^(.{1,30}?)[：:]\s*(.+)$/u);
  if (!match) return null;
  return {
    label: match[1].trim(),
    value: match[2].trim()
  };
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
        structureFingerprint: table.formatFingerprint,
        formatFingerprint: table.tablePropertiesFingerprint,
        rowCount: table.rowCount,
        columnCount: table.columnCount,
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
  });
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
  if (/^[一二三四五六七八九十]+[、.．]/u.test(text) || /^\d+[.．、]/u.test(text)) candidates.push("numberedHeading");
  if (text.length <= 30 && block.original.runs.some((run) => run.rPr?.["w:b"] !== undefined)) candidates.push("heading");
  if (text.length > 80) candidates.push("body");
  if (!candidates.length) candidates.push(textClass(text));
  return candidates;
}

function groupBlocksForRules(sources) {
  const byPosition = new Map();
  for (const source of sources) {
    for (const block of collectDocumentBlocks(source.profile, source.id)) {
      const key = blockSignature(block);
      if (!byPosition.has(key)) byPosition.set(key, []);
      byPosition.get(key).push(block);
    }
  }
  return byPosition;
}

function inferRulesAndConflicts(sources) {
  const rules = [];
  const conflicts = [];
  const dataFields = {};
  const arrays = {};
  const groups = groupBlocksForRules(sources);
  let fieldIndex = 1;
  let arrayIndex = 1;

  for (const [signature, blocks] of groups) {
    const sourceCount = new Set(blocks.map((block) => block.sourceId)).size;
    const texts = new Set(blocks.map((block) => block.text));
    const formats = new Set(blocks.map((block) => block.formatFingerprint));
    const first = blocks[0];
    const evidences = blocks.map((block) => block.evidence);

    if (sourceCount < sources.length) {
      conflicts.push({
        id: `conflict.${String(conflicts.length + 1).padStart(4, "0")}`,
        type: "optionalBlockAmbiguity",
        severity: "medium",
        classification: "unresolvedConflict",
        evidence: evidences,
        recommendedAction: "Confirm whether this block is optional or missing because samples are not structurally aligned.",
        signature,
        presentIn: sourceCount,
        expectedSources: sources.length
      });
    }

    if (first.kind === "paragraph") {
      const labelValue = blocks.map((block) => splitLabelValue(block.text));
      const hasConsistentLabel = labelValue.every(Boolean)
        && new Set(labelValue.map((item) => item.label)).size === 1;
      const valueSet = new Set(labelValue.filter(Boolean).map((item) => item.value));

      if (hasConsistentLabel && valueSet.size > 1) {
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
          label: labelValue[0].label
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
          validationStatus: "unverified"
        });
      } else if (texts.size > 1 && sourceCount > 1) {
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
          validationStatus: "unverified"
        });
      }
    }

    if (first.kind === "table") {
      const rowCounts = new Set(blocks.map((block) => block.rowCount));
      if (rowCounts.size > 1 || first.rowCount > 2) {
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
          validationStatus: "unverified"
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
    [/公司|企业|单位/u, "company.name"],
    [/日期|时间/u, "report.date"],
    [/姓名|负责人|联系人/u, "person.name"],
    [/金额|价款|总额/u, "amount"],
    [/地址|住所/u, "address"]
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

async function analyzeDocxInputs(inputDocxPaths) {
  const sources = [];
  for (let index = 0; index < inputDocxPaths.length; index += 1) {
    const zip = await loadDocx(inputDocxPaths[index]);
    const profile = await profileZip(zip);
    sources.push(buildProfileSource(inputDocxPaths[index], profile, index));
  }

  const atomMap = new Map();
  const structure = inferStructureAndAtoms(sources, atomMap);
  const { rules, conflicts, dataFields, arrays } = inferRulesAndConflicts(sources);
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
    }
  };

  return ir;
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
      conflicts: ir.conflicts.slice(0, 20),
      profileSignals: ir.profileSignals.slice(0, 20)
    };
  }

  return result;
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
