/**
 * DOCX parsing and writing using JSZip and fast-xml-parser
 */

import * as fs from "fs";
import * as path from "path";
import JSZip from "jszip";
import { XMLParser, XMLBuilder } from "fast-xml-parser";

// Track text nodes and their parent run elements
interface WtNodeInfo {
  textNode: any;
  runNode: any | null;
}

export interface DocxSegment {
  id: number;
  text: string;
  translated?: string;
  // Track the w:t nodes and their parent runs
  wtNodeInfos: WtNodeInfo[];
}

export interface ParsedDocx {
  zip: JSZip;
  documentXml: any;
  segments: DocxSegment[];
}

// XML parser/builder options for preserving structure
const parserOptions = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  textNodeName: "#text",
  parseTagValue: false,
  trimValues: false,
};

const builderOptions = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  textNodeName: "#text",
  format: false,
  suppressEmptyNode: false,
  suppressBooleanAttributes: false,
};

/**
 * Collect all w:t text nodes and their parent w:r (run) nodes
 */
function collectWtNodes(node: any, wtNodes: WtNodeInfo[], parentRun: any | null = null): void {
  if (Array.isArray(node)) {
    for (const item of node) {
      collectWtNodes(item, wtNodes, parentRun);
    }
    return;
  }

  if (typeof node !== "object" || node === null) {
    return;
  }

  // Check if this is a w:r (run) node
  if ("w:r" in node) {
    // Recurse into run with this as the parent
    collectWtNodes(node["w:r"], wtNodes, node);
    return;
  }

  // Check if this is a w:t node
  if ("w:t" in node) {
    const wtContent = node["w:t"];
    if (Array.isArray(wtContent)) {
      for (const item of wtContent) {
        if (typeof item === "object" && "#text" in item) {
          wtNodes.push({ textNode: item, runNode: parentRun });
        }
      }
    }
    return;
  }

  // Recurse into child nodes
  for (const key of Object.keys(node)) {
    if (key !== ":@" && key !== "#text") {
      collectWtNodes(node[key], wtNodes, parentRun);
    }
  }
}

/**
 * Extract segments by paragraph (w:p) for better context
 */
function extractSegmentsByParagraph(
  node: any,
  segments: DocxSegment[],
  counter: { id: number }
): void {
  if (Array.isArray(node)) {
    for (const item of node) {
      extractSegmentsByParagraph(item, segments, counter);
    }
    return;
  }

  if (typeof node !== "object" || node === null) {
    return;
  }

  // Check if this is a w:p (paragraph) node
  if ("w:p" in node) {
    const wtNodeInfos: WtNodeInfo[] = [];
    collectWtNodes(node["w:p"], wtNodeInfos);

    if (wtNodeInfos.length > 0) {
      // Merge all text in this paragraph
      const mergedText = wtNodeInfos.map((info) => String(info.textNode["#text"])).join("");

      if (mergedText.trim()) {
        segments.push({
          id: counter.id++,
          text: mergedText,
          wtNodeInfos: wtNodeInfos,
        });
      }
    }
    return; // Don't recurse into paragraph children (already processed)
  }

  // Recurse into other nodes
  for (const key of Object.keys(node)) {
    if (key !== ":@" && key !== "#text") {
      extractSegmentsByParagraph(node[key], segments, counter);
    }
  }
}

/**
 * Remove w:spacing from run properties to allow text to reflow naturally
 */
function clearRunSpacing(runNode: any): void {
  if (!runNode || !runNode["w:r"]) return;

  const runContent = runNode["w:r"];
  if (!Array.isArray(runContent)) return;

  for (const item of runContent) {
    if (item && "w:rPr" in item) {
      const rPr = item["w:rPr"];
      if (Array.isArray(rPr)) {
        // Remove w:spacing elements from run properties
        for (let i = rPr.length - 1; i >= 0; i--) {
          if (rPr[i] && "w:spacing" in rPr[i]) {
            rPr.splice(i, 1);
          }
        }
      }
    }
  }
}

/**
 * Distribute translated text back to w:t nodes
 * Strategy: Put all text in first non-empty node, clear others,
 * and remove fixed spacing from runs to allow natural text flow
 */
function distributeTranslation(segment: DocxSegment): void {
  if (!segment.translated || segment.wtNodeInfos.length === 0) {
    return;
  }

  const translated = segment.translated;
  const infos = segment.wtNodeInfos;

  // Clear spacing from all runs that will be modified
  for (const info of infos) {
    if (info.runNode) {
      clearRunSpacing(info.runNode);
    }
  }

  if (infos.length === 1) {
    // Simple case: single w:t node
    infos[0].textNode["#text"] = translated;
    return;
  }

  // Check if any nodes are whitespace-only (important for table spacing)
  const nodeAnalysis = infos.map((info) => {
    const text = String(info.textNode["#text"]);
    return {
      info,
      text,
      isWhitespaceOnly: /^\s*$/.test(text),
    };
  });

  // Find the first non-whitespace node to put the translation
  const firstContentIndex = nodeAnalysis.findIndex((n) => !n.isWhitespaceOnly);

  if (firstContentIndex === -1) {
    // All nodes are whitespace - put translation in first node
    infos[0].textNode["#text"] = translated;
    return;
  }

  // Put translation in the first content node
  infos[firstContentIndex].textNode["#text"] = translated;

  // Clear other content nodes, but preserve whitespace nodes
  for (let i = 0; i < infos.length; i++) {
    if (i !== firstContentIndex && !nodeAnalysis[i].isWhitespaceOnly) {
      infos[i].textNode["#text"] = "";
    }
  }
}

/**
 * Parse a DOCX file and extract text segments
 */
export async function parseDocx(filePath: string): Promise<ParsedDocx> {
  // Read the DOCX file
  const buffer = fs.readFileSync(filePath);

  // Load as ZIP
  const zip = await JSZip.loadAsync(buffer);

  // Read word/document.xml
  const documentXmlFile = zip.file("word/document.xml");
  if (!documentXmlFile) {
    throw new Error("Invalid DOCX: word/document.xml not found");
  }

  const documentXmlString = await documentXmlFile.async("string");

  // Parse XML
  const parser = new XMLParser(parserOptions);
  const documentXml = parser.parse(documentXmlString);

  // Extract text segments by paragraph
  const segments: DocxSegment[] = [];
  const counter = { id: 0 };
  extractSegmentsByParagraph(documentXml, segments, counter);

  console.log(`Parsed DOCX: found ${segments.length} paragraph segments`);

  return {
    zip,
    documentXml,
    segments,
  };
}

/**
 * Write translated segments back to DOCX
 */
export async function writeDocx(
  parsed: ParsedDocx,
  outputPath: string
): Promise<void> {
  // Apply translations to the original XML structure
  for (const segment of parsed.segments) {
    distributeTranslation(segment);
  }

  // Build XML string
  const builder = new XMLBuilder(builderOptions);
  const newXmlString = builder.build(parsed.documentXml);

  // Update the zip with new document.xml
  parsed.zip.file("word/document.xml", newXmlString);

  // Generate new DOCX buffer
  const outputBuffer = await parsed.zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });

  // Ensure output directory exists
  const outputDir = path.dirname(outputPath);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // Write to file
  fs.writeFileSync(outputPath, outputBuffer);
  console.log(`Wrote translated DOCX: ${outputPath}`);
}
