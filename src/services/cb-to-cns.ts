/**
 * CB Report → CNS Report conversion service
 *
 * New simplified flow:
 * 1. PDF → Adobe → DOCX (preserves all formatting: tables, styles, etc.)
 * 2. Parse DOCX to get text segments with indices
 * 3. LLM identifies which segments belong to 7 categories (returns segment indices)
 * 4. Keep only the relevant segments in the DOCX
 * 5. Continue with existing translation pipeline
 */

import * as fs from "fs";
import * as path from "path";
import JSZip from "jszip";
import { XMLParser, XMLBuilder } from "fast-xml-parser";
import { AzureOpenAI } from "openai";
import { JobState, updateJob } from "../jobs";
import { convertPdfToDocx } from "./adobe";
import { parseDocx, DocxSegment } from "./docx";

// Initialize Azure OpenAI client
let client: AzureOpenAI | null = null;

function getClient(): AzureOpenAI {
  if (!client) {
    const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
    const apiKey = process.env.AZURE_OPENAI_API_KEY;
    const apiVersion = process.env.AZURE_OPENAI_API_VERSION || "2024-10-01-preview";

    if (!endpoint || !apiKey) {
      throw new Error("AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_API_KEY must be set");
    }

    client = new AzureOpenAI({
      endpoint,
      apiKey,
      apiVersion,
      timeout: 300000, // 5 minutes for large PDF analysis
    });
  }
  return client;
}

/**
 * Page range info for each category
 */
export interface PageRanges {
  testItemParticulars: string;      // 1. 試驗樣品特性
  factoryAddresses: string;         // 2. 工廠地址
  energySafeguards: string;         // 3. 安全防護總攬表
  clauses4to10: string;             // 4. 主體條文 4-10 章
  annexBtoG: string;                // 5. 附錄條文 B-G
  criticalComponents: string;       // 6. 4.1.2 重要零件列表
  detailedTestData: string;         // 7. 詳細試驗數據表
}

/**
 * Segment range for each category (which paragraph indices to keep)
 */
interface SegmentRanges {
  testItemParticulars: number[];
  factoryAddresses: number[];
  energySafeguards: number[];
  clauses4to10: number[];
  annexBtoG: number[];
  criticalComponents: number[];
  detailedTestData: number[];
}

/**
 * The CB→CNS analysis prompt - now returns segment indices
 */
const CB_TO_CNS_SYSTEM_PROMPT = `你是一個「CB 報告 → CNS 報告」分析助手。

任務目標：
給你一份 CB 測試報告的段落列表（每段有編號），請你找出屬於以下七大類的段落編號。

========================
【七大類內容】

1. 試驗樣品特性 / Test item particulars
   線索：標題含 "Test item particulars"，表格欄位包含 Product group, Classification of use, Supply connection 等

2. 工廠地址 / Name and address of factory(ies)
   線索：標題含 "Name and address of factory(ies)"，下面是工廠名稱與地址

3. 安全防護總攬表 / OVERVIEW OF ENERGY SOURCES AND SAFEGUARDS
   線索：標題含 "OVERVIEW OF ENERGY SOURCES AND SAFEGUARDS"，有 Clause 5–10、ES1/ES3、PS3 等

4. 主體條文 4～10 章（Clause/Requirement/Result/Verdict 表）
   線索：表格形式，條文編號如 4.x, 5.x, 6.x, 7.x, 8.x, 9.x, 10.x，Verdict 欄位有 P、N/A

5. 附錄條文 Annex B～G
   線索：分段標題包含 B.1, B.2, C.1, D.1, E.1, F.1, G.8 等

6. 4.1.2 重要零件列表 & 零件認證表
   線索：表格標題含 "4.1.2"，記錄 MOV、保險絲、電感等零件

7. 詳細試驗數據表（接觸電流、PS 分級、溫升等）
   線索：表格標題如 "5.7.4 TABLE", "5.7.5 TABLE", "6.2.2 TABLE", "Temperature measurements" 等

========================
【輸出格式】

請只輸出 JSON 格式，列出每個類別包含的段落編號（可以是範圍或單獨編號）：

{
  "testItemParticulars": [10, 11, 12],
  "factoryAddresses": [5, 6, 7, 8],
  "energySafeguards": [20, 21, 22, 23, 24],
  "clauses4to10": [30, 31, 32, "...到", 100],
  "annexBtoG": [110, 111, "...到", 150],
  "criticalComponents": [160, 161, 162],
  "detailedTestData": [200, 201, 202, 203]
}

注意：
- 如果某類別找不到，設為空陣列 []
- 段落編號從 0 開始
- 可以用連續數字表示範圍
- 只輸出 JSON，不要其他文字`;

/**
 * Parse LLM response to get segment indices
 */
function parseSegmentRanges(response: string): SegmentRanges {
  const defaultRanges: SegmentRanges = {
    testItemParticulars: [],
    factoryAddresses: [],
    energySafeguards: [],
    clauses4to10: [],
    annexBtoG: [],
    criticalComponents: [],
    detailedTestData: [],
  };

  try {
    // Extract JSON from response
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      console.warn("No JSON found in LLM response");
      return defaultRanges;
    }

    const parsed = JSON.parse(jsonMatch[0]);

    // Process each category
    const processArray = (arr: any[]): number[] => {
      if (!Array.isArray(arr)) return [];
      const result: number[] = [];
      for (const item of arr) {
        if (typeof item === "number") {
          result.push(item);
        }
      }
      return result;
    };

    return {
      testItemParticulars: processArray(parsed.testItemParticulars),
      factoryAddresses: processArray(parsed.factoryAddresses),
      energySafeguards: processArray(parsed.energySafeguards),
      clauses4to10: processArray(parsed.clauses4to10),
      annexBtoG: processArray(parsed.annexBtoG),
      criticalComponents: processArray(parsed.criticalComponents),
      detailedTestData: processArray(parsed.detailedTestData),
    };
  } catch (e) {
    console.error("Failed to parse segment ranges:", e);
    return defaultRanges;
  }
}

/**
 * Convert segment ranges to page ranges string (for compatibility)
 */
function segmentRangesToPageRanges(ranges: SegmentRanges): PageRanges {
  const formatRange = (arr: number[]): string => {
    if (arr.length === 0) return "not_found";
    const min = Math.min(...arr);
    const max = Math.max(...arr);
    return `segments ${min}-${max}`;
  };

  return {
    testItemParticulars: formatRange(ranges.testItemParticulars),
    factoryAddresses: formatRange(ranges.factoryAddresses),
    energySafeguards: formatRange(ranges.energySafeguards),
    clauses4to10: formatRange(ranges.clauses4to10),
    annexBtoG: formatRange(ranges.annexBtoG),
    criticalComponents: formatRange(ranges.criticalComponents),
    detailedTestData: formatRange(ranges.detailedTestData),
  };
}

/**
 * XML parser/builder options
 */
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
 * Filter DOCX to keep only paragraphs in the specified indices
 * This modifies the DOCX in place to remove unwanted paragraphs
 */
async function filterDocxBySegments(
  docxPath: string,
  outputPath: string,
  keepIndices: Set<number>
): Promise<void> {
  // Read the DOCX file
  const buffer = fs.readFileSync(docxPath);
  const zip = await JSZip.loadAsync(buffer);

  // Read document.xml
  const documentXmlFile = zip.file("word/document.xml");
  if (!documentXmlFile) {
    throw new Error("Invalid DOCX: word/document.xml not found");
  }

  const documentXmlString = await documentXmlFile.async("string");
  const parser = new XMLParser(parserOptions);
  const documentXml = parser.parse(documentXmlString);

  // Find and filter paragraphs
  let paragraphIndex = 0;

  function filterParagraphs(node: any): void {
    if (Array.isArray(node)) {
      // Filter out paragraphs not in keepIndices
      for (let i = node.length - 1; i >= 0; i--) {
        const item = node[i];
        if (typeof item === "object" && item !== null) {
          if ("w:p" in item) {
            // This is a paragraph
            if (!keepIndices.has(paragraphIndex)) {
              node.splice(i, 1); // Remove this paragraph
            }
            paragraphIndex++;
          } else {
            filterParagraphs(item);
          }
        }
      }
      return;
    }

    if (typeof node !== "object" || node === null) {
      return;
    }

    for (const key of Object.keys(node)) {
      if (key !== ":@" && key !== "#text") {
        filterParagraphs(node[key]);
      }
    }
  }

  // Note: For now, we'll keep all paragraphs since filtering might break table structure
  // The LLM will help with translation focus, but we preserve the document structure

  // Build new XML
  const builder = new XMLBuilder(builderOptions);
  const newXmlString = builder.build(documentXml);

  // Update zip and save
  zip.file("word/document.xml", newXmlString);

  const outputBuffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });

  const outputDir = path.dirname(outputPath);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  fs.writeFileSync(outputPath, outputBuffer);
}

/**
 * Main function: Analyze CB PDF and prepare DOCX for translation
 *
 * New flow:
 * 1. PDF → Adobe → DOCX (preserves formatting)
 * 2. Copy DOCX to output path (keep all formatting)
 * 3. Optionally analyze with LLM for logging
 */
export async function analyzeCbAndGenerateCnsDocx(
  job: JobState,
  pdfPath: string,
  outputDocxPath: string
): Promise<{ pageRanges: PageRanges }> {
  const openai = getClient();
  const deployment = process.env.AZURE_OPENAI_DEPLOYMENT_NAME;

  if (!deployment) {
    throw new Error("AZURE_OPENAI_DEPLOYMENT_NAME must be set");
  }

  // Step 1: Convert PDF to DOCX using Adobe (preserves all formatting!)
  updateJob(job, {
    stepMessage: "正在用 Adobe API 轉換 PDF 為 DOCX（保留格式）...",
    progress: 3,
  });

  const tempDocxPath = path.join(path.dirname(outputDocxPath), `temp-adobe-${Date.now()}.docx`);
  await convertPdfToDocx(pdfPath, tempDocxPath);

  if (job.cancelled) {
    throw new Error("Job cancelled");
  }

  // Step 2: Parse DOCX to get segments (for LLM analysis)
  updateJob(job, {
    stepMessage: "正在解析文件結構...",
    progress: 8,
  });

  const parsed = await parseDocx(tempDocxPath);
  const segments = parsed.segments;
  console.log(`Parsed ${segments.length} segments from Adobe DOCX`);

  if (job.cancelled) {
    throw new Error("Job cancelled");
  }

  // Step 3: Send to LLM for analysis (identify which segments are important)
  updateJob(job, {
    status: "converting",
    stepMessage: "正在分析 CB 報告結構...",
    progress: 10,
  });

  // Build segment list for LLM
  const segmentList = segments.map((s, i) => `[${i}] ${s.text.substring(0, 200)}${s.text.length > 200 ? "..." : ""}`).join("\n");

  // Truncate if too long
  const maxChars = 80000;
  const truncatedList = segmentList.length > maxChars
    ? segmentList.substring(0, maxChars) + "\n\n[... 段落列表已截斷 ...]"
    : segmentList;

  const userContent = `以下是 CB 測試報告的段落列表（共 ${segments.length} 段）：\n\n${truncatedList}`;

  let segmentRanges: SegmentRanges;
  let pageRanges: PageRanges;

  try {
    const response = await openai.chat.completions.create(
      {
        model: deployment,
        messages: [
          { role: "system", content: CB_TO_CNS_SYSTEM_PROMPT },
          { role: "user", content: userContent },
        ],
        max_completion_tokens: 4000,
      },
      { signal: job.abortController.signal }
    );

    // Update token usage
    const usage = response.usage;
    if (usage) {
      job.usage.prompt += usage.prompt_tokens || 0;
      job.usage.completion += usage.completion_tokens || 0;
    }

    const content = response.choices[0]?.message?.content;
    if (content) {
      segmentRanges = parseSegmentRanges(content);
      pageRanges = segmentRangesToPageRanges(segmentRanges);
      console.log("Segment ranges found:", segmentRanges);
    } else {
      // Default: keep all segments
      segmentRanges = {
        testItemParticulars: [],
        factoryAddresses: [],
        energySafeguards: [],
        clauses4to10: [],
        annexBtoG: [],
        criticalComponents: [],
        detailedTestData: [],
      };
      pageRanges = segmentRangesToPageRanges(segmentRanges);
    }
  } catch (error: any) {
    if (error.name === "AbortError" || job.cancelled) {
      // Clean up temp file
      try { fs.unlinkSync(tempDocxPath); } catch {}
      throw new Error("Job cancelled");
    }
    console.error("LLM analysis failed (continuing anyway):", error.message);
    // Continue without filtering - just use the Adobe DOCX as-is
    pageRanges = {
      testItemParticulars: "analysis_failed",
      factoryAddresses: "analysis_failed",
      energySafeguards: "analysis_failed",
      clauses4to10: "analysis_failed",
      annexBtoG: "analysis_failed",
      criticalComponents: "analysis_failed",
      detailedTestData: "analysis_failed",
    };
  }

  if (job.cancelled) {
    try { fs.unlinkSync(tempDocxPath); } catch {}
    throw new Error("Job cancelled");
  }

  // Step 4: Copy Adobe DOCX to output path (preserving all formatting!)
  updateJob(job, {
    stepMessage: "正在準備 CNS 報告文件...",
    progress: 14,
  });

  // Simply copy the Adobe-converted DOCX - this preserves all tables, styles, formatting
  fs.copyFileSync(tempDocxPath, outputDocxPath);
  console.log(`Copied Adobe DOCX to: ${outputDocxPath}`);

  // Clean up temp file
  try {
    fs.unlinkSync(tempDocxPath);
  } catch {
    // Ignore cleanup errors
  }

  return {
    pageRanges,
  };
}
