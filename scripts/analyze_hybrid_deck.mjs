#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";

const SCRIPT_VERSION = "1.2.0";
const EMU_PER_INCH = 914400;
const PX_PER_INCH = 96;
const LARGE_RASTER_AREA_RATIO = 0.2;
const MIN_VISIBLE_STROKE_PX = 1;

const HELP = `Usage:
  analyze_hybrid_deck.mjs --pptx <source.pptx> --workspace <directory>

Read-only source analysis for picture-heavy and hybrid PowerPoint decks.

Required options:
  --pptx <path>       Source .pptx. The source is never edited.
  --workspace <path>  Analysis workspace for extracted evidence.

Outputs:
  source-manifest.json      Normalized slide/object/media inventory
  theme-profile.json        Theme, color, typeface, master/layout evidence
  source-inspect.ndjson     Artifact Tool inspection records
  source-renders/           One PNG per slide
  source-layout/            One layout JSON per slide
  source-media/             Extracted ppt/media assets
  source-montage.webp       Verified all-slide contact sheet when available
  source-montage.json       Per-slide hashes and lossless tile verification
  source-original.pptx      Verified byte-identical working copy

Environment:
  RUNTIME_NODE_MODULES must point at the bundled workspace Node.js packages.
  The directory must contain @oai/artifact-tool and jszip; sharp is used when
  available to record raster dimensions and resolution evidence.

Options:
  -h, --help          Show this help.
`;

class CliError extends Error {
  constructor(message, exitCode = 2) {
    super(message);
    this.name = "CliError";
    this.exitCode = exitCode;
  }
}

function parseArgs(argv) {
  if (argv.includes("--help") || argv.includes("-h")) {
    return { help: true };
  }
  if (argv.length === 0) {
    throw new CliError("Both --pptx and --workspace are required.");
  }
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (token === "--pptx" || token === "--workspace") {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) {
        throw new CliError(`${token} requires a value.`);
      }
      result[token.slice(2)] = value;
      index += 1;
      continue;
    }
    if (token.startsWith("--pptx=")) {
      result.pptx = token.slice("--pptx=".length);
      continue;
    }
    if (token.startsWith("--workspace=")) {
      result.workspace = token.slice("--workspace=".length);
      continue;
    }
    throw new CliError(`Unknown argument: ${token}`);
  }
  if (!result.pptx || !result.workspace) {
    throw new CliError("Both --pptx and --workspace are required.");
  }
  return result;
}

async function loadRuntimeModule(packageName, { optional = false } = {}) {
  const modulesDir = process.env.RUNTIME_NODE_MODULES;
  if (!modulesDir) {
    if (optional) return null;
    throw new CliError(
      "RUNTIME_NODE_MODULES is not set. Load the bundled workspace dependencies before running the analyzer.",
    );
  }
  const requireFromRuntime = createRequire(
    path.join(path.resolve(modulesDir), "pptx-refactor-loader.cjs"),
  );
  let resolved;
  try {
    resolved = requireFromRuntime.resolve(packageName);
  } catch (error) {
    if (optional) return null;
    throw new CliError(
      `Cannot resolve ${packageName} from RUNTIME_NODE_MODULES (${modulesDir}): ${error.message}`,
    );
  }
  return import(pathToFileURL(resolved).href);
}

function normalizeZipPath(value) {
  return path.posix.normalize(String(value ?? "").replaceAll("\\", "/").replace(/^\/+/, ""));
}

function normalizePartTarget(basePart, target) {
  const raw = String(target ?? "").replaceAll("\\", "/");
  if (!raw) return null;
  if (raw.startsWith("/")) return normalizeZipPath(raw);
  return normalizeZipPath(path.posix.join(path.posix.dirname(basePart), raw));
}

function xmlDecode(value) {
  return String(value ?? "")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&");
}

function parseAttributes(fragment) {
  const attributes = {};
  const pattern = /([A-Za-z_][\w:.-]*)\s*=\s*(["'])(.*?)\2/g;
  for (const match of fragment.matchAll(pattern)) {
    attributes[match[1]] = xmlDecode(match[3]);
  }
  return attributes;
}

function parseRelationships(xml, sourcePart) {
  const relationships = [];
  const pattern = /<(?:\w+:)?Relationship\b([^>]*?)(?:\/>|>\s*<\/(?:\w+:)?Relationship>)/gi;
  for (const match of String(xml ?? "").matchAll(pattern)) {
    const attrs = parseAttributes(match[1]);
    const external = String(attrs.TargetMode ?? "").toLowerCase() === "external";
    relationships.push({
      id: attrs.Id ?? null,
      type: attrs.Type ?? null,
      target: attrs.Target ?? null,
      targetMode: attrs.TargetMode ?? null,
      external,
      resolvedTarget: external ? null : normalizePartTarget(sourcePart, attrs.Target),
    });
  }
  return relationships;
}

function naturalPartSort(left, right) {
  return left.localeCompare(right, undefined, { numeric: true, sensitivity: "base" });
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function hashFile(filePath) {
  const handle = await fs.open(filePath, "r");
  const digest = crypto.createHash("sha256");
  try {
    for await (const chunk of handle.readableWebStream()) {
      digest.update(Buffer.from(chunk));
    }
  } finally {
    await handle.close().catch(() => {});
  }
  return digest.digest("hex");
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, Buffer.from(await blob.arrayBuffer()));
}

async function writeJson(filePath, value) {
  await fs.writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

function relativeArtifact(workspace, target) {
  if (!target) return null;
  return path.relative(workspace, target).replaceAll("\\", "/");
}

function ensureInsideWorkspace(workspace, candidate) {
  const relative = path.relative(workspace, candidate);
  if (!relative || relative === ".") return;
  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new CliError(`Refusing to write outside workspace: ${candidate}`);
  }
}

async function resetDirectory(workspace, directory) {
  ensureInsideWorkspace(workspace, directory);
  await fs.rm(directory, { recursive: true, force: true });
  await fs.mkdir(directory, { recursive: true });
}

function contentTypeForExtension(extension) {
  const types = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".tif": "image/tiff",
    ".tiff": "image/tiff",
    ".webp": "image/webp",
    ".svg": "image/svg+xml",
    ".emf": "image/x-emf",
    ".wmf": "image/x-wmf",
    ".eps": "application/postscript",
    ".mp3": "audio/mpeg",
    ".wav": "audio/wav",
    ".m4a": "audio/mp4",
    ".mp4": "video/mp4",
    ".mov": "video/quicktime",
    ".avi": "video/x-msvideo",
    ".wmv": "video/x-ms-wmv",
  };
  return types[String(extension ?? "").toLowerCase()] ?? "application/octet-stream";
}

function isRasterContentType(contentType) {
  return /^image\/(?:png|jpeg|gif|bmp|tiff|webp)$/i.test(String(contentType ?? ""));
}

async function safeZipText(zip, partName) {
  const part = zip.file(partName);
  if (!part) return null;
  try {
    return await part.async("string");
  } catch {
    return null;
  }
}

function parseSlideOrder(presentationXml, presentationRels, zipNames) {
  const relById = new Map(presentationRels.map((relationship) => [relationship.id, relationship]));
  const result = [];
  const pattern = /<(?:p:)?sldId\b([^>]*)\/?\s*>/gi;
  for (const match of String(presentationXml ?? "").matchAll(pattern)) {
    const attrs = parseAttributes(match[1]);
    const relationship = relById.get(attrs["r:id"]);
    if (relationship?.resolvedTarget?.startsWith("ppt/slides/slide")) {
      result.push(relationship.resolvedTarget);
    }
  }
  if (result.length > 0) return result;
  return zipNames
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort(naturalPartSort);
}

function parseSlideSize(presentationXml) {
  const match = String(presentationXml ?? "").match(/<(?:p:)?sldSz\b([^>]*)\/?\s*>/i);
  if (!match) return null;
  const attrs = parseAttributes(match[1]);
  const cx = Number(attrs.cx);
  const cy = Number(attrs.cy);
  if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) return null;
  return {
    width: Number(((cx / EMU_PER_INCH) * PX_PER_INCH).toFixed(2)),
    height: Number(((cy / EMU_PER_INCH) * PX_PER_INCH).toFixed(2)),
    unit: "px",
    sourceEmu: { cx, cy },
  };
}

function extractThemeFromXml(xml, partName) {
  const result = {
    sourcePart: partName,
    colorSchemeName: null,
    colors: {},
    fontSchemeName: null,
    fontScheme: {},
  };
  const colorScheme = String(xml ?? "").match(/<a:clrScheme\b([^>]*)>([\s\S]*?)<\/a:clrScheme>/i);
  if (colorScheme) {
    result.colorSchemeName = parseAttributes(colorScheme[1]).name ?? null;
    for (const key of ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]) {
      const colorMatch = colorScheme[2].match(new RegExp(`<a:${key}\\b[^>]*>([\\s\\S]*?)<\\/a:${key}>`, "i"));
      if (!colorMatch) continue;
      const valueMatch = colorMatch[1].match(/<a:(?:srgbClr|sysClr|scrgbClr|prstClr)\b([^>]*)\/?\s*>/i);
      if (!valueMatch) continue;
      const attrs = parseAttributes(valueMatch[1]);
      const raw = attrs.lastClr ?? attrs.val ?? null;
      result.colors[key] = raw && /^[0-9A-Fa-f]{6}$/.test(raw) ? `#${raw.toUpperCase()}` : raw;
    }
  }
  const fontScheme = String(xml ?? "").match(/<a:fontScheme\b([^>]*)>([\s\S]*?)<\/a:fontScheme>/i);
  if (fontScheme) {
    result.fontSchemeName = parseAttributes(fontScheme[1]).name ?? null;
    for (const family of ["majorFont", "minorFont"]) {
      const familyMatch = fontScheme[2].match(new RegExp(`<a:${family}\\b[^>]*>([\\s\\S]*?)<\\/a:${family}>`, "i"));
      if (!familyMatch) continue;
      const latin = familyMatch[1].match(/<a:latin\b([^>]*)\/?\s*>/i);
      const eastAsian = familyMatch[1].match(/<a:ea\b([^>]*)\/?\s*>/i);
      const complexScript = familyMatch[1].match(/<a:cs\b([^>]*)\/?\s*>/i);
      result.fontScheme[family] = {
        latinTypeface: latin ? parseAttributes(latin[1]).typeface ?? null : null,
        eastAsianTypeface: eastAsian ? parseAttributes(eastAsian[1]).typeface ?? null : null,
        complexScriptTypeface: complexScript ? parseAttributes(complexScript[1]).typeface ?? null : null,
      };
    }
  }
  return result;
}

function collectTypefaceEvidence(xmlParts, layoutJsons, rawThemes) {
  const counts = new Map();
  const add = (typeface, source) => {
    const normalized = String(typeface ?? "").trim();
    if (!normalized) return;
    const current = counts.get(normalized) ?? { typeface: normalized, count: 0, sources: new Set() };
    current.count += 1;
    if (source) current.sources.add(source);
    counts.set(normalized, current);
  };

  for (const [partName, xml] of xmlParts) {
    for (const match of String(xml).matchAll(/\btypeface\s*=\s*(["'])(.*?)\1/gi)) {
      add(xmlDecode(match[2]), partName);
    }
  }
  const visit = (value, source) => {
    if (!value || typeof value !== "object") return;
    if (typeof value.typeface === "string") add(value.typeface, source);
    if (Array.isArray(value.typefaces)) {
      for (const typeface of value.typefaces) add(typeface, source);
    }
    for (const child of Object.values(value)) {
      if (child && typeof child === "object") visit(child, source);
    }
  };
  for (const layout of layoutJsons) visit(layout, `layout:slide-${layout.slide?.slide ?? "unknown"}`);
  for (const theme of rawThemes) {
    for (const family of Object.values(theme.fontScheme ?? {})) {
      add(family.latinTypeface, theme.sourcePart);
      add(family.eastAsianTypeface, theme.sourcePart);
      add(family.complexScriptTypeface, theme.sourcePart);
    }
  }
  return [...counts.values()]
    .map((entry) => ({ ...entry, sources: [...entry.sources].sort() }))
    .sort((left, right) => right.count - left.count || left.typeface.localeCompare(right.typeface));
}

function rectangleFromBbox(bbox, slideWidth, slideHeight) {
  if (!Array.isArray(bbox) || bbox.length < 4) return null;
  const [left, top, width, height] = bbox.map(Number);
  if (![left, top, width, height].every(Number.isFinite) || width <= 0 || height <= 0) return null;
  const x1 = Math.max(0, Math.min(slideWidth, left));
  const y1 = Math.max(0, Math.min(slideHeight, top));
  const x2 = Math.max(0, Math.min(slideWidth, left + width));
  const y2 = Math.max(0, Math.min(slideHeight, top + height));
  if (x2 <= x1 || y2 <= y1) return null;
  return { x1, y1, x2, y2 };
}

function isLineLikeElement(element, inspectRecord) {
  const values = [
    element?.kind,
    inspectRecord?.kind,
    element?.geometry,
    element?.shapeType,
    element?.connectorType,
    inspectRecord?.geometry,
    inspectRecord?.shapeType,
  ]
    .filter((value) => value != null)
    .map((value) => String(value).toLowerCase());
  return values.some((value) => (
    value === "line"
    || value.includes("connector")
    || value.includes("polyline")
    || value.includes("straightline")
  ));
}

function finiteStrokeWidthPx(element, inspectRecord) {
  const candidates = [
    element?.lineWidth,
    element?.strokeWidth,
    element?.line?.width,
    element?.stroke?.width,
    inspectRecord?.lineWidth,
    inspectRecord?.strokeWidth,
  ];
  for (const candidate of candidates) {
    const value = Math.abs(Number(candidate));
    if (Number.isFinite(value) && value > 0) return value;
  }
  return MIN_VISIBLE_STROKE_PX;
}

function bboxRecordPx(bbox) {
  if (!Array.isArray(bbox) || bbox.length < 4) return null;
  const [left, top, width, height] = bbox.map(Number);
  if (![left, top, width, height].every(Number.isFinite)) return null;
  return { left, top, width, height, unit: "px" };
}

function effectiveBboxForElement(bbox, element, inspectRecord, slideWidth, slideHeight) {
  const raw = bboxRecordPx(bbox);
  if (!raw) {
    return {
      raw: null,
      effective: null,
      adjustedForStroke: false,
      isLineLike: false,
      strokeWidthPx: null,
    };
  }
  const isLineLike = isLineLikeElement(element, inspectRecord);
  if (!isLineLike || (raw.width > 0 && raw.height > 0)) {
    return {
      raw,
      effective: raw,
      adjustedForStroke: false,
      isLineLike,
      strokeWidthPx: isLineLike ? finiteStrokeWidthPx(element, inspectRecord) : null,
    };
  }

  const strokeWidthPx = Math.max(
    MIN_VISIBLE_STROKE_PX,
    finiteStrokeWidthPx(element, inspectRecord),
  );
  let left = raw.left;
  let top = raw.top;
  let width = raw.width;
  let height = raw.height;
  if (width <= 0) {
    const start = Math.max(0, Math.min(slideWidth, raw.left - strokeWidthPx / 2));
    const end = Math.max(0, Math.min(slideWidth, raw.left + strokeWidthPx / 2));
    left = start;
    width = end - start;
  }
  if (height <= 0) {
    const start = Math.max(0, Math.min(slideHeight, raw.top - strokeWidthPx / 2));
    const end = Math.max(0, Math.min(slideHeight, raw.top + strokeWidthPx / 2));
    top = start;
    height = end - start;
  }
  if (!(width > 0) || !(height > 0)) {
    return {
      raw,
      effective: raw,
      adjustedForStroke: false,
      isLineLike,
      strokeWidthPx,
    };
  }
  return {
    raw,
    effective: { left, top, width, height, unit: "px" },
    adjustedForStroke: true,
    isLineLike,
    strokeWidthPx,
  };
}

function unionArea(rectangles) {
  if (rectangles.length === 0) return 0;
  const xs = [...new Set(rectangles.flatMap((rect) => [rect.x1, rect.x2]))].sort((a, b) => a - b);
  let area = 0;
  for (let index = 0; index < xs.length - 1; index += 1) {
    const left = xs[index];
    const right = xs[index + 1];
    if (right <= left) continue;
    const intervals = rectangles
      .filter((rect) => rect.x1 < right && rect.x2 > left)
      .map((rect) => [rect.y1, rect.y2])
      .sort((a, b) => a[0] - b[0]);
    let coveredY = 0;
    let currentStart = null;
    let currentEnd = null;
    for (const [start, end] of intervals) {
      if (currentStart === null) {
        currentStart = start;
        currentEnd = end;
      } else if (start <= currentEnd) {
        currentEnd = Math.max(currentEnd, end);
      } else {
        coveredY += currentEnd - currentStart;
        currentStart = start;
        currentEnd = end;
      }
    }
    if (currentStart !== null) coveredY += currentEnd - currentStart;
    area += (right - left) * coveredY;
  }
  return area;
}

function roundRatio(value) {
  return Number(Math.max(0, Math.min(1, value)).toFixed(4));
}

function pxToInches(value) {
  return Number((Number(value) / PX_PER_INCH).toFixed(4));
}

function bboxPxToInches(bbox) {
  if (!bbox) return null;
  return {
    x: pxToInches(bbox.left),
    y: pxToInches(bbox.top),
    width: pxToInches(bbox.width),
    height: pxToInches(bbox.height),
    unit: "in",
  };
}

function bboxPxEvidence(bbox) {
  if (!bbox) return null;
  return {
    x: Number(bbox.left),
    y: Number(bbox.top),
    width: Number(bbox.width),
    height: Number(bbox.height),
    unit: "px",
  };
}

function classifySlide({ imageCoverageRatio, largestImageCoverageRatio, imageCount, nativeObjectCount, nativeTextChars, largestRasterLowResolution }) {
  if (imageCount === 0 || (imageCoverageRatio < 0.08 && largestImageCoverageRatio < 0.06)) {
    return {
      classification: "native-editable",
      confidence: imageCount === 0 ? 0.96 : 0.84,
      reasons: imageCount === 0
        ? ["No slide-local raster images were detected."]
        : ["Raster coverage is small and appears decorative."],
    };
  }
  if (imageCoverageRatio >= 0.82 && largestImageCoverageRatio >= 0.78) {
    if (largestRasterLowResolution) {
      return {
        classification: "low-quality-scan",
        confidence: 0.9,
        reasons: [
          "A single raster covers most of the slide.",
          "The raster has fewer pixels than its rendered slide footprint.",
        ],
      };
    }
    return {
      classification: "flattened",
      confidence: nativeObjectCount <= 3 && nativeTextChars < 80 ? 0.96 : 0.86,
      reasons: [
        "A single raster covers most of the slide.",
        nativeObjectCount <= 3
          ? "Few slide-local editable objects remain."
          : "Some editable overlays remain above the dominant raster.",
      ],
    };
  }
  if (imageCoverageRatio >= 0.18 || largestImageCoverageRatio >= 0.16) {
    return {
      classification: "mixed",
      confidence: 0.84,
      reasons: [
        "The slide combines meaningful raster coverage with editable content.",
        `${nativeObjectCount} slide-local non-raster object(s) were detected.`,
      ],
    };
  }
  return {
    classification: "native-editable",
    confidence: 0.78,
    reasons: ["Raster content is limited; native editable structure dominates."],
  };
}

function parseInspectNdjson(ndjson, warnings) {
  const records = [];
  const lines = String(ndjson ?? "").split(/\r?\n/);
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index].trim();
    if (!line) continue;
    try {
      records.push(JSON.parse(line));
    } catch (error) {
      warnings.push(`Could not parse inspect line ${index + 1}: ${error.message}`);
    }
  }
  return records;
}

function walkElements(layout) {
  const result = [];
  const visit = (value) => {
    if (!value || typeof value !== "object") return;
    if (Array.isArray(value)) {
      for (const child of value) visit(child);
      return;
    }
    if (typeof value.kind === "string" && Array.isArray(value.bbox) && value.scope === "slide") {
      result.push(value);
    }
    for (const child of Object.values(value)) visit(child);
  };
  visit(layout.elements ?? []);
  return result;
}

function firstDefined(...values) {
  return values.find((value) => value !== undefined && value !== null && value !== "");
}

function normalizedStyleColor(value) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  if (/^#[0-9a-f]{6}$/i.test(text)) return text.toUpperCase();
  return text;
}

function visualStyleEvidence(layoutJsons) {
  const textClusters = new Map();
  const roleCandidates = [];
  const componentClusters = new Map();
  for (const layout of layoutJsons) {
    const slideNumber = Number(layout.slide?.slide);
    const frame = layout.slide?.frame ?? { width: 1280, height: 720 };
    for (const element of walkElements(layout)) {
      const bbox = Array.isArray(element.bbox) ? element.bbox.map(Number) : null;
      if (!bbox || bbox.length < 4 || !bbox.every(Number.isFinite)) continue;
      const [left, top, width, height] = bbox;
      const text = String(element.text ?? element.textPreview ?? "").trim();
      const style = element.resolvedTextStyle ?? {};
      const firstParagraph = Array.isArray(element.paragraphs) ? element.paragraphs[0] : null;
      const firstRun = Array.isArray(firstParagraph?.runs) ? firstParagraph.runs[0] : null;
      const runStyle = firstRun?.resolvedTextStyle ?? firstRun?.style ?? {};
      if (text) {
        const typeface = firstDefined(style.typeface, runStyle.typeface, element.typeface, null);
        const fontSize = Number(firstDefined(style.fontSize, runStyle.fontSize, element.fontSize, NaN));
        const bold = firstDefined(style.bold, runStyle.bold, element.bold, null);
        const color = normalizedStyleColor(firstDefined(style.color, runStyle.color, element.textColor, null));
        const alignment = firstDefined(style.alignment, firstParagraph?.alignment, null);
        const lineSpacing = Number(firstDefined(style.lineSpacing, firstParagraph?.lineSpacing, NaN));
        const signature = JSON.stringify({
          typeface: typeface ?? null,
          fontSize: Number.isFinite(fontSize) ? Number(fontSize.toFixed(2)) : null,
          bold: bold == null ? null : Boolean(bold),
          color,
          alignment: alignment ?? null,
          lineSpacing: Number.isFinite(lineSpacing) ? Number(lineSpacing.toFixed(3)) : null,
        });
        const cluster = textClusters.get(signature) ?? {
          style: JSON.parse(signature),
          count: 0,
          objectIds: [],
          slides: new Set(),
        };
        cluster.count += 1;
        cluster.objectIds.push(element.aid ?? element.id ?? null);
        cluster.slides.add(slideNumber);
        textClusters.set(signature, cluster);

        const name = String(element.name ?? "");
        let roleHint = "body";
        let confidence = 0.55;
        if (/slide number|页码|编号/i.test(name)) {
          roleHint = "footer";
          confidence = 0.9;
        } else if (/title|标题|headline/i.test(name) || top <= Number(frame.height) * 0.16) {
          roleHint = "slide-title";
          confidence = /title|标题|headline/i.test(name) ? 0.88 : 0.7;
        } else if (/caption|注释|说明/i.test(name)) {
          roleHint = "caption";
          confidence = 0.78;
        }
        roleCandidates.push({
          slideNumber,
          objectId: element.aid ?? element.id ?? null,
          name,
          roleHint,
          confidence,
          bboxPx: { x: left, y: top, width, height, unit: "px" },
          style: JSON.parse(signature),
          textPreview: text.slice(0, 160),
          requiresVisualConfirmation: true,
        });
      }

      const geometry = String(element.geometry ?? element.kind ?? "unknown").toLowerCase();
      if (element.kind === "shape" && width > 0 && height > 0) {
        const signature = `${geometry}|${Math.round(width / 4) * 4}|${Math.round(height / 4) * 4}|${normalizedStyleColor(element.fillColor) ?? "none"}|${Number(element.lineWidth ?? 0).toFixed(2)}`;
        const cluster = componentClusters.get(signature) ?? {
          signature,
          geometry,
          approximateSizePx: { width: Math.round(width / 4) * 4, height: Math.round(height / 4) * 4 },
          fillColor: normalizedStyleColor(element.fillColor),
          lineWidthPx: Number(element.lineWidth ?? 0),
          instances: [],
        };
        cluster.instances.push({ slideNumber, objectId: element.aid ?? element.id ?? null, name: element.name ?? "", bboxPx: { x: left, y: top, width, height, unit: "px" } });
        componentClusters.set(signature, cluster);
      }
    }
  }
  return {
    textStyleClusters: [...textClusters.values()]
      .map((entry) => ({ ...entry, slides: [...entry.slides].sort((a, b) => a - b), objectIds: entry.objectIds.filter(Boolean) }))
      .sort((a, b) => b.count - a.count),
    roleCandidates,
    componentCandidates: [...componentClusters.values()]
      .filter((entry) => entry.instances.length >= 2)
      .sort((a, b) => b.instances.length - a.instances.length),
    advisoryOnly: true,
  };
}

function buildPackageUsageContext(zipNames, relationships, slideParts) {
  const normalizedRelationships = relationships
    .filter((relationship) => relationship && !relationship.external && relationship.resolvedTarget)
    .map((relationship) => ({
      ...relationship,
      ownerPart: normalizeZipPath(relationship.ownerPart ?? relationship.sourcePart),
      resolvedTarget: normalizeZipPath(relationship.resolvedTarget),
    }));
  const byOwner = new Map();
  const byTarget = new Map();
  for (const relationship of normalizedRelationships) {
    const ownerValues = byOwner.get(relationship.ownerPart) ?? [];
    ownerValues.push(relationship);
    byOwner.set(relationship.ownerPart, ownerValues);
    const targetValues = byTarget.get(relationship.resolvedTarget) ?? [];
    targetValues.push(relationship);
    byTarget.set(relationship.resolvedTarget, targetValues);
  }

  const slideNumberByPart = new Map(
    slideParts.map((partName, index) => [normalizeZipPath(partName), index + 1]),
  );
  const allLayouts = new Set(zipNames
    .filter((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(name))
    .map(normalizeZipPath));
  const allMasters = new Set(zipNames
    .filter((name) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(name))
    .map(normalizeZipPath));
  const layoutToSlides = new Map();
  const slideToLayout = new Map();
  for (const [slidePart, slideNumber] of slideNumberByPart) {
    const layoutRelationship = (byOwner.get(slidePart) ?? []).find((relationship) => (
      /\/slideLayout$/i.test(String(relationship.type ?? ""))
      || /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(relationship.resolvedTarget)
    ));
    if (!layoutRelationship) continue;
    const layoutPart = layoutRelationship.resolvedTarget;
    slideToLayout.set(slidePart, layoutPart);
    const values = layoutToSlides.get(layoutPart) ?? [];
    values.push(slideNumber);
    layoutToSlides.set(layoutPart, values);
  }

  const masterToSlides = new Map();
  const layoutToMaster = new Map();
  for (const [layoutPart, slideNumbers] of layoutToSlides) {
    const masterRelationship = (byOwner.get(layoutPart) ?? []).find((relationship) => (
      /\/slideMaster$/i.test(String(relationship.type ?? ""))
      || /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(relationship.resolvedTarget)
    ));
    if (!masterRelationship) continue;
    const masterPart = masterRelationship.resolvedTarget;
    layoutToMaster.set(layoutPart, masterPart);
    const values = new Set(masterToSlides.get(masterPart) ?? []);
    for (const slideNumber of slideNumbers) values.add(slideNumber);
    masterToSlides.set(masterPart, [...values].sort((a, b) => a - b));
  }

  const activeLayouts = new Set(layoutToSlides.keys());
  const activeMasters = new Set(masterToSlides.keys());

  const scopeForParts = (rawParts) => {
    const parts = [...new Set(rawParts.filter(Boolean).map(normalizeZipPath))];
    const activeSlides = new Set();
    const activeLayoutParts = new Set();
    const activeMasterParts = new Set();
    const unusedLayouts = new Set();
    const unusedMasters = new Set();
    const referenceEvidence = [];

    const registerOwner = (ownerPart, relationship = null) => {
      const normalizedOwner = normalizeZipPath(ownerPart);
      if (slideNumberByPart.has(normalizedOwner)) {
        activeSlides.add(slideNumberByPart.get(normalizedOwner));
      } else if (layoutToSlides.has(normalizedOwner)) {
        activeLayoutParts.add(normalizedOwner);
        for (const slideNumber of layoutToSlides.get(normalizedOwner)) activeSlides.add(slideNumber);
      } else if (masterToSlides.has(normalizedOwner)) {
        activeMasterParts.add(normalizedOwner);
        for (const slideNumber of masterToSlides.get(normalizedOwner)) activeSlides.add(slideNumber);
      } else if (allLayouts.has(normalizedOwner)) {
        unusedLayouts.add(normalizedOwner);
      } else if (allMasters.has(normalizedOwner)) {
        unusedMasters.add(normalizedOwner);
      }
      if (relationship) {
        referenceEvidence.push({
          relationshipPart: relationship.sourcePart ?? null,
          ownerPart: normalizedOwner,
          targetPart: relationship.resolvedTarget ?? null,
          type: relationship.type ?? null,
        });
      }
    };

    for (const partName of parts) {
      registerOwner(partName);
      for (const relationship of byTarget.get(partName) ?? []) {
        // A master owns every layout in its template family. A slide that uses
        // that master does not thereby depend on every sibling layout.
        if (
          allLayouts.has(partName)
          && allMasters.has(relationship.ownerPart)
          && /\/slideLayout$/i.test(String(relationship.type ?? ""))
        ) {
          continue;
        }
        registerOwner(relationship.ownerPart, relationship);
      }
    }

    const activeSlideList = [...activeSlides].sort((a, b) => a - b);
    const activeLayoutList = [...activeLayoutParts].sort(naturalPartSort);
    const activeMasterList = [...activeMasterParts].sort(naturalPartSort);
    const unusedLayoutList = [...unusedLayouts].sort(naturalPartSort);
    const unusedMasterList = [...unusedMasters].sort(naturalPartSort);
    const activeDependency = activeSlideList.length > 0;
    const unusedTemplateMetadata = !activeDependency
      && (unusedLayoutList.length > 0 || unusedMasterList.length > 0);
    return {
      scope: activeDependency
        ? "active-slide-dependency"
        : unusedTemplateMetadata
          ? "unused-template-metadata"
          : "package-metadata",
      activeDependency,
      unusedTemplateMetadata,
      activeSlides: activeSlideList,
      activeLayouts: activeLayoutList,
      activeMasters: activeMasterList,
      unusedLayouts: unusedLayoutList,
      unusedMasters: unusedMasterList,
      references: referenceEvidence,
    };
  };

  return {
    slideNumberByPart,
    slideToLayout,
    layoutToSlides,
    layoutToMaster,
    masterToSlides,
    activeLayouts,
    activeMasters,
    scopeForParts,
  };
}

function textPreviewFromXml(xml, maximumLength = 240) {
  const values = [];
  for (const match of String(xml ?? "").matchAll(/<(?:a:)?t\b[^>]*>([\s\S]*?)<\/(?:a:)?t>/gi)) {
    const value = xmlDecode(match[1]).replace(/\s+/g, " ").trim();
    if (value) values.push(value);
  }
  const joined = values.join(" ");
  return joined.length > maximumLength ? `${joined.slice(0, maximumLength - 1)}…` : joined;
}

function detectSpecialFeatures(zipNames, relevantXmlParts, relationships, usageContext) {
  const lowerNames = zipNames.map((name) => name.toLowerCase());
  const xmlByPart = new Map(relevantXmlParts);
  const xmlMatches = (pattern) => relevantXmlParts
    .filter(([, xml]) => pattern.test(xml))
    .map(([partName]) => partName);
  const pathMatches = (predicate) => zipNames.filter((name, index) => predicate(lowerNames[index], name));
  const external = relationships.filter((relationship) => relationship.external);
  const relType = (relationship) => String(relationship.type ?? "").toLowerCase();
  const linkedMedia = external.filter((relationship) => /(?:image|audio|video|media)$/.test(relType(relationship)));
  const hyperlinks = external.filter((relationship) => /hyperlink$/.test(relType(relationship)));
  const legacyVector = pathMatches((name) => /\.(?:emf|wmf|eps)$/.test(name));
  const mediaParts = pathMatches((name) => /^ppt\/media\/.*\.(?:mp3|wav|m4a|aac|wma|mp4|m4v|mov|avi|wmv|mpeg|mpg)$/i.test(name));
  const noteParts = pathMatches((name) => /^ppt\/notesslides\/notesslide\d+\.xml$/i.test(name));
  const noteDetails = noteParts.map((partName) => {
    const textPreview = textPreviewFromXml(xmlByPart.get(partName));
    return { part: partName, textPreview, hasText: Boolean(textPreview) };
  });
  const commentParts = pathMatches((name) => (
    /^ppt\/comments\/.*\.xml$/i.test(name)
    || /^ppt\/(?:moderncomments|commentsmodern)\/.*\.xml$/i.test(name)
  ));
  const commentAuthorParts = pathMatches((name) => (
    /^ppt\/commentauthors\.xml$/i.test(name)
    || /^ppt\/(?:commentauthors|authors)\/.*\.xml$/i.test(name)
  ));
  const tagParts = pathMatches((name) => /^ppt\/tags\/.*\.xml$/i.test(name));
  const wdpParts = pathMatches((name) => /^ppt\/media\/.*\.wdp$/i.test(name));
  const knownMediaExtension = /\.(?:png|jpe?g|gif|bmp|tiff?|webp|svg|emf|wmf|eps|wdp|mp3|wav|m4a|aac|wma|mp4|m4v|mov|avi|wmv|mpeg|mpg)$/i;
  const unsupportedMediaParts = pathMatches((name) => name.startsWith("ppt/media/") && !knownMediaExtension.test(name));

  const make = ({
    objectParts = [],
    evidenceParts = [],
    scopeParts = [],
    details = [],
    blockingPolicy = "none",
    manualReview = false,
    reviewReason = null,
  } = {}) => {
    const uniqueObjects = [...new Set(objectParts)].sort(naturalPartSort);
    const uniqueEvidence = [...new Set(evidenceParts)].sort(naturalPartSort);
    const uniqueParts = [...new Set([...uniqueObjects, ...uniqueEvidence])].sort(naturalPartSort);
    const present = uniqueParts.length > 0 || details.length > 0;
    const scope = usageContext.scopeForParts([...uniqueParts, ...scopeParts]);
    const blocking = present && (
      blockingPolicy === "always"
      || (blockingPolicy === "active" && scope.activeDependency)
    );
    return {
      present,
      count: uniqueObjects.length || uniqueEvidence.length || details.length,
      objectCount: uniqueObjects.length,
      referenceCount: scope.references.length,
      parts: uniqueParts,
      objectParts: uniqueObjects,
      evidenceParts: uniqueEvidence,
      details,
      ...scope,
      blocking,
      manualReview: present && (manualReview || blocking),
      reviewReason: present ? reviewReason : null,
    };
  };

  const commentAuthorDetails = commentAuthorParts.map((partName) => {
    const xml = String(xmlByPart.get(partName) ?? "");
    return {
      part: partName,
      authorCount: [...xml.matchAll(/<(?:[A-Za-z0-9_.-]+:)?cmAuthor\b/gi)].length,
    };
  });
  const commentAuthorEntryCount = commentAuthorDetails.reduce((sum, detail) => sum + detail.authorCount, 0);
  const commentAuthorsFeature = make({
    objectParts: commentAuthorParts,
    details: commentAuthorDetails,
    manualReview: true,
    reviewReason: "Comment-author history is package metadata that may be dropped even when no comment parts remain.",
  });
  if (commentAuthorsFeature.present) {
    commentAuthorsFeature.partCount = commentAuthorsFeature.objectParts.length;
    commentAuthorsFeature.entryCount = commentAuthorEntryCount;
    commentAuthorsFeature.objectCount = commentAuthorEntryCount;
    commentAuthorsFeature.count = commentAuthorEntryCount;
  }

  return {
    macros: make({
      objectParts: pathMatches((name) => name.includes("vbaproject") || name.endsWith(".vba")),
      blockingPolicy: "always",
      manualReview: true,
      reviewReason: "Macro-bearing packages require an explicit preservation workflow.",
    }),
    ole: make({
      objectParts: pathMatches((name) => name.startsWith("ppt/embeddings/")),
      evidenceParts: xmlMatches(/<(?:p:)?oleObj\b/i),
      blockingPolicy: "active",
      manualReview: true,
      reviewReason: "Active OLE dependencies block reconstruction; unused template OLE still requires round-trip review.",
    }),
    activeX: make({
      objectParts: pathMatches((name) => name.startsWith("ppt/activex/")),
      evidenceParts: xmlMatches(/activeX/i),
      blockingPolicy: "active",
      manualReview: true,
      reviewReason: "ActiveX may not survive import/export and must be reviewed explicitly.",
    }),
    smartArt: make({
      objectParts: pathMatches((name) => name.startsWith("ppt/diagrams/")),
      evidenceParts: xmlMatches(/schemas\.openxmlformats\.org\/drawingml\/2006\/diagram|<dgm:/i),
      blockingPolicy: "active",
      manualReview: true,
      reviewReason: "Active SmartArt must be preserved or converted under an explicit exception.",
    }),
    media: make({
      objectParts: mediaParts,
      blockingPolicy: "active",
      manualReview: true,
      reviewReason: "Audio and video require active-dependency and round-trip preservation checks.",
    }),
    linkedMedia: make({
      scopeParts: linkedMedia.map((relationship) => relationship.ownerPart),
      blockingPolicy: "active",
      manualReview: true,
      reviewReason: "Linked media can break when the deck is copied or exported.",
      details: linkedMedia.map((relationship) => ({
        sourcePart: relationship.sourcePart,
        ownerPart: relationship.ownerPart ?? null,
        target: relationship.target,
        type: relationship.type,
      })),
    }),
    externalLinks: make({
      scopeParts: external.map((relationship) => relationship.ownerPart),
      manualReview: true,
      reviewReason: "External relationships must be checked for portability and preservation.",
      details: external.map((relationship) => ({
        sourcePart: relationship.sourcePart,
        ownerPart: relationship.ownerPart ?? null,
        target: relationship.target,
        type: relationship.type,
      })),
    }),
    hyperlinks: make({
      scopeParts: hyperlinks.map((relationship) => relationship.ownerPart),
      details: hyperlinks.map((relationship) => ({
        sourcePart: relationship.sourcePart,
        ownerPart: relationship.ownerPart ?? null,
        target: relationship.target,
      })),
    }),
    transitions: make({
      evidenceParts: xmlMatches(/<(?:p:)?transition\b/i),
      manualReview: true,
      reviewReason: "Transitions are reported by active or unused layout scope and require visual round-trip review.",
    }),
    animations: make({
      evidenceParts: xmlMatches(/<(?:p:)?timing\b/i),
      manualReview: true,
      reviewReason: "Animation timing may be lost during reconstruction.",
    }),
    equations: make({
      evidenceParts: xmlMatches(/<(?:m:)?oMath(?:Para)?\b|office\/2006\/math/i),
      blockingPolicy: "active",
      manualReview: true,
      reviewReason: "Active equations require a preservation decision.",
    }),
    ink: make({
      objectParts: pathMatches((name) => name.startsWith("ppt/ink/") || name.endsWith(".inkml")),
      evidenceParts: xmlMatches(/inkml|<p:contentPart\b/i),
      blockingPolicy: "active",
      manualReview: true,
      reviewReason: "Active Ink content requires a preservation decision.",
    }),
    notes: make({
      objectParts: noteParts,
      manualReview: noteDetails.some((note) => note.hasText),
      reviewReason: "Speaker notes are delivery content and must survive the round trip.",
      details: noteDetails,
    }),
    comments: make({
      objectParts: commentParts,
      manualReview: true,
      reviewReason: "Classic or modern comments may not survive reconstruction.",
    }),
    commentAuthors: commentAuthorsFeature,
    tags: make({
      objectParts: tagParts,
      manualReview: true,
      reviewReason: "Custom slide tags can carry workflow metadata and require preservation review.",
    }),
    customShows: make({
      evidenceParts: xmlMatches(/<(?:p:)?custShow(?:Lst)?\b/i),
      manualReview: true,
      reviewReason: "Custom shows depend on stable slide identities and ordering.",
    }),
    customXml: make({
      objectParts: pathMatches((name) => name.startsWith("customxml/")),
      manualReview: true,
      reviewReason: "Custom XML parts may be application metadata and require explicit preservation review.",
    }),
    embeddedFonts: make({
      objectParts: pathMatches((name) => name.startsWith("ppt/fonts/") || name.endsWith(".odttf")),
      evidenceParts: xmlMatches(/<(?:p:)?embeddedFont(?:Lst)?\b/i),
      manualReview: true,
      reviewReason: "Embedded fonts require licensing and round-trip preservation checks.",
    }),
    svg: make({
      objectParts: pathMatches((name) => /^ppt\/media\/.*\.svg$/i.test(name)),
      manualReview: true,
      reviewReason: "SVG assets must be checked for import/export fidelity.",
    }),
    wdp: make({
      objectParts: wdpParts,
      blockingPolicy: "active",
      manualReview: true,
      reviewReason: "WDP/HD Photo is not a standard raster path in the editing runtime; active inherited use blocks reconstruction.",
    }),
    unsupportedMedia: make({
      objectParts: unsupportedMediaParts,
      blockingPolicy: "active",
      manualReview: true,
      reviewReason: "Unknown media formats require an explicit preservation or conversion decision.",
    }),
    legacyVector: make({
      objectParts: legacyVector,
      blockingPolicy: "active",
      manualReview: true,
      reviewReason: "Legacy vector assets require active-dependency and render-fidelity review.",
    }),
    emf: make({
      objectParts: legacyVector.filter((name) => name.toLowerCase().endsWith(".emf")),
      blockingPolicy: "active",
      manualReview: true,
      reviewReason: "EMF assets require preservation review.",
    }),
    wmf: make({
      objectParts: legacyVector.filter((name) => name.toLowerCase().endsWith(".wmf")),
      blockingPolicy: "active",
      manualReview: true,
      reviewReason: "WMF assets require preservation review.",
    }),
    eps: make({
      objectParts: legacyVector.filter((name) => name.toLowerCase().endsWith(".eps")),
      blockingPolicy: "active",
      manualReview: true,
      reviewReason: "EPS assets require preservation review.",
    }),
  };
}

async function inspectRasterMetadata(sharpFactory, bytes, contentType, warnings, sourcePart) {
  if (!sharpFactory || !isRasterContentType(contentType)) return null;
  try {
    const metadata = await sharpFactory(bytes, { limitInputPixels: false }).metadata();
    return {
      width: metadata.width ?? null,
      height: metadata.height ?? null,
      density: metadata.density ?? null,
      format: metadata.format ?? null,
      space: metadata.space ?? null,
      channels: metadata.channels ?? null,
      hasAlpha: metadata.hasAlpha ?? null,
      orientation: metadata.orientation ?? null,
    };
  } catch (error) {
    warnings.push(`Could not inspect raster metadata for ${sourcePart}: ${error.message}`);
    return null;
  }
}

async function createVerifiedMontage({
  sharpFactory,
  slideArtifactPaths,
  slideCount,
  workspace,
  montagePath,
  warnings,
}) {
  const evidencePath = path.join(workspace, "source-montage.json");
  await fs.rm(montagePath, { force: true }).catch(() => {});
  await fs.rm(evidencePath, { force: true }).catch(() => {});
  if (!sharpFactory) {
    warnings.push("Verified montage was not produced because sharp is unavailable.");
    return { available: false, path: null, evidencePath: null, evidence: null };
  }
  if (slideCount < 1 || slideArtifactPaths.length !== slideCount) {
    warnings.push(
      `Verified montage was not produced: expected ${slideCount} rendered slide(s), found ${slideArtifactPaths.length}.`,
    );
    return { available: false, path: null, evidencePath: null, evidence: null };
  }

  const tileWidth = 320;
  const previewHeight = 200;
  const labelHeight = 24;
  const tileHeight = previewHeight + labelHeight;
  const columns = Math.max(1, Math.ceil(Math.sqrt(slideCount * (16 / 9))));
  const rows = Math.ceil(slideCount / columns);
  const canvasWidth = columns * tileWidth;
  const canvasHeight = rows * tileHeight;
  const tileBuffers = [];
  const tiles = [];

  try {
    for (const artifact of slideArtifactPaths) {
      const renderBytes = await fs.readFile(artifact.renderPath);
      const renderMetadata = await sharpFactory(renderBytes, { limitInputPixels: false }).metadata();
      if (!(renderMetadata.width > 0) || !(renderMetadata.height > 0)) {
        throw new Error(`slide ${artifact.slideNumber} render has no valid dimensions`);
      }
      const preview = await sharpFactory(renderBytes, { limitInputPixels: false })
        .resize(tileWidth, previewHeight, {
          fit: "contain",
          background: { r: 255, g: 255, b: 255, alpha: 1 },
        })
        .ensureAlpha()
        .png()
        .toBuffer();
      const label = Buffer.from(
        `<svg xmlns="http://www.w3.org/2000/svg" width="${tileWidth}" height="${labelHeight}">`
          + `<rect width="100%" height="100%" fill="#eef2f7"/>`
          + `<text x="${tileWidth / 2}" y="17" text-anchor="middle" font-family="Arial, sans-serif" font-size="13" fill="#334155">`
          + `Slide ${artifact.slideNumber}</text></svg>`,
      );
      const tileBuffer = await sharpFactory({
        create: {
          width: tileWidth,
          height: tileHeight,
          channels: 4,
          background: { r: 255, g: 255, b: 255, alpha: 1 },
        },
      })
        .composite([
          { input: preview, left: 0, top: 0 },
          { input: label, left: 0, top: previewHeight },
        ])
        .png()
        .toBuffer();
      const column = (artifact.slideNumber - 1) % columns;
      const row = Math.floor((artifact.slideNumber - 1) / columns);
      const rawTile = await sharpFactory(tileBuffer).ensureAlpha().raw().toBuffer();
      tileBuffers.push({
        input: tileBuffer,
        left: column * tileWidth,
        top: row * tileHeight,
      });
      tiles.push({
        slideNumber: artifact.slideNumber,
        sourceRender: relativeArtifact(workspace, artifact.renderPath),
        sourceSha256: sha256(renderBytes),
        sourceDimensionsPx: {
          width: renderMetadata.width,
          height: renderMetadata.height,
          unit: "px",
        },
        grid: { row: row + 1, column: column + 1 },
        bboxPx: {
          x: column * tileWidth,
          y: row * tileHeight,
          width: tileWidth,
          height: tileHeight,
          unit: "px",
        },
        tilePixelSha256: sha256(rawTile),
      });
    }

    const montageBytes = await sharpFactory({
      create: {
        width: canvasWidth,
        height: canvasHeight,
        channels: 4,
        background: { r: 226, g: 232, b: 240, alpha: 1 },
      },
    })
      .composite(tileBuffers)
      .webp({ lossless: true, quality: 100, effort: 4 })
      .toBuffer();
    const montageMetadata = await sharpFactory(montageBytes).metadata();
    if (montageMetadata.width !== canvasWidth || montageMetadata.height !== canvasHeight) {
      throw new Error(
        `montage canvas mismatch (${montageMetadata.width}x${montageMetadata.height}, expected ${canvasWidth}x${canvasHeight})`,
      );
    }

    for (const tile of tiles) {
      const extractedRaw = await sharpFactory(montageBytes, { limitInputPixels: false })
        .extract({
          left: tile.bboxPx.x,
          top: tile.bboxPx.y,
          width: tile.bboxPx.width,
          height: tile.bboxPx.height,
        })
        .ensureAlpha()
        .raw()
        .toBuffer();
      if (sha256(extractedRaw) !== tile.tilePixelSha256) {
        throw new Error(`lossless tile verification failed for slide ${tile.slideNumber}`);
      }
    }

    await fs.writeFile(montagePath, montageBytes);
    const evidence = {
      schemaVersion: "1.1",
      generatedAt: new Date().toISOString(),
      pageCount: slideCount,
      tileCount: tiles.length,
      allSlidesIncluded: tiles.length === slideCount
        && tiles.every((tile, index) => tile.slideNumber === index + 1),
      canvasPx: { width: canvasWidth, height: canvasHeight, unit: "px" },
      grid: { rows, columns, tileWidth, tileHeight, previewHeight, labelHeight },
      montageSha256: sha256(montageBytes),
      verification: {
        method: "lossless-webp-tile-pixel-sha256",
        allTilesMatched: true,
      },
      tiles,
    };
    if (!evidence.allSlidesIncluded) {
      throw new Error("montage tile ledger does not contain every slide exactly once");
    }
    await writeJson(evidencePath, evidence);
    return { available: true, path: montagePath, evidencePath, evidence };
  } catch (error) {
    await fs.rm(montagePath, { force: true }).catch(() => {});
    await fs.rm(evidencePath, { force: true }).catch(() => {});
    warnings.push(`Verified montage was not produced: ${error.message}`);
    return { available: false, path: null, evidencePath: null, evidence: null };
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    process.stdout.write(HELP);
    return;
  }

  const sourcePath = path.resolve(args.pptx);
  const workspace = path.resolve(args.workspace);
  if (path.extname(sourcePath).toLowerCase() !== ".pptx") {
    throw new CliError(`Only .pptx input is supported: ${sourcePath}`);
  }
  let sourceStat;
  try {
    sourceStat = await fs.stat(sourcePath);
  } catch {
    throw new CliError(`Source PPTX not found: ${sourcePath}`);
  }
  if (!sourceStat.isFile()) {
    throw new CliError(`Source PPTX is not a regular file: ${sourcePath}`);
  }
  if (path.normalize(sourcePath) === path.normalize(workspace)) {
    throw new CliError("--workspace must be a directory distinct from the source file.");
  }

  await fs.mkdir(workspace, { recursive: true });
  const workspaceStat = await fs.stat(workspace);
  if (!workspaceStat.isDirectory()) {
    throw new CliError(`Workspace is not a directory: ${workspace}`);
  }

  const renderDir = path.join(workspace, "source-renders");
  const layoutDir = path.join(workspace, "source-layout");
  const mediaDir = path.join(workspace, "source-media");
  await resetDirectory(workspace, renderDir);
  await resetDirectory(workspace, layoutDir);
  await resetDirectory(workspace, mediaDir);

  const sourceHashBefore = await hashFile(sourcePath);
  const sourceCopyPath = path.join(workspace, "source-original.pptx");
  if (path.normalize(sourcePath).toLowerCase() === path.normalize(sourceCopyPath).toLowerCase()) {
    throw new CliError("The source file cannot be the analyzer's source-original.pptx output.");
  }
  await fs.copyFile(sourcePath, sourceCopyPath);
  const copyHash = await hashFile(sourceCopyPath);
  if (copyHash !== sourceHashBefore) {
    throw new CliError("The working copy hash does not match the source hash.", 1);
  }

  const warnings = [];
  const jszipModule = await loadRuntimeModule("jszip");
  const JSZip = jszipModule.default ?? jszipModule;
  const sharpModule = await loadRuntimeModule("sharp", { optional: true });
  const sharpFactory = sharpModule ? sharpModule.default ?? sharpModule : null;
  if (!sharpFactory) warnings.push("sharp is unavailable; raster dimensions and low-resolution hints were skipped.");

  const sourceBytes = await fs.readFile(sourceCopyPath);
  if (sourceBytes.subarray(0, 2).toString("ascii") !== "PK") {
    throw new CliError("The input is not an unencrypted Open XML ZIP package. Encrypted or legacy PowerPoint files are not supported.");
  }
  let zip;
  try {
    zip = await JSZip.loadAsync(sourceBytes);
  } catch (error) {
    throw new CliError(`Unable to open PPTX package (encrypted or corrupt file): ${error.message}`);
  }
  const zipNames = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort(naturalPartSort);
  if (zip.file("EncryptionInfo") || zip.file("EncryptedPackage")) {
    throw new CliError("Encrypted PowerPoint files are not supported.");
  }

  const presentationXml = await safeZipText(zip, "ppt/presentation.xml");
  if (!presentationXml) throw new CliError("Invalid PPTX: ppt/presentation.xml is missing.");
  const presentationRelsXml = await safeZipText(zip, "ppt/_rels/presentation.xml.rels");
  const presentationRelationships = parseRelationships(presentationRelsXml, "ppt/presentation.xml")
    .map((relationship) => ({
      ...relationship,
      ownerPart: "ppt/presentation.xml",
      sourcePart: "ppt/_rels/presentation.xml.rels",
    }));
  const slideParts = parseSlideOrder(presentationXml, presentationRelationships, zipNames);
  const declaredSlideSize = parseSlideSize(presentationXml);

  const relevantXmlParts = [];
  const allRelationships = [];
  const xmlCandidateNames = zipNames.filter((name) => /\.(?:xml|rels)$/i.test(name));
  for (const partName of xmlCandidateNames) {
    const xml = await safeZipText(zip, partName);
    if (xml === null) continue;
    relevantXmlParts.push([partName, xml]);
    if (partName.toLowerCase().endsWith(".rels")) {
      let ownerPart = partName;
      const relMatch = partName.match(/^(.*)\/_rels\/([^/]+)\.rels$/i);
      if (relMatch) ownerPart = `${relMatch[1]}/${relMatch[2]}`;
      allRelationships.push(
        ...parseRelationships(xml, ownerPart).map((relationship) => ({
          ...relationship,
          ownerPart,
          sourcePart: partName,
        })),
      );
    }
  }
  const packageUsage = buildPackageUsageContext(zipNames, allRelationships, slideParts);

  const mediaItems = [];
  const mediaByPart = new Map();
  for (const partName of zipNames.filter((name) => name.toLowerCase().startsWith("ppt/media/"))) {
    const entryBytes = await zip.file(partName).async("nodebuffer");
    const outputName = path.posix.basename(partName);
    const outputPath = path.join(mediaDir, outputName);
    ensureInsideWorkspace(workspace, outputPath);
    await fs.writeFile(outputPath, entryBytes);
    const extension = path.extname(outputName).toLowerCase();
    const contentType = contentTypeForExtension(extension);
    const item = {
      sourcePart: partName,
      extractedPath: relativeArtifact(workspace, outputPath),
      fileName: outputName,
      extension,
      contentType,
      sizeBytes: entryBytes.length,
      sha256: sha256(entryBytes),
      rasterMetadata: await inspectRasterMetadata(sharpFactory, entryBytes, contentType, warnings, partName),
      usedOnSlides: [],
      directlyUsedOnSlides: [],
      inheritedOnSlides: [],
      usageScope: null,
    };
    mediaItems.push(item);
    mediaByPart.set(normalizeZipPath(partName), item);
  }

  const slideMediaRefs = new Map();
  for (let slideIndex = 0; slideIndex < slideParts.length; slideIndex += 1) {
    const slidePart = slideParts[slideIndex];
    const relationshipPart = `${path.posix.dirname(slidePart)}/_rels/${path.posix.basename(slidePart)}.rels`;
    const relationshipXml = await safeZipText(zip, relationshipPart);
    const relationships = parseRelationships(relationshipXml, slidePart);
    const relationshipById = new Map(relationships.map((relationship) => [relationship.id, relationship]));
    const slideXml = await safeZipText(zip, slidePart);
    const visualRelationshipIds = [];
    for (const match of String(slideXml ?? "").matchAll(/<a:blip\b([^>]*)\/?\s*>/gi)) {
      const attrs = parseAttributes(match[1]);
      const relationshipId = attrs["r:embed"] ?? attrs["r:link"];
      if (relationshipId) visualRelationshipIds.push(relationshipId);
    }
    const refsInVisualOrder = visualRelationshipIds
      .map((relationshipId) => relationshipById.get(relationshipId))
      .filter((relationship) => relationship && !relationship.external && relationship.resolvedTarget)
      .map((relationship) => normalizeZipPath(relationship.resolvedTarget));
    const refs = refsInVisualOrder.length > 0
      ? refsInVisualOrder
      : relationships
        .filter((relationship) => /\/image$/i.test(relationship.type ?? "") && relationship.resolvedTarget)
        .map((relationship) => normalizeZipPath(relationship.resolvedTarget));
    slideMediaRefs.set(slideIndex + 1, refs);
    for (const mediaRef of new Set(refs)) {
      const mediaItem = mediaByPart.get(mediaRef);
      if (mediaItem) mediaItem.directlyUsedOnSlides.push(slideIndex + 1);
    }
  }

  for (const mediaItem of mediaItems) {
    const usage = packageUsage.scopeForParts([mediaItem.sourcePart]);
    const direct = new Set(mediaItem.directlyUsedOnSlides);
    mediaItem.inheritedOnSlides = usage.activeSlides.filter((slideNumber) => !direct.has(slideNumber));
    mediaItem.usedOnSlides = [...new Set([...direct, ...usage.activeSlides])].sort((a, b) => a - b);
    mediaItem.usageScope = usage.scope;
    mediaItem.activeLayouts = usage.activeLayouts;
    mediaItem.activeMasters = usage.activeMasters;
    mediaItem.unusedLayouts = usage.unusedLayouts;
    mediaItem.unusedMasters = usage.unusedMasters;
  }

  const duplicateHashMap = new Map();
  for (const item of mediaItems) {
    const values = duplicateHashMap.get(item.sha256) ?? [];
    values.push(item.sourcePart);
    duplicateHashMap.set(item.sha256, values);
  }
  const duplicateGroups = [...duplicateHashMap.entries()]
    .filter(([, parts]) => parts.length > 1)
    .map(([hash, parts]) => ({ sha256: hash, count: parts.length, sourceParts: parts.sort(naturalPartSort) }))
    .sort((left, right) => right.count - left.count || left.sha256.localeCompare(right.sha256));

  const artifactModule = await loadRuntimeModule("@oai/artifact-tool");
  const { FileBlob, PresentationFile } = artifactModule;
  process.stderr.write(`[analyze] Importing ${sourcePath}\n`);
  let presentation;
  try {
    presentation = await PresentationFile.importPptx(await FileBlob.load(sourceCopyPath));
  } catch (error) {
    throw new CliError(`Artifact Tool could not import the PPTX: ${error.message}`, 1);
  }

  const inspection = await presentation.inspect({
    kind: "deck,slide,textbox,shape,image,table,chart,notes,thread,layout",
    maxChars: 50_000_000,
  });
  const inspectNdjson = String(inspection.ndjson ?? "");
  const inspectPath = path.join(workspace, "source-inspect.ndjson");
  await fs.writeFile(inspectPath, inspectNdjson.endsWith("\n") ? inspectNdjson : `${inspectNdjson}\n`, "utf8");
  const inspectRecords = parseInspectNdjson(inspectNdjson, warnings);
  const inspectById = new Map(inspectRecords.filter((record) => record.id).map((record) => [record.id, record]));
  const inspectSlides = new Map(inspectRecords.filter((record) => record.kind === "slide").map((record) => [record.slide, record]));
  const layoutEvidence = inspectRecords.filter((record) => record.kind === "layout");

  const layoutJsons = [];
  const slideArtifactPaths = [];
  const slideCount = presentation.slides.items.length;
  for (let slideIndex = 0; slideIndex < slideCount; slideIndex += 1) {
    const slideNumber = slideIndex + 1;
    const stem = `slide-${String(slideNumber).padStart(3, "0")}`;
    const slide = presentation.slides.items[slideIndex];
    process.stderr.write(`[analyze] Rendering slide ${slideNumber}/${slideCount}\n`);
    const renderPath = path.join(renderDir, `${stem}.png`);
    const layoutPath = path.join(layoutDir, `${stem}.layout.json`);
    try {
      await writeBlob(renderPath, await presentation.export({ slide, format: "png", scale: 1 }));
    } catch (error) {
      warnings.push(`Slide ${slideNumber} render failed: ${error.message}`);
    }
    let layoutJson = null;
    try {
      const layoutBlob = await slide.export({ format: "layout" });
      const layoutText = await layoutBlob.text();
      await fs.writeFile(layoutPath, layoutText, "utf8");
      layoutJson = JSON.parse(layoutText);
      layoutJsons.push(layoutJson);
    } catch (error) {
      warnings.push(`Slide ${slideNumber} layout export failed: ${error.message}`);
    }
    slideArtifactPaths.push({ slideNumber, renderPath, layoutPath, layoutJson });
  }

  const montagePath = path.join(workspace, "source-montage.webp");
  const montage = await createVerifiedMontage({
    sharpFactory,
    slideArtifactPaths,
    slideCount,
    workspace,
    montagePath,
    warnings,
  });

  const rawThemes = [];
  for (const partName of zipNames.filter((name) => /^ppt\/theme\/theme\d+\.xml$/i.test(name))) {
    const xml = await safeZipText(zip, partName);
    if (xml !== null) rawThemes.push(extractThemeFromXml(xml, partName));
  }
  const fonts = collectTypefaceEvidence(relevantXmlParts, layoutJsons, rawThemes);
  const visualEvidence = visualStyleEvidence(layoutJsons);
  const firstLayoutTheme = layoutJsons.find((layout) => layout.theme)?.theme ?? null;
  const themeProfile = {
    schemaVersion: "1.1",
    generatedAt: new Date().toISOString(),
    sourceSha256: sourceHashBefore,
    effectiveTheme: firstLayoutTheme,
    rawThemes,
    fonts,
    textStyleClusters: visualEvidence.textStyleClusters,
    roleCandidates: visualEvidence.roleCandidates,
    componentCandidates: visualEvidence.componentCandidates,
    visualEvidenceAdvisoryOnly: true,
    masters: layoutEvidence
      .filter((record) => record.type === "master")
      .map((record) => ({
        layoutId: record.layoutId ?? null,
        name: record.name ?? "",
        type: record.type,
        placeholders: record.placeholders ?? [],
      })),
    layouts: layoutEvidence
      .filter((record) => record.type !== "master")
      .map((record) => ({
        layoutId: record.layoutId ?? null,
        name: record.name ?? "",
        type: record.type ?? null,
        placeholders: record.placeholders ?? [],
      })),
    perSlide: layoutJsons.map((layout) => ({
      slideNumber: layout.slide?.slide ?? null,
      layoutId: layout.slide?.layoutId ?? null,
      layoutName: layout.slide?.layoutName ?? null,
      masterLayoutId: layout.slide?.masterLayoutId ?? null,
      theme: layout.theme ?? null,
    })),
  };
  const themeProfilePath = path.join(workspace, "theme-profile.json");
  await writeJson(themeProfilePath, themeProfile);

  const objects = [];
  const largeRasters = [];
  const slides = [];
  const classificationCounts = {
    "native-editable": 0,
    mixed: 0,
    flattened: 0,
    "low-quality-scan": 0,
  };

  for (const artifact of slideArtifactPaths) {
    const slideNumber = artifact.slideNumber;
    const layout = artifact.layoutJson;
    const inspectSlide = inspectSlides.get(slideNumber) ?? {};
    const frame = layout?.slide?.frame ?? declaredSlideSize;
    const slideWidth = Number(frame?.width);
    const slideHeight = Number(frame?.height);
    if (!(slideWidth > 0) || !(slideHeight > 0)) {
      throw new CliError(`Unable to derive canvas dimensions for slide ${slideNumber}; refusing to assume a default size.`);
    }
    const slideArea = slideWidth * slideHeight;
    const elements = layout ? walkElements(layout) : inspectRecords
      .filter((record) => record.slide === slideNumber && record.bbox && !["slide", "layout", "notes", "thread"].includes(record.kind))
      .map((record) => ({
        aid: record.id,
        kind: record.kind,
        name: record.name,
        bbox: record.bbox,
        text: record.text,
        textPreview: record.textPreview,
        alt: record.alt,
        scope: "slide",
      }));

    const relationshipRefs = [...(slideMediaRefs.get(slideNumber) ?? [])];
    let relationshipCursor = 0;
    const runtimeAssetMap = new Map();
    const slideObjects = [];
    for (const element of elements) {
      const inspectRecord = inspectById.get(element.aid) ?? {};
      const bboxArray = Array.isArray(element.bbox) ? element.bbox.map((value) => Number(value)) : null;
      const bboxEvidence = effectiveBboxForElement(
        bboxArray,
        element,
        inspectRecord,
        slideWidth,
        slideHeight,
      );
      const effectiveBboxArray = bboxEvidence.effective
        ? [
          bboxEvidence.effective.left,
          bboxEvidence.effective.top,
          bboxEvidence.effective.width,
          bboxEvidence.effective.height,
        ]
        : null;
      const rect = rectangleFromBbox(effectiveBboxArray, slideWidth, slideHeight);
      const clippedArea = rect ? (rect.x2 - rect.x1) * (rect.y2 - rect.y1) : 0;
      const rawAssetId = element.asset?.assetId ?? element.fillImage?.assetId ?? null;
      let mediaRef = rawAssetId && typeof rawAssetId === "string" ? normalizeZipPath(rawAssetId) : null;
      if (rawAssetId && runtimeAssetMap.has(rawAssetId)) mediaRef = runtimeAssetMap.get(rawAssetId);
      const declaredContentType = element.contentType ?? element.asset?.contentType ?? element.fillImage?.contentType ?? null;
      const hasImageAsset = element.kind === "image" || Boolean(element.fillImage) || /^image\//i.test(String(declaredContentType ?? ""));
      if (hasImageAsset && (!mediaRef || !mediaByPart.has(mediaRef))) {
        while (relationshipCursor < relationshipRefs.length && !mediaByPart.has(relationshipRefs[relationshipCursor])) {
          relationshipCursor += 1;
        }
        if (relationshipCursor < relationshipRefs.length) {
          mediaRef = relationshipRefs[relationshipCursor];
          relationshipCursor += 1;
          if (rawAssetId) runtimeAssetMap.set(rawAssetId, mediaRef);
        }
      }
      const mediaItem = mediaRef ? mediaByPart.get(mediaRef) : null;
      const contentType = mediaItem?.contentType ?? declaredContentType;
      const vectorImage = /(?:svg\+xml|x-emf|x-wmf|postscript)/i.test(String(contentType ?? ""));
      const hasRasterAsset = hasImageAsset && !vectorImage;
      const object = {
        objectId: element.aid ?? inspectRecord.id ?? `${inspectSlide.id ?? `slide-${slideNumber}`}/object-${slideObjects.length + 1}`,
        slideNumber,
        kind: inspectRecord.kind ?? element.kind ?? "unknown",
        name: element.name ?? inspectRecord.name ?? "",
        bbox: bboxEvidence.effective,
        rawBbox: bboxEvidence.adjustedForStroke ? bboxEvidence.raw : null,
        bboxAdjustment: bboxEvidence.adjustedForStroke ? {
          type: "line-stroke-effective-bbox",
          reason: "A zero-width or zero-height line/connector bbox was expanded to its visible stroke footprint.",
          strokeWidthPx: bboxEvidence.strokeWidthPx,
          minimumVisibleStrokePx: MIN_VISIBLE_STROKE_PX,
        } : null,
        geometry: element.geometry ?? inspectRecord.geometry ?? null,
        strokeWidthPx: bboxEvidence.isLineLike ? bboxEvidence.strokeWidthPx : null,
        areaRatio: slideArea > 0 ? roundRatio(clippedArea / slideArea) : 0,
        text: element.text ?? inspectRecord.text ?? null,
        textPreview: element.textPreview ?? inspectRecord.textPreview ?? null,
        alt: element.alt ?? inspectRecord.alt ?? null,
        isPlaceholder: Boolean(inspectRecord.placeholder ?? inspectRecord.isPlaceholder),
        hasImageAsset,
        hasRasterAsset,
        mediaRef: mediaItem?.sourcePart ?? mediaRef,
        mediaSha256: mediaItem?.sha256 ?? null,
        contentType: mediaItem?.contentType ?? contentType,
      };
      slideObjects.push(object);
      objects.push(object);
    }

    const imageObjects = slideObjects.filter((object) => object.hasImageAsset && object.bbox);
    const rasterObjects = imageObjects.filter((object) => object.hasRasterAsset);
    const imageRects = imageObjects
      .map((object) => rectangleFromBbox(
        [object.bbox.left, object.bbox.top, object.bbox.width, object.bbox.height],
        slideWidth,
        slideHeight,
      ))
      .filter(Boolean);
    const imageCoverageRatio = slideArea > 0 ? roundRatio(unionArea(imageRects) / slideArea) : 0;
    const largestImage = [...imageObjects].sort((left, right) => right.areaRatio - left.areaRatio)[0] ?? null;
    const largestRaster = [...rasterObjects].sort((left, right) => right.areaRatio - left.areaRatio)[0] ?? null;
    const largestImageCoverageRatio = largestImage?.areaRatio ?? 0;
    const largestMedia = largestRaster?.mediaRef ? mediaByPart.get(normalizeZipPath(largestRaster.mediaRef)) : null;
    const rasterMetadata = largestMedia?.rasterMetadata;
    const largestRasterLowResolution = Boolean(
      largestRaster && rasterMetadata?.width && rasterMetadata?.height
      && (rasterMetadata.width < largestRaster.bbox.width * 0.8
        || rasterMetadata.height < largestRaster.bbox.height * 0.8),
    );
    const nativeObjects = slideObjects.filter((object) => !object.hasImageAsset);
    const nativeTextChars = nativeObjects.reduce((sum, object) => sum + String(object.text ?? "").length, 0);
    const classification = classifySlide({
      imageCoverageRatio,
      largestImageCoverageRatio,
      imageCount: imageObjects.length,
      nativeObjectCount: nativeObjects.length,
      nativeTextChars,
      largestRasterLowResolution,
    });
    classificationCounts[classification.classification] += 1;

    for (const object of rasterObjects.filter((candidate) => candidate.areaRatio >= LARGE_RASTER_AREA_RATIO)) {
      largeRasters.push({
        objectId: object.objectId,
        slideNumber,
        bbox: bboxPxToInches(object.bbox),
        areaRatio: object.areaRatio,
        mediaRef: object.mediaRef,
        mediaSha256: object.mediaSha256,
        classificationHint: classification.classification,
        reason: object.areaRatio >= 0.78
          ? "Raster occupies most of the slide and may contain flattened content."
          : "Raster occupies a substantial content region and requires an explicit conversion decision.",
      });
    }

    const countKind = (kind) => slideObjects.filter((object) => object.kind === kind).length;
    slides.push({
      slideNumber,
      slideId: inspectSlide.id ?? layout?.slide?.aid ?? null,
      sourcePart: slideParts[slideNumber - 1] ?? null,
      title: inspectSlide.title ?? "",
      classification: classification.classification,
      classificationSuggestion: classification.classification,
      classificationConfidence: classification.confidence,
      classificationReasons: classification.reasons,
      dimensions: {
        width: pxToInches(slideWidth),
        height: pxToInches(slideHeight),
        unit: "in",
      },
      renderDimensionsPx: { width: slideWidth, height: slideHeight, unit: "px" },
      layoutId: layout?.slide?.layoutId ?? null,
      layoutName: layout?.slide?.layoutName ?? null,
      masterLayoutId: layout?.slide?.masterLayoutId ?? null,
      imageCoverageRatio,
      largestImageCoverageRatio,
      textBoxCount: countKind("textbox"),
      shapeCount: countKind("shape"),
      imageCount: imageObjects.length,
      tableCount: countKind("table"),
      chartCount: countKind("chart"),
      nativeObjectCount: nativeObjects.length,
      nativeTextChars,
      objectIds: slideObjects.map((object) => object.objectId),
      largeRasterObjectIds: largeRasters.filter((item) => item.slideNumber === slideNumber).map((item) => item.objectId),
      nativeTextCount: nativeObjects.filter((object) => String(object.text ?? "").length > 0).length,
      objects: slideObjects.map((object) => ({
        id: object.objectId,
        kind: object.kind,
        name: object.name,
        bbox: bboxPxToInches(object.bbox),
        rawBboxPx: bboxPxEvidence(object.rawBbox),
        bboxAdjustment: object.bboxAdjustment,
        geometry: object.geometry,
        strokeWidthPx: object.strokeWidthPx,
        slideLocal: true,
        sha256: object.mediaSha256,
        coverageRatio: object.areaRatio,
        textPreview: object.textPreview,
        contentType: object.contentType,
        mediaRef: object.mediaRef,
      })),
      renderPath: relativeArtifact(workspace, artifact.renderPath),
      layoutPath: relativeArtifact(workspace, artifact.layoutPath),
    });
  }

  const specialFeatures = detectSpecialFeatures(
    zipNames,
    relevantXmlParts,
    allRelationships,
    packageUsage,
  );
  const blockingFeatures = Object.entries(specialFeatures)
    .filter(([, value]) => value.present && value.blocking)
    .map(([name]) => name);
  const manualReviewFeatures = Object.entries(specialFeatures)
    .filter(([, value]) => value.present && value.manualReview)
    .map(([name]) => name);
  if (blockingFeatures.length > 0) {
    warnings.push(`Active compatibility blockers detected: ${blockingFeatures.join(", ")}.`);
  }
  const reviewOnlyFeatures = manualReviewFeatures.filter((name) => !blockingFeatures.includes(name));
  if (reviewOnlyFeatures.length > 0) {
    warnings.push(`Manual compatibility review required for: ${reviewOnlyFeatures.join(", ")}.`);
  }

  const sourceHashAfter = await hashFile(sourcePath);
  if (sourceHashAfter !== sourceHashBefore) {
    throw new CliError("The source file changed during analysis; discard this evidence package and retry.", 1);
  }
  const sourceStatAfter = await fs.stat(sourcePath);
  const specialObjects = Object.entries(specialFeatures)
    .filter(([, value]) => value.present)
    .map(([type, value]) => ({ type, ...value }));
  const manifestObjects = objects.map((object) => {
    const { rawBbox, ...serializable } = object;
    return {
      ...serializable,
      bbox: bboxPxToInches(object.bbox),
      rawBboxPx: bboxPxEvidence(rawBbox),
    };
  });
  const iconCandidates = manifestObjects
    .filter((object) => object.hasRasterAsset && Number(object.areaRatio) > 0 && Number(object.areaRatio) <= 0.12)
    .map((object) => ({
      slideNumber: object.slideNumber,
      objectId: object.objectId,
      bbox: object.bbox,
      areaRatio: object.areaRatio,
      mediaRef: object.mediaRef,
      mediaSha256: object.mediaSha256,
      candidateClass: "generic-icon-or-authentic-small-image",
      requiresVisualClassification: true,
    }));
  const effectiveSlideSizePx = declaredSlideSize ?? slides[0]?.renderDimensionsPx ?? null;
  const effectiveSlideSizeIn = effectiveSlideSizePx ? {
    width: pxToInches(effectiveSlideSizePx.width),
    height: pxToInches(effectiveSlideSizePx.height),
    unit: "in",
  } : null;
  const manifest = {
    schemaVersion: "1.1",
    analyzerVersion: SCRIPT_VERSION,
    generatedAt: new Date().toISOString(),
    sourcePptx: relativeArtifact(workspace, sourceCopyPath),
    sourceSha256: sourceHashBefore,
    slideSize: effectiveSlideSizeIn,
    coordinateSystem: {
      canonicalUnit: "in",
      origin: "top-left",
      bboxFields: ["x", "y", "width", "height"],
      note: "All manifest bbox and dimensions fields use inches. Pixel values appear only in explicitly suffixed render-evidence fields.",
    },
    slideCount,
    themeProfile: relativeArtifact(workspace, themeProfilePath),
    specialObjects,
    requiresManualCompatibilityReview: manualReviewFeatures.length > 0,
    blockingFeatures,
    manualReviewFeatures,
    source: {
      inputPath: sourcePath,
      copiedPath: relativeArtifact(workspace, sourceCopyPath),
      sha256: sourceHashBefore,
      copySha256: copyHash,
      sourceUnchanged: true,
      sizeBytes: sourceStatAfter.size,
      mtimeUtc: sourceStatAfter.mtime.toISOString(),
      extension: ".pptx",
    },
    deck: {
      slideCount,
      slideSize: effectiveSlideSizeIn,
      renderSlideSizePx: effectiveSlideSizePx,
      masters: themeProfile.masters,
      layouts: themeProfile.layouts,
      theme: themeProfile.effectiveTheme,
      fonts,
      classificationCounts,
      specialFeatures,
      requiresManualCompatibilityReview: manualReviewFeatures.length > 0,
      blockingFeatureNames: blockingFeatures,
      manualReviewFeatureNames: manualReviewFeatures,
    },
    classificationPolicy: {
      labels: ["native-editable", "mixed", "flattened", "low-quality-scan"],
      advisoryOnly: true,
      largeRasterAreaRatio: LARGE_RASTER_AREA_RATIO,
      note: "Coverage heuristics are triage evidence only; every slide must still be inspected visually at full size.",
    },
    slides,
    objects: manifestObjects,
    largeRasters,
    visualEvidence: {
      themeProfile: relativeArtifact(workspace, themeProfilePath),
      roleCandidateCount: visualEvidence.roleCandidates.length,
      componentCandidateCount: visualEvidence.componentCandidates.length,
      iconCandidates,
      advisoryOnly: true,
    },
    media: {
      count: mediaItems.length,
      items: mediaItems,
      duplicateGroups,
    },
    artifacts: {
      inspectPath: relativeArtifact(workspace, inspectPath),
      renderDir: relativeArtifact(workspace, renderDir),
      layoutDir: relativeArtifact(workspace, layoutDir),
      mediaDir: relativeArtifact(workspace, mediaDir),
      montagePath: montage.available ? relativeArtifact(workspace, montage.path) : null,
      montageEvidencePath: montage.available ? relativeArtifact(workspace, montage.evidencePath) : null,
      montage: montage.evidence,
      themeProfilePath: relativeArtifact(workspace, themeProfilePath),
    },
    warnings,
  };
  const manifestPath = path.join(workspace, "source-manifest.json");
  await writeJson(manifestPath, manifest);

  const summary = `${JSON.stringify({
    ok: true,
    source: sourcePath,
    workspace,
    manifest: manifestPath,
    slideCount,
    classificationCounts,
    largeRasterCount: largeRasters.length,
    mediaCount: mediaItems.length,
    duplicateMediaGroups: duplicateGroups.length,
    requiresManualCompatibilityReview: manualReviewFeatures.length > 0,
    blockingFeatures,
    manualReviewFeatures,
    warningCount: warnings.length,
  }, null, 2)}\n`;
  await new Promise((resolve) => process.stdout.write(summary, resolve));
  // The bundled Windows renderer can leave a native teardown hook that exits
  // non-zero after a successful custom-canvas import. All outputs are already
  // closed here, so terminate explicitly to keep the CLI contract reliable.
  process.exit(0);
}

main().catch((error) => {
  const exitCode = Number.isInteger(error.exitCode) ? error.exitCode : 1;
  process.stderr.write(`analyze_hybrid_deck: ${error.message}\n`);
  process.exitCode = exitCode;
});
