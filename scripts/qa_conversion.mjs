#!/usr/bin/env node

import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import posixPath from "node:path/posix";
import { createRequire } from "node:module";
import { parseArgs as parseNodeArgs } from "node:util";

const LARGE_RASTER_RATIO = 0.2;
const BOUNDS_TOLERANCE_PX = 2.5;
const ROUNDTRIP_MEAN_DIFF_LIMIT = 2;
const ROUNDTRIP_PIXEL_DIFF_LIMIT = 0.02;
const ALLOWED_ACTIONS = new Set([
  "keep-native",
  "rebuild-text",
  "rebuild-shape",
  "rebuild-table",
  "rebuild-chart",
  "extract-raster",
  "regenerate-icon",
  "retain-raster",
  "manual-review",
]);
const RASTER_ACTIONS = new Set(["extract-raster", "regenerate-icon", "retain-raster"]);
const ALLOWED_LARGE_RASTER_ACTIONS = new Set(["extract-raster", "retain-raster", "manual-review"]);
const VISUAL_SCHEMA_VERSION = "1.1";
const SEMANTIC_NAME_PATTERN = /(?:^|\s)mppe\|((?:[a-z]+=[a-z0-9_-]+\|?)+)/;
const DEFAULT_ROLE_FONT_TOLERANCE_PX = 0.5;
const DEFAULT_ROLE_LINE_SPACING_TOLERANCE = 0.03;
const DEFAULT_ICON_OPTICAL_TARGET = 0.72;
const DEFAULT_ICON_OPTICAL_TOLERANCE = 0.08;
const DEFAULT_ICON_CENTROID_TOLERANCE = 0.04;

function usage() {
  return [
    "Usage:",
    "  $RUNTIME_NODE qa_conversion.mjs --source <source.pptx> --final <editable.pptx> --plan <conversion-plan.json> --workspace <dir>",
    "",
    "Writes:",
    "  <workspace>/qa-report.json",
    "  <workspace>/qa-ledger.txt",
    "  <workspace>/qa-artifacts/<final-hash-prefix>/{renders,layouts,roundtrip,diffs,...}",
    "",
    "The script imports source/final with Artifact Tool, exports and re-imports a",
    "round-trip PPTX, renders every slide, emits layout JSON and visual diffs,",
    "and validates the conversion plan against the exact final file hash.",
  ].join("\n");
}

function parseCli(argv) {
  return parseNodeArgs({
    args: argv,
    options: {
      help: { type: "boolean" },
      source: { type: "string" },
      final: { type: "string" },
      plan: { type: "string" },
      workspace: { type: "string" },
    },
    allowPositionals: false,
    strict: true,
  }).values;
}

function requireString(args, key) {
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") throw new Error(`Missing required --${key}.`);
  return value;
}

function isWithin(child, parent) {
  const relative = path.relative(path.resolve(parent), path.resolve(child));
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

function runtimeRequire() {
  const modulesPath = process.env.RUNTIME_NODE_MODULES;
  if (!modulesPath || !path.isAbsolute(modulesPath)) {
    throw new Error("RUNTIME_NODE_MODULES must be set to the absolute bundled Node modules path.");
  }
  return createRequire(path.join(modulesPath, "__make_pptx_editable_runtime__.cjs"));
}

async function fileExists(filePath) {
  return fs.stat(filePath).then((stat) => stat.isFile()).catch(() => false);
}

async function directoryExists(filePath) {
  return fs.stat(filePath).then((stat) => stat.isDirectory()).catch(() => false);
}

async function sha256File(filePath) {
  const hash = crypto.createHash("sha256");
  const file = await fs.open(filePath, "r");
  try {
    for await (const chunk of file.createReadStream()) hash.update(chunk);
  } finally {
    await file.close();
  }
  return hash.digest("hex");
}

function sha256Bytes(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function canonicalJson(value) {
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(",")}]`;
  if (value && typeof value === "object") return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${canonicalJson(value[key])}`).join(",")}}`;
  return JSON.stringify(value);
}

function canonicalJsonSha256(value) {
  return sha256Bytes(Buffer.from(canonicalJson(value), "utf8"));
}

async function atomicWrite(filePath, bytes) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const temporary = `${filePath}.tmp-${process.pid}-${Date.now()}`;
  await fs.writeFile(temporary, bytes);
  await fs.rename(temporary, filePath).catch(async (error) => {
    if (error?.code !== "EEXIST" && error?.code !== "EPERM") throw error;
    await fs.rm(filePath, { force: true });
    await fs.rename(temporary, filePath);
  });
}

async function writeBlob(filePath, blob) {
  if (!blob || typeof blob.arrayBuffer !== "function") throw new Error(`Expected a Blob for ${filePath}.`);
  await atomicWrite(filePath, Buffer.from(await blob.arrayBuffer()));
}

function slidesFromPresentation(presentation) {
  if (Array.isArray(presentation.slides?.items)) return presentation.slides.items;
  if (Number.isInteger(presentation.slides?.count) && typeof presentation.slides.getItem === "function") {
    return Array.from({ length: presentation.slides.count }, (_, index) => presentation.slides.getItem(index));
  }
  throw new Error("Could not enumerate imported presentation slides.");
}

function parseInspect(ndjson) {
  const records = [];
  const parseErrors = [];
  for (const [index, line] of String(ndjson || "").split(/\r?\n/).entries()) {
    if (!line.trim()) continue;
    try {
      records.push(JSON.parse(line));
    } catch (error) {
      parseErrors.push({ line: index + 1, message: error.message });
    }
  }
  return { records, parseErrors };
}

function bboxOf(value) {
  const raw = value?.bbox ?? value?.bounds ?? value?.box;
  if (Array.isArray(raw) && raw.length >= 4) {
    const [left, top, width, height] = raw.slice(0, 4).map(Number);
    if ([left, top, width, height].every(Number.isFinite)) return { left, top, width, height };
  }
  if (raw && typeof raw === "object") {
    const left = Number(raw.left ?? raw.x);
    const top = Number(raw.top ?? raw.y);
    const width = Number(raw.width ?? raw.w);
    const height = Number(raw.height ?? raw.h);
    if ([left, top, width, height].every(Number.isFinite)) return { left, top, width, height };
  }
  const left = Number(value?.left ?? value?.x);
  const top = Number(value?.top ?? value?.y);
  const width = Number(value?.width ?? value?.w);
  const height = Number(value?.height ?? value?.h);
  if ([left, top, width, height].every(Number.isFinite)) return { left, top, width, height };
  return undefined;
}

function area(box) {
  return Math.max(0, Number(box?.width) || 0) * Math.max(0, Number(box?.height) || 0);
}

function intersectionArea(a, b) {
  if (!a || !b) return 0;
  const left = Math.max(a.left, b.left);
  const top = Math.max(a.top, b.top);
  const right = Math.min(a.left + a.width, b.left + b.width);
  const bottom = Math.min(a.top + a.height, b.top + b.height);
  return Math.max(0, right - left) * Math.max(0, bottom - top);
}

function overlapMetrics(a, b) {
  if (!a || !b) return { overlap: 0, aCoverage: 0, bCoverage: 0, iou: 0, areaRatio: Infinity };
  const aArea = area(a);
  const bArea = area(b);
  const overlap = intersectionArea(a, b);
  const union = aArea + bArea - overlap;
  return {
    overlap,
    aCoverage: aArea > 0 ? overlap / aArea : 0,
    bCoverage: bArea > 0 ? overlap / bArea : 0,
    iou: union > 0 ? overlap / union : 0,
    areaRatio: aArea > 0 && bArea > 0 ? Math.max(aArea, bArea) / Math.min(aArea, bArea) : Infinity,
  };
}

function objectContainedInRegion(regionBox, objectBox, minimumObjectCoverage = 0.8) {
  if (!regionBox || !objectBox) return false;
  const metrics = overlapMetrics(regionBox, objectBox);
  const centerX = objectBox.left + objectBox.width / 2;
  const centerY = objectBox.top + objectBox.height / 2;
  const centerInside = centerX >= regionBox.left && centerX <= regionBox.left + regionBox.width &&
    centerY >= regionBox.top && centerY <= regionBox.top + regionBox.height;
  return centerInside && metrics.bCoverage >= minimumObjectCoverage;
}

function boxesSubstantiallyMatch(a, b, minimumCoverage = 0.8, maximumAreaRatio = 1.35) {
  if (!a || !b) return false;
  const metrics = overlapMetrics(a, b);
  return metrics.aCoverage >= minimumCoverage && metrics.bCoverage >= minimumCoverage &&
    metrics.areaRatio <= maximumAreaRatio;
}

function unitScale(unit, frame) {
  const normalized = String(unit || "px").trim().toLowerCase();
  if (["px", "pixel", "pixels"].includes(normalized)) return { x: 1, y: 1 };
  if (["in", "inch", "inches"].includes(normalized)) return { x: 96, y: 96 };
  if (["pt", "point", "points"].includes(normalized)) return { x: 96 / 72, y: 96 / 72 };
  if (normalized === "emu") return { x: 1 / 9525, y: 1 / 9525 };
  if (normalized === "cm") return { x: 96 / 2.54, y: 96 / 2.54 };
  if (normalized === "mm") return { x: 96 / 25.4, y: 96 / 25.4 };
  if (["%", "percent", "percentage"].includes(normalized)) {
    return { x: frame.width / 100, y: frame.height / 100 };
  }
  throw new Error(`Unsupported bbox unit "${unit}".`);
}

function bboxToPixels(raw, frame) {
  const box = bboxOf({ bbox: raw });
  if (!box) return undefined;
  const scale = unitScale(raw?.unit, frame);
  return {
    left: box.left * scale.x,
    top: box.top * scale.y,
    width: box.width * scale.x,
    height: box.height * scale.y,
  };
}

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFKC")
    .replace(/[•●▪‣]/g, "")
    .replace(/\s+/gu, "")
    .trim();
}

function expectedTextItems(value, fallbackConfidence) {
  const out = [];
  const visit = (item) => {
    if (typeof item === "string" || typeof item === "number") {
      if (String(item).trim()) out.push({ text: String(item), confidence: fallbackConfidence, needsReview: false });
      return;
    }
    if (Array.isArray(item)) {
      for (const child of item) visit(child);
      return;
    }
    if (item && typeof item === "object") {
      const text = item.text ?? item.value ?? item.transcript;
      if (text !== undefined) {
        out.push({
          text: String(text),
          confidence: Number.isFinite(Number(item.confidence)) ? Number(item.confidence) : fallbackConfidence,
          needsReview: item.needsReview === true,
        });
      }
    }
  };
  visit(value);
  return out;
}

function flattenPlanRegions(plan) {
  const regions = [];
  const seen = new Set();
  const addRegions = (items, defaults = {}, nesting = {}) => {
    if (!Array.isArray(items)) return;
    for (const item of items) {
      if (!item || typeof item !== "object" || seen.has(item)) continue;
      seen.add(item);
      const slideNumber = Number(item.slideNumber ?? item.slide ?? defaults.slideNumber);
      const sourceObjectIds = [
        ...(Array.isArray(item.sourceObjectIds) ? item.sourceObjectIds : []),
        ...(Array.isArray(item.sourceElementIds) ? item.sourceElementIds : []),
        ...(item.sourceObjectId ? [item.sourceObjectId] : []),
        ...(item.sourceElementId ? [item.sourceElementId] : []),
      ].map(String).filter(Boolean);
      const confidence = Number.isFinite(Number(item.confidence)) ? Number(item.confidence) : 1;
      const regionId = String(item.regionId ?? item.subregionId ?? item.id ?? `s${String(slideNumber).padStart(2, "0")}-r${regions.length + 1}`);
      const region = {
        ...item,
        regionId,
        slideNumber,
        sourceObjectIds,
        action: String(item.action ?? "").trim(),
        targetType: String(item.targetType ?? item.outputType ?? "").trim(),
        confidence,
        expectedTextItems: expectedTextItems(item.expectedText ?? item.text ?? [], confidence),
        parentRegionId: nesting.parentRegionId,
        rootRegionId: nesting.rootRegionId ?? regionId,
        semanticDepth: Number(nesting.semanticDepth ?? 0),
        isSemanticSubregion: Boolean(nesting.parentRegionId),
      };
      regions.push(region);
      const childNesting = {
        parentRegionId: regionId,
        rootRegionId: region.rootRegionId,
        semanticDepth: region.semanticDepth + 1,
      };
      addRegions(item.subregions, { slideNumber }, childNesting);
    }
  };
  addRegions(plan.regions);
  addRegions(plan.items);
  addRegions(plan.actions);
  for (const slide of Array.isArray(plan.slides) ? plan.slides : []) {
    const defaults = { slideNumber: slide.slideNumber ?? slide.slide };
    addRegions(slide.regions, defaults);
    addRegions(slide.items, defaults);
    addRegions(slide.actions, defaults);
  }
  return regions;
}

function elementsFromLayout(layout) {
  return Array.isArray(layout?.elements) ? layout.elements : [];
}

function inheritedElementsFromLayout(layout) {
  return (Array.isArray(layout?.inheritedLayers) ? layout.inheritedLayers : [])
    .flatMap((layer) => (Array.isArray(layer?.elements) ? layer.elements.map((element) => ({ ...element, inheritedScope: layer.scope })) : []));
}

function frameOf(layout) {
  const frame = layout?.slide?.frame;
  if (!frame || ![frame.width, frame.height].every((number) => Number.isFinite(Number(number)))) return undefined;
  return { left: Number(frame.left || 0), top: Number(frame.top || 0), width: Number(frame.width), height: Number(frame.height) };
}

async function renderDeck(presentation, label, artifactRoot) {
  const slides = slidesFromPresentation(presentation);
  const renderDir = path.join(artifactRoot, "renders", label);
  const layoutDir = path.join(artifactRoot, "layouts", label);
  await fs.mkdir(renderDir, { recursive: true });
  await fs.mkdir(layoutDir, { recursive: true });
  const slideArtifacts = [];
  for (let index = 0; index < slides.length; index += 1) {
    const slide = slides[index];
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    const pngPath = path.join(renderDir, `${stem}.png`);
    const layoutPath = path.join(layoutDir, `${stem}.layout.json`);
    await writeBlob(pngPath, await presentation.export({ slide, format: "png", scale: 1 }));
    const layoutBlob = await slide.export({ format: "layout" });
    const layoutText = await layoutBlob.text();
    await atomicWrite(layoutPath, Buffer.from(layoutText, "utf8"));
    slideArtifacts.push({
      slideNumber: index + 1,
      pngPath,
      layoutPath,
      layout: JSON.parse(layoutText),
    });
  }
  const inspect = await presentation.inspect({
    kind: "deck,slide,textbox,shape,image,table,chart,notes,layout",
    maxChars: 1_500_000,
  });
  const inspectPath = path.join(artifactRoot, `${label}-inspect.ndjson`);
  await atomicWrite(inspectPath, Buffer.from(inspect.ndjson || "", "utf8"));
  const parsed = parseInspect(inspect.ndjson);
  return { label, slides, slideArtifacts, inspectPath, inspect: parsed };
}

async function imageDiff(sharp, firstPath, secondPath, outputPath, masks = []) {
  const first = await sharp(firstPath).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
  const secondMeta = await sharp(secondPath).metadata();
  const dimensionsMatch = first.info.width === secondMeta.width && first.info.height === secondMeta.height;
  const second = await sharp(secondPath)
    .resize(first.info.width, first.info.height, { fit: "fill" })
    .ensureAlpha()
    .raw()
    .toBuffer({ resolveWithObject: true });
  const count = first.info.width * first.info.height;
  const diff = Buffer.alloc(count * 4);
  let total = 0;
  let changed = 0;
  let maximum = 0;
  let outsideTotal = 0;
  let outsideChanged = 0;
  let outsideCount = 0;
  let maskedCount = 0;
  for (let index = 0; index < count; index += 1) {
    const x = index % first.info.width;
    const y = Math.floor(index / first.info.width);
    const masked = masks.some((box) => x >= box.left && x <= box.left + box.width && y >= box.top && y <= box.top + box.height);
    if (masked) maskedCount += 1;
    else outsideCount += 1;
    let pixelMaximum = 0;
    for (let channel = 0; channel < 3; channel += 1) {
      const delta = Math.abs(first.data[index * 4 + channel] - second.data[index * 4 + channel]);
      total += delta;
      if (!masked) outsideTotal += delta;
      pixelMaximum = Math.max(pixelMaximum, delta);
      maximum = Math.max(maximum, delta);
      diff[index * 4 + channel] = masked ? 224 : Math.min(255, delta * 4);
    }
    diff[index * 4 + 3] = 255;
    if (pixelMaximum > 12) changed += 1;
    if (!masked && pixelMaximum > 12) outsideChanged += 1;
  }
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await sharp(diff, { raw: { width: first.info.width, height: first.info.height, channels: 4 } })
    .png({ compressionLevel: 9 })
    .toFile(outputPath);
  return {
    dimensionsMatch,
    width: first.info.width,
    height: first.info.height,
    meanAbsoluteDifference: total / (count * 3),
    changedPixelRatio: changed / count,
    outsideMaskMeanAbsoluteDifference: outsideCount > 0 ? outsideTotal / (outsideCount * 3) : 0,
    outsideMaskChangedPixelRatio: outsideCount > 0 ? outsideChanged / outsideCount : 0,
    maskedPixelRatio: maskedCount / count,
    masks,
    maximumChannelDifference: maximum,
    diffPath: outputPath,
  };
}

function decodeXml(value) {
  return String(value || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

async function loadPackage(JSZip, filePath) {
  return JSZip.loadAsync(await fs.readFile(filePath));
}

function packageFeatureCounts(zip) {
  const names = Object.keys(zip.files);
  const count = (predicate) => names.filter(predicate).length;
  return {
    slideMasters: count((name) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(name)),
    slideLayouts: count((name) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(name)),
    notesSlides: count((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(name)),
    macros: count((name) => /vbaProject\.bin$/i.test(name)),
    activeX: count((name) => /^ppt\/activeX\//i.test(name)),
    oleEmbeddings: count((name) => /^ppt\/embeddings\//i.test(name)),
    smartArtParts: count((name) => /^ppt\/diagrams\//i.test(name)),
    audioVideo: count((name) => /^ppt\/media\/.*\.(mp3|m4a|wav|wma|mp4|mov|avi|wmv|mkv)$/i.test(name)),
    vectors: count((name) => /^ppt\/media\/.*\.(emf|wmf|eps)$/i.test(name)),
    inkParts: count((name) => /(^|\/)ink\//i.test(name) || /\.inkml$/i.test(name)),
  };
}

async function scanXmlFeatures(zip) {
  const slideNames = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name));
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("string") ?? "";
  const result = { transitions: 0, animations: 0, equations: 0, customShows: 0, emptyPlaceholders: [], placeholderText: [] };
  const placeholderPattern = /^(slide number|date|footer|click to add (title|subtitle|text)|(title|subtitle|name|text|body) goes here)$/i;
  for (const name of slideNames) {
    const xml = await zip.file(name).async("string");
    const slideNumber = Number(name.match(/slide(\d+)\.xml$/i)?.[1]);
    result.transitions += (xml.match(/<p:transition\b/g) || []).length;
    result.animations += (xml.match(/<p:timing\b/g) || []).length;
    result.equations += (xml.match(/<(?:m|mml):oMath\b/g) || []).length;
    for (const shape of xml.match(/<p:sp\b[\s\S]*?<\/p:sp>/g) || []) {
      if (!/<p:ph\b/.test(shape)) continue;
      const text = [...shape.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((match) => decodeXml(match[1])).join("").trim();
      const id = shape.match(/<p:cNvPr\b[^>]*\bid="([^"]+)"/)?.[1];
      const nameAttr = decodeXml(shape.match(/<p:cNvPr\b[^>]*\bname="([^"]*)"/)?.[1]);
      const type = shape.match(/<p:ph\b[^>]*\btype="([^"]+)"/)?.[1];
      if (!text) result.emptyPlaceholders.push({ slideNumber, objectId: id, name: nameAttr, type });
      if (text && placeholderPattern.test(text)) result.placeholderText.push({ slideNumber, objectId: id, name: nameAttr, text });
    }
  }
  result.customShows = (presentationXml.match(/<p:custShow\b/g) || []).length;
  return result;
}

async function scanRelationships(zip) {
  const names = new Set(Object.keys(zip.files));
  const broken = [];
  let external = 0;
  for (const relsName of [...names].filter((name) => name.endsWith(".rels"))) {
    const xml = await zip.file(relsName).async("string");
    let sourcePart = "";
    if (relsName !== "_rels/.rels") {
      const directory = posixPath.dirname(posixPath.dirname(relsName));
      const base = posixPath.basename(relsName, ".rels");
      sourcePart = posixPath.join(directory === "." ? "" : directory, base);
    }
    const baseDir = sourcePart ? posixPath.dirname(sourcePart) : "";
    for (const match of xml.matchAll(/<Relationship\b([^>]+?)\/?>(?:<\/Relationship>)?/g)) {
      const attrs = match[1];
      const target = decodeXml(attrs.match(/\bTarget="([^"]+)"/)?.[1]);
      const id = attrs.match(/\bId="([^"]+)"/)?.[1];
      const targetMode = attrs.match(/\bTargetMode="([^"]+)"/)?.[1];
      if (!target) continue;
      if (targetMode === "External" || /^[a-z][a-z0-9+.-]*:/i.test(target)) {
        external += 1;
        continue;
      }
      const withoutFragment = target.split("#")[0];
      const resolved = withoutFragment.startsWith("/")
        ? withoutFragment.slice(1)
        : posixPath.normalize(posixPath.join(baseDir, withoutFragment));
      if (resolved && !names.has(resolved)) broken.push({ relsPart: relsName, relationshipId: id, target, resolved });
    }
  }
  return { broken, external };
}

function assetIdOf(element) {
  const asset = element?.asset ?? element?.fillImage;
  if (asset && typeof asset === "object" && typeof asset.assetId === "string") return asset.assetId.replace(/^\//, "");
  if (typeof asset === "string") return asset.match(/assetId=([^;\s}]+)/)?.[1]?.replace(/^\//, "");
  return undefined;
}

async function rasterMetadata(sharp, zip, layoutArtifacts) {
  const cache = new Map();
  const results = [];
  for (const artifact of layoutArtifacts) {
    const frame = frameOf(artifact.layout);
    for (const element of elementsFromLayout(artifact.layout).filter((item) => String(item.kind).toLowerCase() === "image")) {
      const assetId = assetIdOf(element);
      if (!assetId || !zip.file(assetId)) continue;
      if (!cache.has(assetId)) {
        const bytes = await zip.file(assetId).async("nodebuffer");
        const metadata = await sharp(bytes, { failOn: "none" }).metadata().catch(() => ({}));
        let alpha;
        if (metadata.width && metadata.height && metadata.hasAlpha && metadata.width * metadata.height <= 16_000_000) {
          const raw = await sharp(bytes).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
          let transparent = 0;
          let edgeVisible = 0;
          let visible = 0;
          let visibleX = 0;
          let visibleY = 0;
          let borderOpaque = 0;
          let borderPixels = 0;
          let visibleLeft = raw.info.width;
          let visibleTop = raw.info.height;
          let visibleRight = -1;
          let visibleBottom = -1;
          for (let y = 0; y < raw.info.height; y += 1) {
            for (let x = 0; x < raw.info.width; x += 1) {
              const value = raw.data[(y * raw.info.width + x) * raw.info.channels + 3];
              if (value <= 8) transparent += 1;
              else {
                visible += 1;
                visibleX += x;
                visibleY += y;
                visibleLeft = Math.min(visibleLeft, x);
                visibleTop = Math.min(visibleTop, y);
                visibleRight = Math.max(visibleRight, x);
                visibleBottom = Math.max(visibleBottom, y);
              }
              if (x === 0 || y === 0 || x === raw.info.width - 1 || y === raw.info.height - 1) {
                borderPixels += 1;
                if (value > 8) edgeVisible += 1;
                if (value >= 250) borderOpaque += 1;
              }
            }
          }
          const transparentPixelRatio = transparent / (raw.info.width * raw.info.height);
          const borderOpaqueRatio = borderOpaque / Math.max(1, borderPixels);
          alpha = {
            transparentPixelRatio,
            edgeVisiblePixelRatio: edgeVisible / Math.max(1, raw.info.width * 2 + raw.info.height * 2 - 4),
            borderOpaqueRatio,
            alphaVisibleCentroid: visible > 0 ? {
              normalizedX: (visibleX / visible) / Math.max(1, raw.info.width - 1),
              normalizedY: (visibleY / visible) / Math.max(1, raw.info.height - 1),
            } : null,
            visibleBounds: visible > 0 ? {
              left: visibleLeft,
              top: visibleTop,
              width: visibleRight - visibleLeft + 1,
              height: visibleBottom - visibleTop + 1,
            } : null,
            opticalCoverage: visible > 0
              ? Math.max(visibleRight - visibleLeft + 1, visibleBottom - visibleTop + 1) / Math.max(raw.info.width, raw.info.height)
              : 0,
            opaqueTileScore: Math.min(1, (transparentPixelRatio < 0.001 ? 0.65 : 0) + (borderOpaqueRatio >= 0.98 ? 0.35 : 0)),
          };
        }
        cache.set(assetId, { assetId, bytes: bytes.length, sha256: sha256Bytes(bytes), ...metadata, alpha });
      }
      const box = bboxOf(element);
      const media = cache.get(assetId);
      const effectiveScale = box && media.width && media.height
        ? Math.min(media.width / Math.max(1, box.width), media.height / Math.max(1, box.height))
        : undefined;
      results.push({
        slideNumber: artifact.slideNumber,
        objectId: element.aid ?? element.id,
        name: element.name,
        bbox: box,
        areaRatio: box && frame ? area(box) / area(frame) : undefined,
        assetId,
        effectiveScale,
        ...media,
      });
    }
  }
  return results;
}

function linesOf(element) {
  const lines = element?.textLayout?.lines;
  if (!Array.isArray(lines)) return [];
  return lines.map((line) => typeof line === "string" ? line : String(line?.text ?? ""));
}

function titleLike(element) {
  return /(^|\b)(title|headline)(\b|$)|标题/u.test(String(element?.name ?? element?.placeholder?.type ?? ""));
}

function parseSemanticName(value) {
  const match = String(value ?? "").match(SEMANTIC_NAME_PATTERN);
  if (!match) return undefined;
  const metadata = {};
  for (const token of match[1].replace(/\|$/, "").split("|")) {
    const [key, item] = token.split("=");
    if (key && item) metadata[key] = item;
  }
  return Object.keys(metadata).length ? metadata : undefined;
}

function primaryTextStyle(element) {
  const paragraph = Array.isArray(element?.paragraphs) ? element.paragraphs.find((item) => String(item?.text ?? "").trim()) : undefined;
  const run = Array.isArray(paragraph?.runs) ? paragraph.runs.find((item) => String(item?.text ?? "").trim()) : undefined;
  const candidates = [run, paragraph?.resolvedTextStyle, element?.resolvedTextStyle, element];
  const pick = (key) => candidates.find((candidate) => candidate?.[key] !== undefined)?.[key];
  return {
    typeface: pick("typeface"),
    fontSize: Number(pick("fontSize") ?? element?.resolvedFontSize),
    color: String(pick("color") ?? "").toLowerCase(),
    bold: pick("bold"),
    alignment: pick("alignment"),
    lineSpacing: Number(paragraph?.lineSpacing ?? pick("lineSpacing")),
    insets: element?.resolvedTextStyle?.insets,
  };
}

function profileNumber(profile, ...keys) {
  for (const key of keys) {
    const value = Number(profile?.[key]);
    if (Number.isFinite(value)) return value;
  }
  return undefined;
}

function nearlyEqual(actual, expected, tolerance) {
  return Number.isFinite(actual) && Number.isFinite(expected) && Math.abs(actual - expected) <= tolerance;
}

function semanticElementRecords(slideArtifacts) {
  return slideArtifacts.flatMap((artifact) => elementsFromLayout(artifact.layout).map((element) => ({
    slideNumber: artifact.slideNumber,
    frame: frameOf(artifact.layout),
    element,
    bbox: bboxOf(element),
    semantic: parseSemanticName(element?.name),
  })));
}

function recordId(record) {
  return String(record?.id ?? record?.aid ?? record?.objectId ?? "");
}

function recordText(record) {
  return String(record?.text ?? record?.textPreview ?? "");
}

function kindMatches(record, desired) {
  const kind = String(record?.kind ?? "").toLowerCase();
  if (desired === "text") return (kind === "textbox" || kind === "shape") && recordText(record).trim() !== "";
  if (desired === "shape") return ["shape", "group", "connector", "line"].includes(kind);
  if (desired === "image") return kind === "image";
  if (desired === "object") return !["", "deck", "layout", "master", "slide", "notes", "thread"].includes(kind) && Boolean(bboxOf(record));
  return kind === desired;
}

function targetComponents(value) {
  return String(value ?? "")
    .toLowerCase()
    .split(/[+/,|]/u)
    .map((item) => item.trim().replace(/^(native|preserved)-/, ""))
    .map((item) => ({ shapes: "shape", connectors: "connector", images: "image", icons: "icon", tables: "table", charts: "chart", objects: "object" })[item] ?? item)
    .filter(Boolean);
}

function targetContractValid(region, components) {
  const values = new Set(components);
  if (values.size === 0) return false;
  if (region.action === "rebuild-text") return values.has("text");
  if (region.action === "rebuild-shape") return ["shape", "group", "connector"].some((item) => values.has(item));
  if (region.action === "rebuild-table") return values.has("table");
  if (region.action === "rebuild-chart") return values.has("chart");
  if (region.action === "retain-raster") return values.has("image");
  if (["extract-raster", "regenerate-icon"].includes(region.action)) {
    return [...values].every((item) => ["image", "icon"].includes(item)) && [...values].some((item) => ["image", "icon"].includes(item));
  }
  return true;
}

function finalTargetKinds(region) {
  const kinds = new Set();
  if (region.action === "rebuild-text") kinds.add("text");
  if (region.action === "rebuild-shape") kinds.add("shape");
  if (region.action === "rebuild-table") kinds.add("table");
  if (region.action === "rebuild-chart") kinds.add("chart");
  if (RASTER_ACTIONS.has(region.action)) kinds.add("image");
  for (const component of targetComponents(region.targetType)) {
    if (component === "text") kinds.add("text");
    else if (["shape", "group", "connector"].includes(component)) kinds.add("shape");
    else if (["image", "icon"].includes(component)) kinds.add("image");
    else if (["table", "chart", "object"].includes(component)) kinds.add(component);
  }
  return [...kinds];
}

function recordMatchesRegionEvidence(record, regionBox, region, desiredKind) {
  const box = bboxOf(record);
  if (!box || !regionBox) return false;
  if (RASTER_ACTIONS.has(region.action)) return boxesSubstantiallyMatch(regionBox, box);
  if (["table", "chart"].includes(desiredKind)) {
    const metrics = overlapMetrics(regionBox, box);
    return objectContainedInRegion(regionBox, box, 0.8) && metrics.aCoverage >= 0.2;
  }
  return objectContainedInRegion(regionBox, box, 0.8);
}

function nativeTextRecord(record) {
  const kind = String(record?.kind ?? "").toLowerCase();
  return ["textbox", "shape", "table", "chart"].includes(kind) && recordText(record).trim() !== "";
}

function sameStringSet(left, right) {
  const a = [...new Set((left ?? []).map(String).filter(Boolean))].sort();
  const b = [...new Set((right ?? []).map(String).filter(Boolean))].sort();
  return JSON.stringify(a) === JSON.stringify(b);
}

function normalizePackagePart(value) {
  return String(value ?? "").trim().replace(/\\/g, "/").replace(/^\/+/, "");
}

function relativeTo(workspace, filePath) {
  return path.relative(workspace, filePath).split(path.sep).join("/");
}

function ledgerText(report) {
  const lines = [
    `PPTX editability QA: ${report.status.toUpperCase()}`,
    `Generated: ${report.generatedAt}`,
    `Source: ${report.inputs.source}`,
    `Source SHA-256: ${report.hashes.sourceBefore}`,
    `Final: ${report.inputs.final}`,
    `Final SHA-256: ${report.hashes.finalBefore}`,
    `Round-trip SHA-256: ${report.hashes.roundtrip}`,
    "",
    "Summary",
    `- Slides: source ${report.summary.sourceSlides}; final ${report.summary.finalSlides}; round-trip ${report.summary.roundtripSlides}`,
    `- Plan regions: ${report.summary.planRegions}`,
    `- Semantic subregions: ${report.summary.semanticSubregions}`,
    `- Errors: ${report.summary.errors}`,
    `- Warnings: ${report.summary.warnings}`,
    `- Final large slide-local rasters: ${report.summary.finalLargeRasters}`,
    `- Manual-review regions: ${report.exceptions.manualReview.length}`,
    "",
    "Editability coverage",
    `- text-editable: ${report.editability.textEditable}`,
    `- structure-editable: ${report.editability.structureEditable}`,
    `- data-editable: ${report.editability.dataEditable}`,
    `- raster-replaceable: ${report.editability.rasterReplaceable}`,
    "",
    "Exceptions",
  ];
  const exceptionGroups = [
    ["Retained raster", report.exceptions.retainedRaster],
    ["Regenerated icons", report.exceptions.regeneratedIcons],
    ["Manual review", report.exceptions.manualReview],
  ];
  for (const [label, values] of exceptionGroups) {
    lines.push(`${label}:`);
    if (values.length === 0) lines.push("- none");
    for (const item of values) lines.push(`- slide ${item.slideNumber}, ${item.regionId}: ${item.reason || item.action}`);
  }
  lines.push("", "Failed checks and advisories");
  const findings = report.checks.filter((check) => check.status !== "pass");
  if (findings.length === 0) lines.push("- none");
  for (const check of findings) {
    const location = [check.slideNumber ? `slide ${check.slideNumber}` : "", check.regionId || ""].filter(Boolean).join(", ");
    lines.push(`- ${check.severity.toUpperCase()} ${check.id}${location ? ` (${location})` : ""}: ${check.message}`);
  }
  lines.push(
    "",
    "External presentation checks still required",
    "- Run the active $presentations template-fidelity check.",
    "- Run the active $presentations slide overflow test and corroborate inherited findings with layout JSON.",
    "- Inspect every final and round-trip slide individually at full size; montages alone are insufficient.",
    "",
  );
  return lines.join("\n");
}

async function main() {
  const args = parseCli(process.argv.slice(2));
  if (args.help) {
    console.log(usage());
    return;
  }

  const source = path.resolve(requireString(args, "source"));
  const finalPptx = path.resolve(requireString(args, "final"));
  const planPath = path.resolve(requireString(args, "plan"));
  const workspace = path.resolve(requireString(args, "workspace"));
  if (source === finalPptx) throw new Error("--final must differ from --source; QA never accepts an in-place overwrite.");
  for (const [label, filePath, extension] of [["source", source, ".pptx"], ["final", finalPptx, ".pptx"], ["plan", planPath, ".json"]]) {
    if (path.extname(filePath).toLowerCase() !== extension) throw new Error(`--${label} must use a ${extension} extension.`);
    if (!(await fileExists(filePath))) throw new Error(`Missing --${label} file: ${filePath}`);
  }
  await fs.mkdir(workspace, { recursive: true });
  if (!(await directoryExists(workspace))) throw new Error(`Could not create QA workspace: ${workspace}`);

  const sourceHashBefore = await sha256File(source);
  const finalHashBefore = await sha256File(finalPptx);
  const artifactRoot = path.join(workspace, "qa-artifacts", finalHashBefore.slice(0, 16));
  if (!isWithin(artifactRoot, path.join(workspace, "qa-artifacts")) || artifactRoot === workspace) {
    throw new Error(`Unsafe QA artifact path: ${artifactRoot}`);
  }
  await fs.rm(artifactRoot, { recursive: true, force: true });
  await fs.mkdir(artifactRoot, { recursive: true });

  const plan = JSON.parse(await fs.readFile(planPath, "utf8"));
  const regions = flattenPlanRegions(plan);
  const manifestCandidates = [
    typeof plan.sourceManifest === "string" && plan.sourceManifest.trim()
      ? (path.isAbsolute(plan.sourceManifest) ? path.resolve(plan.sourceManifest) : path.resolve(path.dirname(planPath), plan.sourceManifest))
      : undefined,
    path.join(workspace, "source-manifest.json"),
    path.join(path.dirname(planPath), "source-manifest.json"),
  ].filter(Boolean);
  let sourceManifestPath;
  for (const candidate of [...new Set(manifestCandidates)]) {
    if (await fileExists(candidate)) {
      sourceManifestPath = candidate;
      break;
    }
  }
  const sourceManifest = sourceManifestPath
    ? JSON.parse(await fs.readFile(sourceManifestPath, "utf8"))
    : undefined;

  const requireFromRuntime = runtimeRequire();
  const artifactTool = requireFromRuntime("@oai/artifact-tool");
  const sharpModule = requireFromRuntime("sharp");
  const JSZipModule = requireFromRuntime("jszip");
  const sharp = sharpModule.default ?? sharpModule;
  const JSZip = JSZipModule.default ?? JSZipModule;
  const { FileBlob, PresentationFile } = artifactTool;

  const sourceDeck = await PresentationFile.importPptx(await FileBlob.load(source));
  const finalDeck = await PresentationFile.importPptx(await FileBlob.load(finalPptx));
  const sourceEvidence = await renderDeck(sourceDeck, "source", artifactRoot);
  const finalEvidence = await renderDeck(finalDeck, "final", artifactRoot);

  const roundtripDir = path.join(artifactRoot, "roundtrip");
  await fs.mkdir(roundtripDir, { recursive: true });
  const roundtripPptx = path.join(roundtripDir, "final-roundtrip.pptx");
  await (await PresentationFile.exportPptx(finalDeck)).save(roundtripPptx);
  const roundtripDeck = await PresentationFile.importPptx(await FileBlob.load(roundtripPptx));
  const roundtripEvidence = await renderDeck(roundtripDeck, "roundtrip", artifactRoot);
  const roundtripHash = await sha256File(roundtripPptx);

  const sourceZip = await loadPackage(JSZip, source);
  const finalZip = await loadPackage(JSZip, finalPptx);
  const roundtripZip = await loadPackage(JSZip, roundtripPptx);
  const sourcePackageFeatures = packageFeatureCounts(sourceZip);
  const finalPackageFeatures = packageFeatureCounts(finalZip);
  const roundtripPackageFeatures = packageFeatureCounts(roundtripZip);
  const sourceXmlFeatures = await scanXmlFeatures(sourceZip);
  const finalXmlFeatures = await scanXmlFeatures(finalZip);
  const roundtripXmlFeatures = await scanXmlFeatures(roundtripZip);
  const finalRelationships = await scanRelationships(finalZip);
  const roundtripRelationships = await scanRelationships(roundtripZip);

  const sourceVsFinalDiffs = [];
  const finalVsRoundtripDiffs = [];
  const diffDir = path.join(artifactRoot, "diffs");
  const comparableSlides = Math.min(sourceEvidence.slideArtifacts.length, finalEvidence.slideArtifacts.length);
  for (let index = 0; index < comparableSlides; index += 1) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    const sourceFrame = frameOf(sourceEvidence.slideArtifacts[index].layout);
    const redesignMasks = sourceFrame
      ? regions.filter((region) => region.slideNumber === index + 1 && region.intent === "user-approved-redesign")
        .map((region) => {
          try { return bboxToPixels(region.bbox, sourceFrame); } catch { return undefined; }
        }).filter(Boolean)
      : [];
    sourceVsFinalDiffs.push({
      slideNumber: index + 1,
      ...(await imageDiff(
        sharp,
        sourceEvidence.slideArtifacts[index].pngPath,
        finalEvidence.slideArtifacts[index].pngPath,
        path.join(diffDir, "source-vs-final", `${stem}.png`),
        redesignMasks,
      )),
    });
  }
  const roundtripComparable = Math.min(finalEvidence.slideArtifacts.length, roundtripEvidence.slideArtifacts.length);
  for (let index = 0; index < roundtripComparable; index += 1) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    finalVsRoundtripDiffs.push({
      slideNumber: index + 1,
      ...(await imageDiff(
        sharp,
        finalEvidence.slideArtifacts[index].pngPath,
        roundtripEvidence.slideArtifacts[index].pngPath,
        path.join(diffDir, "final-vs-roundtrip", `${stem}.png`),
      )),
    });
  }

  const finalRasterMetadata = await rasterMetadata(sharp, finalZip, finalEvidence.slideArtifacts);
  const checks = [];
  const addCheck = (condition, severity, id, message, context = {}) => {
    checks.push({ id, severity, message, ...context, status: condition ? "pass" : "fail" });
    return condition;
  };

  addCheck(Boolean(sourceManifestPath), "error", "source-manifest-required", "QA requires source-manifest.json so every source large raster can be accounted for. Set plan.sourceManifest or place the manifest beside the plan or in the QA workspace.", { searched: [...new Set(manifestCandidates)] });
  addCheck(sourceHashBefore !== finalHashBefore, "error", "source-final-distinct", "Source and final SHA-256 values must differ.");
  const planSourceHash = typeof plan.sourceSha256 === "string" ? plan.sourceSha256.toLowerCase() : "";
  const manifestSourceHash = typeof sourceManifest?.sourceSha256 === "string" ? sourceManifest.sourceSha256.toLowerCase() : "";
  addCheck(/^[0-9a-f]{64}$/i.test(planSourceHash) && planSourceHash === sourceHashBefore, "error", "plan-source-hash", "Plan sourceSha256 is mandatory and must match the source file used for QA.", { declared: plan.sourceSha256, actual: sourceHashBefore });
  addCheck(/^[0-9a-f]{64}$/i.test(manifestSourceHash) && manifestSourceHash === sourceHashBefore, "error", "manifest-source-hash", "source-manifest.json must include sourceSha256 for the exact source file used for QA.", { manifest: sourceManifestPath, declared: sourceManifest?.sourceSha256, actual: sourceHashBefore });
  const sourceOriginalPath = path.join(workspace, "source-original.pptx");
  const sourceCopyMatches = await fileExists(sourceOriginalPath) ? await sha256File(sourceOriginalPath) === sourceHashBefore : false;
  addCheck(sourceCopyMatches, "warning", "source-copy", "QA workspace should contain a byte-identical source-original.pptx copy.");
  addCheck(/_editable\.pptx$/i.test(finalPptx), "warning", "editable-filename", "Final filename should use the conventional _editable.pptx suffix.");

  addCheck(sourceEvidence.inspect.parseErrors.length === 0, "error", "source-inspect-parse", "Source inspect NDJSON must parse without errors.");
  addCheck(finalEvidence.inspect.parseErrors.length === 0, "error", "final-inspect-parse", "Final inspect NDJSON must parse without errors.");
  addCheck(roundtripEvidence.inspect.parseErrors.length === 0, "error", "roundtrip-inspect-parse", "Round-trip inspect NDJSON must parse without errors.");
  addCheck(sourceEvidence.slides.length === finalEvidence.slides.length, "error", "slide-count-source-final", "Final slide count must equal the source slide count.");
  addCheck(finalEvidence.slides.length === roundtripEvidence.slides.length, "error", "slide-count-roundtrip", "Round-trip slide count must equal the final slide count.");

  const slideCount = Math.min(sourceEvidence.slideArtifacts.length, finalEvidence.slideArtifacts.length);
  for (let index = 0; index < slideCount; index += 1) {
    const sourceArtifact = sourceEvidence.slideArtifacts[index];
    const finalArtifact = finalEvidence.slideArtifacts[index];
    const sourceFrame = frameOf(sourceArtifact.layout);
    const finalFrame = frameOf(finalArtifact.layout);
    addCheck(
      Boolean(sourceFrame && finalFrame && Math.abs(sourceFrame.width - finalFrame.width) < 0.01 && Math.abs(sourceFrame.height - finalFrame.height) < 0.01),
      "error",
      "canvas-preserved",
      "Canvas dimensions and orientation must match the source.",
      { slideNumber: index + 1, sourceFrame, finalFrame },
    );
    addCheck(
      sourceArtifact.layout?.slide?.layoutId === finalArtifact.layout?.slide?.layoutId &&
        sourceArtifact.layout?.slide?.masterLayoutId === finalArtifact.layout?.slide?.masterLayoutId,
      "error",
      "layout-hierarchy-preserved",
      "Slide layout and master linkage must be preserved.",
      { slideNumber: index + 1 },
    );
  }

  for (const key of ["slideMasters", "slideLayouts", "notesSlides", "macros", "activeX", "oleEmbeddings", "smartArtParts", "audioVideo", "vectors", "inkParts"]) {
    addCheck(
      finalPackageFeatures[key] >= sourcePackageFeatures[key],
      "error",
      `package-feature-${key}`,
      `Final package must not lose source ${key} parts.`,
      { sourceCount: sourcePackageFeatures[key], finalCount: finalPackageFeatures[key] },
    );
    addCheck(
      roundtripPackageFeatures[key] >= finalPackageFeatures[key],
      "error",
      `roundtrip-feature-${key}`,
      `Round-trip package must not lose final ${key} parts.`,
      { finalCount: finalPackageFeatures[key], roundtripCount: roundtripPackageFeatures[key] },
    );
  }
  for (const key of ["transitions", "animations", "equations", "customShows"]) {
    addCheck(
      finalXmlFeatures[key] >= sourceXmlFeatures[key],
      "error",
      `xml-feature-${key}`,
      `Final package must not lose source ${key}.`,
      { sourceCount: sourceXmlFeatures[key], finalCount: finalXmlFeatures[key] },
    );
  }
  addCheck(finalRelationships.broken.length === 0, "error", "final-relationships", "Final PPTX must have no broken internal relationships.", { broken: finalRelationships.broken });
  addCheck(roundtripRelationships.broken.length === 0, "error", "roundtrip-relationships", "Round-trip PPTX must have no broken internal relationships.", { broken: roundtripRelationships.broken });
  addCheck(finalXmlFeatures.emptyPlaceholders.length === 0, "error", "empty-placeholders", "Final PPTX must not contain empty structural placeholders.", { placeholders: finalXmlFeatures.emptyPlaceholders });
  addCheck(finalXmlFeatures.placeholderText.length === 0, "error", "placeholder-prompts", "Final PPTX must not contain unresolved placeholder prompt text.", { placeholders: finalXmlFeatures.placeholderText });

  const sourceRecords = sourceEvidence.inspect.records;
  const finalRecords = finalEvidence.inspect.records;
  const roundtripRecords = roundtripEvidence.inspect.records;
  const sourceIds = new Set([
    ...sourceRecords.flatMap((record) => [record.id, record.aid, record.objectId]),
    ...(Array.isArray(sourceManifest?.objects) ? sourceManifest.objects.map((object) => object.objectId ?? object.id) : []),
  ].filter(Boolean).map(String));
  const sourceRecordById = new Map();
  for (const record of sourceRecords) {
    for (const id of [record?.id, record?.aid, record?.objectId].filter(Boolean).map(String)) {
      if (!sourceRecordById.has(id)) sourceRecordById.set(id, record);
    }
  }
  const finalRecordById = new Map();
  for (const record of finalRecords) {
    for (const id of [record?.id, record?.aid, record?.objectId].filter(Boolean).map(String)) {
      if (!finalRecordById.has(id)) finalRecordById.set(id, record);
    }
  }
  const sourceObjectById = new Map();
  for (const object of Array.isArray(sourceManifest?.objects) ? sourceManifest.objects : []) {
    const id = String(object?.objectId ?? object?.id ?? "");
    if (id) sourceObjectById.set(id, object);
  }
  for (const slide of Array.isArray(sourceManifest?.slides) ? sourceManifest.slides : []) {
    for (const object of Array.isArray(slide?.objects) ? slide.objects : []) {
      const id = String(object?.objectId ?? object?.id ?? "");
      if (id && !sourceObjectById.has(id)) sourceObjectById.set(id, object);
    }
  }

  const sourceObjectEvidence = (objectId) => sourceObjectById.get(objectId) ?? sourceRecordById.get(objectId);
  const sourceObjectHashes = (objectId) => {
    const object = sourceObjectEvidence(objectId);
    const hashes = [object?.mediaSha256, object?.sha256]
      .filter((value) => typeof value === "string" && /^[0-9a-f]{64}$/i.test(value))
      .map((value) => value.toLowerCase());
    const mediaRef = normalizePackagePart(object?.mediaRef);
    const manifestMedia = Array.isArray(sourceManifest?.media?.items)
      ? sourceManifest.media.items.find((item) => normalizePackagePart(item?.sourcePart ?? item?.mediaRef) === mediaRef)
      : undefined;
    if (typeof manifestMedia?.sha256 === "string" && /^[0-9a-f]{64}$/i.test(manifestMedia.sha256)) hashes.push(manifestMedia.sha256.toLowerCase());
    return [...new Set(hashes)];
  };

  const manifestLargeRasterIds = new Set((Array.isArray(sourceManifest?.largeRasters) ? sourceManifest.largeRasters : [])
    .map((item) => String(item.objectId ?? item.id ?? "")).filter(Boolean));
  for (const objectId of manifestLargeRasterIds) {
    addCheck(
      regions.some((region) => !region.isSemanticSubregion && region.sourceObjectIds.includes(objectId)),
      "error",
      "large-raster-plan-coverage",
      "Every source large raster must be covered by a top-level parent region; semantic subregions never own source deletion.",
      { sourceObjectId: objectId },
    );
  }

  const finalFrames = new Map();
  for (const artifact of finalEvidence.slideArtifacts) {
    finalFrames.set(artifact.slideNumber, frameOf(artifact.layout));
  }
  const finalRecordsBySlide = new Map();
  for (const record of finalRecords) {
    const slideNumber = Number(record.slide);
    if (!Number.isInteger(slideNumber)) continue;
    if (!finalRecordsBySlide.has(slideNumber)) finalRecordsBySlide.set(slideNumber, []);
    finalRecordsBySlide.get(slideNumber).push(record);
  }

  const visualSchema = String(plan.schemaVersion ?? "1.0") === VISUAL_SCHEMA_VERSION;
  const visualPolicy = plan.visualPolicy ?? {};
  const styleProfile = plan.styleProfile ?? {};
  const roleProfiles = styleProfile.roles && typeof styleProfile.roles === "object" ? styleProfile.roles : {};
  const componentProfiles = styleProfile.componentFamilies && typeof styleProfile.componentFamilies === "object" ? styleProfile.componentFamilies : {};
  const iconProfile = styleProfile.iconFamily && typeof styleProfile.iconFamily === "object" ? styleProfile.iconFamily : {};
  const semanticElements = semanticElementRecords(finalEvidence.slideArtifacts);
  const feedbackIssues = Array.isArray(plan.feedbackIssues) ? plan.feedbackIssues : [];
  const visualConsistency = {
    schemaVersion: String(plan.schemaVersion ?? "1.0"),
    semanticObjects: semanticElements.filter((item) => item.semantic).length,
    roleCounts: {},
    componentInstances: {},
    calibration: plan.calibration ?? null,
    feedbackIssues: feedbackIssues.map((issue) => ({
      issueId: issue.issueId,
      scope: issue.scope,
      status: issue.status,
      slideNumber: issue.slideNumber,
    })),
  };

  if (visualSchema) {
    addCheck(visualPolicy.normalizationScope === "full-deck-role-based", "error", "visual-policy-scope", "Schema 1.1 requires full-deck role-based normalization across native and rebuilt objects.", { actual: visualPolicy.normalizationScope });
    addCheck(visualPolicy.iconMode === "extract-or-regenerate-generic", "error", "visual-policy-icons", "Schema 1.1 requires clean extraction or regeneration of every generic raster icon while preserving authentic assets.", { actual: visualPolicy.iconMode });
    addCheck(visualPolicy.qaStrictness === "tiered", "error", "visual-policy-qa", "Schema 1.1 requires tiered QA: deterministic defects fail and subjective drift warns.", { actual: visualPolicy.qaStrictness });
    addCheck(["automatic", "user-gated"].includes(visualPolicy.calibrationMode), "error", "visual-policy-calibration-mode", "Calibration mode must be automatic or user-gated.", { actual: visualPolicy.calibrationMode });
    const calibration = plan.calibration ?? {};
    addCheck(calibration.mode === visualPolicy.calibrationMode, "error", "calibration-mode", "Calibration mode must match visualPolicy.calibrationMode.", { policy: visualPolicy.calibrationMode, calibration: calibration.mode });
    addCheck(["complete", "approved"].includes(calibration.status), "error", "calibration-status", "Full-deck QA requires completed automatic calibration or user-approved calibration.", { status: calibration.status });
    addCheck(Array.isArray(calibration.representativeSlides) && calibration.representativeSlides.length > 0, "error", "calibration-slides", "Calibration requires at least one representative slide.", { representativeSlides: calibration.representativeSlides });
    addCheck(Array.isArray(calibration.evidence) && calibration.evidence.length > 0, "error", "calibration-evidence", "Calibration requires render/inspection evidence.", { evidence: calibration.evidence });
    addCheck(/^[0-9a-f]{64}$/i.test(String(calibration.frozenProfileSha256 ?? "")) && String(calibration.frozenProfileSha256).toLowerCase() === canonicalJsonSha256(styleProfile), "error", "calibration-profile-hash", "Calibration must bind the exact canonical style profile with SHA-256.", { frozenProfileSha256: calibration.frozenProfileSha256, expected: canonicalJsonSha256(styleProfile) });

    for (const issue of feedbackIssues) {
      const issueContext = { issueId: issue.issueId, slideNumber: issue.slideNumber, scope: issue.scope, status: issue.status };
      addCheck(issue.status !== "open", "error", "feedback-issue-open", "Every annotated feedback issue must be fixed or explicitly waived before delivery.", issueContext);
      if (issue.status === "waived") addCheck(typeof issue.waiverReason === "string" && issue.waiverReason.trim() !== "", "error", "feedback-waiver-reason", "A waived feedback issue requires a nonblank reason.", issueContext);
      if (["component-family", "deck"].includes(issue.scope)) {
        addCheck(Array.isArray(issue.inspectedInstanceIds) && issue.inspectedInstanceIds.length > 0, "error", "feedback-propagation-evidence", "Family/deck feedback must enumerate every peer instance inspected after propagation.", issueContext);
      }
      if (issue.status === "fixed") {
        addCheck(regions.some((region) => Array.isArray(region.issueRefs) && region.issueRefs.includes(issue.issueId)), "error", "feedback-fix-binding", "A fixed feedback issue must be bound to at least one conversion region through issueRefs.", issueContext);
      }
    }

    for (const item of semanticElements) {
      const text = String(item.element?.text ?? item.element?.textPreview ?? "").trim();
      if (text) addCheck(Boolean(item.semantic?.role), "error", "semantic-text-role", "Every meaningful slide-local text object in schema 1.1 requires a semantic role name.", { slideNumber: item.slideNumber, objectId: item.element?.aid ?? item.element?.id, name: item.element?.name, textPreview: text.slice(0, 80) });
      if (!item.semantic) continue;
      const role = item.semantic.role;
      if (role) visualConsistency.roleCounts[role] = (visualConsistency.roleCounts[role] ?? 0) + 1;
      if (role && text) {
        const profile = roleProfiles[role];
        addCheck(Boolean(profile), "error", "semantic-role-profile", "Every used text role must exist in styleProfile.roles.", { slideNumber: item.slideNumber, objectId: item.element?.aid ?? item.element?.id, role });
        if (profile) {
          const actual = primaryTextStyle(item.element);
          const expectedFontSize = profileNumber(profile, "fontSizePx", "fontSize");
          if (Number.isFinite(expectedFontSize)) addCheck(nearlyEqual(actual.fontSize, expectedFontSize, Number(profile.fontSizeTolerancePx ?? DEFAULT_ROLE_FONT_TOLERANCE_PX)), "error", "role-font-size", "Text font size differs from its frozen role token.", { slideNumber: item.slideNumber, objectId: item.element?.aid ?? item.element?.id, role, expected: expectedFontSize, actual: actual.fontSize });
          if (profile.typeface) addCheck(String(actual.typeface ?? "").toLowerCase() === String(profile.typeface).toLowerCase(), "error", "role-typeface", "Text typeface differs from its frozen role token.", { slideNumber: item.slideNumber, objectId: item.element?.aid ?? item.element?.id, role, expected: profile.typeface, actual: actual.typeface });
          if (profile.color) addCheck(actual.color === String(profile.color).toLowerCase(), "error", "role-color", "Text color differs from its frozen role token.", { slideNumber: item.slideNumber, objectId: item.element?.aid ?? item.element?.id, role, expected: profile.color, actual: actual.color });
          if (profile.bold !== undefined) addCheck(Boolean(actual.bold) === Boolean(profile.bold), "error", "role-bold", "Text weight differs from its frozen role token.", { slideNumber: item.slideNumber, objectId: item.element?.aid ?? item.element?.id, role, expected: profile.bold, actual: actual.bold });
          if (profile.alignment) addCheck(String(actual.alignment ?? "").toLowerCase() === String(profile.alignment).toLowerCase(), "error", "role-alignment", "Text alignment differs from its frozen role token.", { slideNumber: item.slideNumber, objectId: item.element?.aid ?? item.element?.id, role, expected: profile.alignment, actual: actual.alignment });
          const expectedLineSpacing = profileNumber(profile, "lineSpacing");
          if (Number.isFinite(expectedLineSpacing) && Number.isFinite(actual.lineSpacing)) addCheck(nearlyEqual(actual.lineSpacing, expectedLineSpacing, Number(profile.lineSpacingTolerance ?? DEFAULT_ROLE_LINE_SPACING_TOLERANCE)), "error", "role-line-spacing", "Text line spacing differs from its frozen role token.", { slideNumber: item.slideNumber, objectId: item.element?.aid ?? item.element?.id, role, expected: expectedLineSpacing, actual: actual.lineSpacing });
          const maxLines = profileNumber(profile, "maxLines");
          if (Number.isFinite(maxLines)) addCheck(linesOf(item.element).length <= maxLines, "error", "role-max-lines", "Text exceeds the maximum line count for its semantic role.", { slideNumber: item.slideNumber, objectId: item.element?.aid ?? item.element?.id, role, maxLines, lines: linesOf(item.element) });
        }
      }
    }

    const componentInstances = new Map();
    for (const item of semanticElements.filter((entry) => entry.semantic?.family && entry.semantic?.instance && entry.semantic?.part && entry.bbox)) {
      const key = `${item.semantic.family}|${item.semantic.instance}`;
      if (!componentInstances.has(key)) componentInstances.set(key, { family: item.semantic.family, instance: item.semantic.instance, parts: new Map(), slideNumber: item.slideNumber });
      componentInstances.get(key).parts.set(item.semantic.part, item);
    }
    for (const instance of componentInstances.values()) {
      addCheck(Boolean(componentProfiles[instance.family]), "error", "component-family-profile", "Every used component family must exist in styleProfile.componentFamilies.", { family: instance.family, instance: instance.instance, slideNumber: instance.slideNumber });
      visualConsistency.componentInstances[instance.family] = (visualConsistency.componentInstances[instance.family] ?? 0) + 1;
    }
    const families = new Map();
    for (const instance of componentInstances.values()) {
      if (!families.has(instance.family)) families.set(instance.family, []);
      families.get(instance.family).push(instance);
    }
    for (const [family, instances] of families) {
      if (instances.length < 2) continue;
      const profile = componentProfiles[family] ?? {};
      const flexible = new Set(Array.isArray(profile.flexible) ? profile.flexible : []);
      const baseline = instances[0];
      const baselineFrame = baseline.parts.get("frame")?.bbox;
      if (!baselineFrame) {
        addCheck(false, "error", "component-frame", "Repeated component instances require a semantic part=frame anchor.", { family, instance: baseline.instance });
        continue;
      }
      for (const instance of instances.slice(1)) {
        const frameBox = instance.parts.get("frame")?.bbox;
        addCheck(Boolean(frameBox), "error", "component-frame", "Repeated component instances require a semantic part=frame anchor.", { family, instance: instance.instance });
        if (!frameBox) continue;
        const tolerance = Number(profile.tolerancePx ?? Math.max(2, 0.0025 * Math.max(instance.parts.get("frame")?.frame?.width ?? 0, instance.parts.get("frame")?.frame?.height ?? 0)));
        for (const [part, baselinePart] of baseline.parts) {
          const peer = instance.parts.get(part);
          addCheck(Boolean(peer), "error", "component-part-present", "Every repeated component instance must contain the same named parts.", { family, instance: instance.instance, part });
          if (!peer) continue;
          const baseLocal = { left: baselinePart.bbox.left - baselineFrame.left, top: baselinePart.bbox.top - baselineFrame.top, width: baselinePart.bbox.width, height: baselinePart.bbox.height };
          const peerLocal = { left: peer.bbox.left - frameBox.left, top: peer.bbox.top - frameBox.top, width: peer.bbox.width, height: peer.bbox.height };
          for (const axis of ["left", "top", "width", "height"]) {
            if (flexible.has(`${part}.${axis}`)) continue;
            addCheck(Math.abs(baseLocal[axis] - peerLocal[axis]) <= tolerance, "error", "component-geometry", "Repeated component parts must share the frozen relative geometry unless the axis is explicitly flexible.", { family, instance: instance.instance, part, axis, baseline: baseLocal[axis], actual: peerLocal[axis], tolerance });
          }
        }
      }
    }

    const bulletGroups = new Map();
    for (const item of semanticElements.filter((entry) => entry.semantic?.role === "bullet-body" && entry.semantic?.family && entry.semantic?.instance && entry.semantic?.item)) {
      const key = `${item.semantic.family}|${item.semantic.instance}|${item.semantic.item}`;
      if (!bulletGroups.has(key)) bulletGroups.set(key, {});
      bulletGroups.get(key)[item.semantic.part] = item;
    }
    const bulletTextLeftByList = new Map();
    for (const [key, pair] of bulletGroups) {
      addCheck(Boolean(pair.dot && pair.text), "error", "heavy-bullet-pair", "A heavy bullet item requires one named dot shape and one named text object.", { key });
      if (!pair.dot || !pair.text) continue;
      const fontSize = primaryTextStyle(pair.text.element).fontSize;
      const diameter = (pair.dot.bbox.width + pair.dot.bbox.height) / 2;
      const gap = pair.text.bbox.left - (pair.dot.bbox.left + pair.dot.bbox.width);
      addCheck(Math.abs(pair.dot.bbox.width - pair.dot.bbox.height) <= 1.5, "error", "heavy-bullet-circle", "Heavy bullet dots must render as circles.", { key, bbox: pair.dot.bbox });
      if (Number.isFinite(fontSize) && fontSize > 0) {
        addCheck(diameter / fontSize >= 0.38 && diameter / fontSize <= 0.45, "error", "heavy-bullet-size", "Heavy bullet diameter must remain proportional to the body font size.", { key, diameter, fontSize, ratio: diameter / fontSize });
        addCheck(gap / fontSize >= 0.55 && gap / fontSize <= 0.75, "error", "heavy-bullet-gap", "The gap between a heavy bullet and its text must remain proportional to the font size.", { key, gap, fontSize, ratio: gap / fontSize });
        const dotCenter = pair.dot.bbox.top + pair.dot.bbox.height / 2;
        const firstLineCenter = pair.text.bbox.top + fontSize * 0.5;
        addCheck(Math.abs(dotCenter - firstLineCenter) <= fontSize * 0.15, "error", "heavy-bullet-baseline", "Heavy bullet center must align with the first text line, not the whole text box.", { key, dotCenter, firstLineCenter, tolerance: fontSize * 0.15 });
      }
      const listKey = `${pair.text.semantic.family}|${pair.text.semantic.instance}`;
      if (!bulletTextLeftByList.has(listKey)) bulletTextLeftByList.set(listKey, pair.text.bbox.left);
      addCheck(Math.abs(bulletTextLeftByList.get(listKey) - pair.text.bbox.left) <= 1, "error", "heavy-bullet-text-indent", "Wrapped and peer bullet text must share one vertical text indent.", { key, expectedLeft: bulletTextLeftByList.get(listKey), actualLeft: pair.text.bbox.left });
    }
  }

  for (const region of regions) {
    addCheck(Number.isInteger(region.slideNumber) && region.slideNumber >= 1 && region.slideNumber <= finalEvidence.slides.length, "error", "region-slide", "Plan region must reference a valid one-based slide.", { slideNumber: region.slideNumber, regionId: region.regionId });
    addCheck(ALLOWED_ACTIONS.has(region.action), "error", "region-action", `Unsupported conversion action "${region.action}".`, { slideNumber: region.slideNumber, regionId: region.regionId });
    addCheck(typeof region.reason === "string" && region.reason.trim() !== "", "error", "region-reason", "Every conversion action requires a nonblank evidence-based reason.", { slideNumber: region.slideNumber, regionId: region.regionId, action: region.action });
    const components = targetComponents(region.targetType);
    addCheck(targetContractValid(region, components), "error", "region-target-contract", `targetType "${region.targetType}" is incompatible with action "${region.action}".`, { slideNumber: region.slideNumber, regionId: region.regionId, action: region.action, targetType: region.targetType });
    if (["rebuild-text", "rebuild-shape", "rebuild-table", "rebuild-chart", "extract-raster", "regenerate-icon"].includes(region.action)) {
      addCheck(Number.isFinite(region.confidence) && region.confidence >= 0.8, "error", "region-confidence", "Destructive reconstruction actions require confidence >= 0.80; otherwise use manual-review.", { slideNumber: region.slideNumber, regionId: region.regionId, confidence: region.confidence });
    }
    for (const sourceObjectId of region.sourceObjectIds) {
      addCheck(sourceIds.has(sourceObjectId), "error", "region-source-object", "Planned source object ID must exist in source inspect/manifest evidence.", { slideNumber: region.slideNumber, regionId: region.regionId, sourceObjectId });
      const sourceObject = sourceObjectEvidence(sourceObjectId);
      const sourceSlideNumber = Number(sourceObject?.slideNumber ?? sourceObject?.slide);
      addCheck(!Number.isInteger(sourceSlideNumber) || sourceSlideNumber === region.slideNumber, "error", "region-source-slide", "Planned source object must belong to the same slide as its conversion region.", { slideNumber: region.slideNumber, regionId: region.regionId, sourceObjectId, sourceSlideNumber });
    }
    const frame = finalFrames.get(region.slideNumber);
    let regionBox;
    try {
      regionBox = frame ? bboxToPixels(region.bbox, frame) : undefined;
    } catch (error) {
      addCheck(false, "error", "region-bbox-unit", error.message, { slideNumber: region.slideNumber, regionId: region.regionId });
    }
    addCheck(Boolean(regionBox && area(regionBox) > 0), "error", "region-bbox", "Plan region must have a finite positive bbox.", { slideNumber: region.slideNumber, regionId: region.regionId });
    if (regionBox && frame) {
      addCheck(
        regionBox.left >= -BOUNDS_TOLERANCE_PX && regionBox.top >= -BOUNDS_TOLERANCE_PX &&
          regionBox.left + regionBox.width <= frame.width + BOUNDS_TOLERANCE_PX &&
          regionBox.top + regionBox.height <= frame.height + BOUNDS_TOLERANCE_PX,
        "error",
        "region-inside-canvas",
        "Plan region must remain inside the actual source/final slide canvas.",
        { slideNumber: region.slideNumber, regionId: region.regionId, bbox: regionBox, frame },
      );
    }

    const slideRecords = finalRecordsBySlide.get(region.slideNumber) ?? [];
    const desiredKinds = finalTargetKinds(region);
    for (const desiredKind of desiredKinds) {
      const matchingTargets = slideRecords.filter((record) => kindMatches(record, desiredKind) && recordMatchesRegionEvidence(record, regionBox, region, desiredKind));
      addCheck(
        matchingTargets.length > 0,
        region.action === "manual-review" ? "warning" : "error",
        "region-output-type",
        `Region must contain the planned native/output object type: ${desiredKind}.`,
        { slideNumber: region.slideNumber, regionId: region.regionId, action: region.action, targetType: region.targetType, desiredKind, matches: matchingTargets.map(recordId) },
      );
    }

    if (region.action === "keep-native" && region.sourceObjectIds.length > 0) {
      for (const sourceObjectId of region.sourceObjectIds) {
        const sourceObject = sourceObjectEvidence(sourceObjectId);
        const finalObject = finalRecordById.get(sourceObjectId);
        addCheck(Boolean(finalObject && Number(finalObject.slide) === region.slideNumber), "error", "keep-native-present", "Every keep-native source object ID must remain present on the same final slide; a different nearby object is not sufficient evidence.", { slideNumber: region.slideNumber, regionId: region.regionId, sourceObjectId });
        if (sourceObject && finalObject && frame) {
          let sourceBox;
          try { sourceBox = bboxToPixels(sourceObject.bbox ?? sourceObject.bounds ?? sourceObject.box, frame); } catch { sourceBox = undefined; }
          addCheck(Boolean(sourceBox && boxesSubstantiallyMatch(sourceBox, bboxOf(finalObject), 0.9, 1.15)), "error", "keep-native-geometry", "A keep-native object must retain substantially the same geometry.", { slideNumber: region.slideNumber, regionId: region.regionId, sourceObjectId, sourceBBox: sourceBox, finalBBox: bboxOf(finalObject) });
          const hashes = sourceObjectHashes(sourceObjectId);
          if (hashes.length > 0 && kindMatches(finalObject, "image")) {
            const finalMedia = finalRasterMetadata.find((media) => media.slideNumber === region.slideNumber && (media.objectId === sourceObjectId || boxesSubstantiallyMatch(sourceBox, media.bbox, 0.9, 1.15)));
            addCheck(Boolean(finalMedia && hashes.includes(String(finalMedia.sha256 ?? "").toLowerCase())), "error", "keep-native-raster-hash", "A keep-native raster must preserve the exact embedded source media bytes.", { slideNumber: region.slideNumber, regionId: region.regionId, sourceObjectId, expectedHashes: hashes, actualHash: finalMedia?.sha256 });
          }
        }
      }
    }
    const nativeTextActions = new Set(["keep-native", "rebuild-text", "rebuild-shape", "rebuild-table", "rebuild-chart"]);
    if (region.action === "rebuild-text") {
      addCheck(region.expectedTextItems.length > 0, "error", "expected-text-required", "rebuild-text requires nonblank expectedText.", { slideNumber: region.slideNumber, regionId: region.regionId });
    }
    if (nativeTextActions.has(region.action) && region.expectedTextItems.length > 0) {
      const regionTextRecords = slideRecords.filter((record) => nativeTextRecord(record) && objectContainedInRegion(regionBox, bboxOf(record), 0.8));
      const orderedText = [...regionTextRecords]
        .sort((a, b) => (bboxOf(a)?.top ?? 0) - (bboxOf(b)?.top ?? 0) || (bboxOf(a)?.left ?? 0) - (bboxOf(b)?.left ?? 0))
        .map(recordText)
        .join("\n");
      for (const expected of region.expectedTextItems) {
        const expectedNormalized = normalizeText(expected.text);
        const matchingRecords = regionTextRecords.filter((record) => normalizeText(recordText(record)).includes(expectedNormalized));
        const found = expectedNormalized !== "" && (matchingRecords.length > 0 || normalizeText(orderedText).includes(expectedNormalized));
        addCheck(found, "error", "expected-text-accounted", `Expected text is missing from native final text inside the planned region: "${expected.text}".`, { slideNumber: region.slideNumber, regionId: region.regionId, regionBBox: regionBox });
        if (["rebuild-text", "rebuild-shape", "rebuild-table", "rebuild-chart"].includes(region.action)) {
          addCheck(!(expected.needsReview || Number(expected.confidence) < 0.8), "error", "uncertain-text", "Low-confidence or review-marked text must not be silently finalized as rebuilt text.", { slideNumber: region.slideNumber, regionId: region.regionId, text: expected.text, confidence: expected.confidence });
        }
        addCheck(matchingRecords.length <= 1, "warning", "duplicate-expected-text", `Expected text appears in multiple native objects: "${expected.text}".`, { slideNumber: region.slideNumber, regionId: region.regionId, matches: matchingRecords.map(recordId) });
      }
    }
    if (region.action === "retain-raster" && regionBox) {
      const expectedHashes = new Set(region.sourceObjectIds.flatMap(sourceObjectHashes));
      const retainedMedia = finalRasterMetadata.filter((media) =>
        media.slideNumber === region.slideNumber && boxesSubstantiallyMatch(regionBox, media.bbox));
      addCheck(retainedMedia.length > 0, "error", "retained-raster-present", "A retain-raster region must still contain a substantially matching raster, not merely a small image somewhere inside the region.", { slideNumber: region.slideNumber, regionId: region.regionId });
      addCheck(expectedHashes.size > 0, "error", "retained-raster-source-hash-evidence", "A retain-raster region requires a source-manifest media hash for its bound source object.", { slideNumber: region.slideNumber, regionId: region.regionId, sourceObjectIds: region.sourceObjectIds });
      if (expectedHashes.size > 0 && retainedMedia.length > 0) {
        addCheck(retainedMedia.some((media) => expectedHashes.has(String(media.sha256 ?? "").toLowerCase())), "error", "retained-raster-hash", "A retained raster must preserve the source media hash. Use extract-raster when intentional pixel changes are required.", { slideNumber: region.slideNumber, regionId: region.regionId, expectedHashes: [...expectedHashes], actualHashes: retainedMedia.map((media) => media.sha256).filter(Boolean) });
      }
    }
    if (region.action === "manual-review") {
      addCheck(false, "warning", "manual-review", region.reason || "Conversion plan contains an unresolved manual-review item.", { slideNumber: region.slideNumber, regionId: region.regionId });
    }

    if (["rebuild-text", "rebuild-shape", "rebuild-table", "rebuild-chart"].includes(region.action) && regionBox && frame) {
      const coveringRaster = slideRecords.find((record) => {
        if (!kindMatches(record, "image")) return false;
        const box = bboxOf(record);
        return box && area(box) / area(frame) >= LARGE_RASTER_RATIO && intersectionArea(box, regionBox) / Math.max(1, area(regionBox)) >= 0.7;
      });
      addCheck(!coveringRaster, "error", "rebuild-region-raster-remains", "A large raster still covers a region planned for native reconstruction.", { slideNumber: region.slideNumber, regionId: region.regionId, imageId: coveringRaster ? recordId(coveringRaster) : undefined });
    }
  }

  const finalLargeRasters = [];
  for (const [slideNumber, slideRecords] of finalRecordsBySlide) {
    const frame = finalFrames.get(slideNumber);
    if (!frame) continue;
    for (const record of slideRecords.filter((item) => kindMatches(item, "image"))) {
      const box = bboxOf(record);
      const ratio = box ? area(box) / area(frame) : 0;
      if (ratio < LARGE_RASTER_RATIO) continue;
      const allowedRegion = regions.find((region) => {
        if (region.isSemanticSubregion || region.slideNumber !== slideNumber || !ALLOWED_LARGE_RASTER_ACTIONS.has(region.action)) return false;
        try {
          return boxesSubstantiallyMatch(bboxToPixels(region.bbox, frame), box);
        } catch {
          return false;
        }
      });
      finalLargeRasters.push({ slideNumber, objectId: recordId(record), bbox: box, areaRatio: ratio, allowedByRegion: allowedRegion?.regionId });
      addCheck(Boolean(allowedRegion), "error", "unclassified-large-raster", "No unclassified large slide-local raster may remain in the final deck.", { slideNumber, objectId: recordId(record), areaRatio: ratio });
    }
  }

  for (const artifact of finalEvidence.slideArtifacts) {
    const frame = frameOf(artifact.layout);
    if (!frame) continue;
    for (const element of elementsFromLayout(artifact.layout)) {
      const box = bboxOf(element);
      if (!box) continue;
      addCheck(
        box.left >= -BOUNDS_TOLERANCE_PX && box.top >= -BOUNDS_TOLERANCE_PX &&
          box.left + box.width <= frame.width + BOUNDS_TOLERANCE_PX &&
          box.top + box.height <= frame.height + BOUNDS_TOLERANCE_PX,
        "error",
        "slide-local-bounds",
        "Slide-local content must stay inside the actual slide canvas.",
        { slideNumber: artifact.slideNumber, objectId: element.aid ?? element.id, name: element.name, bbox: box },
      );
    }
    const inheritedOutside = inheritedElementsFromLayout(artifact.layout).filter((element) => {
      const box = bboxOf(element);
      return box && (box.left < -BOUNDS_TOLERANCE_PX || box.top < -BOUNDS_TOLERANCE_PX || box.left + box.width > frame.width + BOUNDS_TOLERANCE_PX || box.top + box.height > frame.height + BOUNDS_TOLERANCE_PX);
    });
    addCheck(inheritedOutside.length === 0, "warning", "inherited-bounds", "Inherited template furniture extends outside the slide; corroborate generic overflow findings with this layout evidence.", { slideNumber: artifact.slideNumber, count: inheritedOutside.length });

    for (const element of elementsFromLayout(artifact.layout)) {
      const text = String(element?.text ?? element?.textPreview ?? "");
      addCheck(!/(^|\n)\s*[•●▪‣]\s*/u.test(text), "error", "literal-bullet", "Body text must not imitate bullets with a literal bullet glyph.", { slideNumber: artifact.slideNumber, objectId: element.aid ?? element.id, name: element.name });
      const lines = linesOf(element);
      const semantic = parseSemanticName(element?.name);
      if (titleLike(element) && !semantic?.role) {
        addCheck(lines.length <= 1, "warning", "title-wrap", "A title-like object wraps to multiple lines; confirm this is intentional.", { slideNumber: artifact.slideNumber, objectId: element.aid ?? element.id, name: element.name, lines });
      }
      if (lines.length > 1) {
        const last = normalizeText(lines.at(-1));
        addCheck(!(last.length === 1 && !/^\d$/u.test(last)), "warning", "single-character-orphan", "A multi-line text object ends with an isolated one-character line.", { slideNumber: artifact.slideNumber, objectId: element.aid ?? element.id, name: element.name, lines });
      }
      for (const paragraph of Array.isArray(element?.paragraphs) ? element.paragraphs : []) {
        if (!paragraph?.bulletCharacter) continue;
        const marginLeft = Number(paragraph.marginLeft);
        const indent = Number(paragraph.indent);
        addCheck(Number.isFinite(marginLeft) && marginLeft > 0 && Number.isFinite(indent) && indent < 0, "warning", "bullet-hanging-indent", "Structured bullets should have a positive left margin and negative hanging indent.", { slideNumber: artifact.slideNumber, objectId: element.aid ?? element.id, marginLeft: paragraph.marginLeft, indent: paragraph.indent });
      }
      const fontSize = Number(element?.resolvedFontSize ?? element?.resolvedTextStyle?.fontSize);
      if (text.trim() && Number.isFinite(fontSize) && fontSize < 10 && !/page|slide number|页码|编号/i.test(String(element?.name ?? ""))) {
        addCheck(false, "warning", "microtext", "Text smaller than 10 px was found; confirm it was not introduced merely to force a fit.", { slideNumber: artifact.slideNumber, objectId: element.aid ?? element.id, name: element.name, fontSize });
      }
    }
  }

  const finalImages = finalRecords.filter((record) => kindMatches(record, "image"));
  for (const image of finalImages) {
    addCheck(typeof image.alt === "string" && image.alt.trim() !== "", "warning", "image-alt", "Every slide-local raster should have meaningful alt text or an object description.", { slideNumber: image.slide, objectId: recordId(image), name: image.name });
  }
  for (const media of finalRasterMetadata) {
    addCheck(media.effectiveScale === undefined || media.effectiveScale >= 1, "warning", "raster-resolution", "Raster effective resolution is below one source pixel per displayed pixel.", { slideNumber: media.slideNumber, objectId: media.objectId, assetId: media.assetId, effectiveScale: media.effectiveScale });
  }

  const planAssets = Array.isArray(plan.assets) ? plan.assets : [];
  const compositeVisualAssets = regions
    .filter((region) => region.sourceConsumption !== "semantic-child" && region.visualAssets && typeof region.visualAssets === "object")
    .flatMap((region) => Array.isArray(region.visualAssets.items)
      ? region.visualAssets.items.map((item) => ({ parentRegion: region, item }))
      : []);
  if (visualSchema) {
    for (const asset of planAssets) {
      if (asset.assetClass === "generic-icon") {
        addCheck(["extract-raster", "regenerate-icon"].includes(asset.action), "error", "generic-icon-default-action", "Every generic raster icon must use clean extract-raster or regenerate-icon under the schema 1.1 default policy.", { assetId: asset.assetId, slideNumber: asset.slideNumber, action: asset.action });
        if (asset.action === "extract-raster") addCheck(asset.extractionQuality === "clean", "error", "generic-icon-extraction-quality", "A generic icon retained from the source must declare clean extraction quality.", { assetId: asset.assetId, slideNumber: asset.slideNumber, extractionQuality: asset.extractionQuality });
      }
      if (["logo", "photo", "product-ui", "evidence-screenshot", "official-diagram"].includes(asset.assetClass)) addCheck(asset.action !== "regenerate-icon", "error", "authentic-asset-not-generated", "Logos, photos, product UI, evidence screenshots, and official diagrams must not be regenerated.", { assetId: asset.assetId, slideNumber: asset.slideNumber, assetClass: asset.assetClass, action: asset.action });
    }
  }
  const assetForRegion = (region) => {
    const slideAssets = planAssets.filter((item) => Number(item.slideNumber) === region.slideNumber);
    return slideAssets.find((item) => typeof item.regionId === "string" && item.regionId === region.regionId) ??
      slideAssets.find((item) => typeof region.assetId === "string" && item.assetId === region.assetId) ??
      slideAssets.find((item) => item.action === region.action && Array.isArray(item.sourceObjectIds) && sameStringSet(item.sourceObjectIds, region.sourceObjectIds));
  };
  const finalMediaForRegion = (region) => {
    const frame = finalFrames.get(region.slideNumber);
    if (!frame) return [];
    let regionBox;
    try { regionBox = bboxToPixels(region.bbox, frame); } catch { return []; }
    return finalRasterMetadata.filter((media) => media.slideNumber === region.slideNumber && boxesSubstantiallyMatch(regionBox, media.bbox));
  };
  const verifyAssetSource = async (asset, region, label) => {
    const assetId = asset.assetId;
    const declaredHash = typeof asset.sourceSha256 === "string" ? asset.sourceSha256.toLowerCase() : "";
    const expectedHashes = new Set(region.sourceObjectIds.flatMap(sourceObjectHashes));
    const expectedMediaRefs = new Set(region.sourceObjectIds
      .map((id) => normalizePackagePart(sourceObjectEvidence(id)?.mediaRef))
      .filter(Boolean));
    addCheck(asset.action === region.action, "error", `${label}-asset-action`, "Asset action must exactly match its conversion region action.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, assetAction: asset.action, regionAction: region.action });
    addCheck(sameStringSet(asset.sourceObjectIds, region.sourceObjectIds), "error", `${label}-asset-source-objects`, "Asset sourceObjectIds must exactly match the bound conversion region; partial overlap is not provenance.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, assetSourceObjectIds: asset.sourceObjectIds, regionSourceObjectIds: region.sourceObjectIds });
    addCheck(typeof asset.reason === "string" && asset.reason.trim() !== "", "error", `${label}-asset-reason`, "Raster asset provenance requires a nonblank reason.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId });
    const editability = Array.isArray(asset.editability) ? asset.editability : [asset.editability];
    addCheck(editability.filter(Boolean).map(String).includes("raster-replaceable"), "error", `${label}-asset-editability`, "Raster assets must be labeled raster-replaceable.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId });
    addCheck(/^[0-9a-f]{64}$/i.test(declaredHash), "error", `${label}-asset-source-hash`, "Raster asset sourceSha256 must be a valid SHA-256 value.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, declared: asset.sourceSha256 });
    addCheck(expectedHashes.size > 0, "error", `${label}-asset-manifest-hash`, "The source manifest must provide media hash evidence for every raster asset source object.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, sourceObjectIds: region.sourceObjectIds });
    addCheck(expectedHashes.has(declaredHash), "error", `${label}-asset-source-binding`, "Asset sourceSha256 must match media hash evidence for its exact source object IDs.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, declared: declaredHash, expected: [...expectedHashes] });

    const hasInputPath = typeof asset.inputPath === "string" && asset.inputPath.trim() !== "";
    const hasMediaRef = typeof asset.mediaRef === "string" && asset.mediaRef.trim() !== "";
    addCheck(hasInputPath || hasMediaRef, "error", `${label}-asset-source`, "Raster asset provenance requires inputPath or mediaRef.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId });
    if (hasInputPath) {
      const inputPath = path.isAbsolute(asset.inputPath) ? path.resolve(asset.inputPath) : path.resolve(path.dirname(planPath), asset.inputPath);
      const exists = await fileExists(inputPath);
      addCheck(exists, "error", `${label}-asset-input-file`, "Asset inputPath must resolve to an existing file.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, inputPath });
      if (exists) {
        const actualHash = await sha256File(inputPath);
        addCheck(actualHash === declaredHash, "error", `${label}-asset-input-hash`, "Asset inputPath bytes must match sourceSha256.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, inputPath, declared: declaredHash, actual: actualHash });
      }
    }
    if (hasMediaRef) {
      const mediaRef = normalizePackagePart(asset.mediaRef);
      addCheck(expectedMediaRefs.size === 0 || expectedMediaRefs.has(mediaRef), "error", `${label}-asset-media-ref-binding`, "Asset mediaRef must match the package part recorded for its exact source object IDs.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, mediaRef, expected: [...expectedMediaRefs] });
      const packageEntry = sourceZip.file(mediaRef);
      addCheck(Boolean(packageEntry), "error", `${label}-asset-media-ref`, "Asset mediaRef must resolve inside the exact source PPTX package.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, mediaRef });
      if (packageEntry) {
        const actualHash = sha256Bytes(await packageEntry.async("nodebuffer"));
        addCheck(actualHash === declaredHash, "error", `${label}-asset-media-hash`, "Source PPTX mediaRef bytes must match sourceSha256.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, mediaRef, declared: declaredHash, actual: actualHash });
      }
    }
  };
  const verifyAssetOutput = async (asset, region, label, required) => {
    const assetId = asset.assetId;
    const hasOutputPath = typeof asset.outputPath === "string" && asset.outputPath.trim() !== "";
    const declaredHash = typeof asset.outputSha256 === "string" ? asset.outputSha256.toLowerCase() : "";
    const hasOutputHash = /^[0-9a-f]{64}$/i.test(declaredHash);
    if (required || hasOutputPath || asset.outputSha256 !== undefined) {
      addCheck(hasOutputPath && hasOutputHash, "error", `${label}-asset-output`, "Asset outputPath and a valid outputSha256 are required together.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, outputPath: asset.outputPath, outputSha256: asset.outputSha256 });
    }
    if (!hasOutputPath || !hasOutputHash) return;
    const outputPath = path.isAbsolute(asset.outputPath) ? path.resolve(asset.outputPath) : path.resolve(path.dirname(planPath), asset.outputPath);
    const exists = await fileExists(outputPath);
    addCheck(exists, "error", `${label}-asset-output-file`, "Asset outputPath must resolve to an existing file.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, outputPath });
    if (exists) {
      const actualHash = await sha256File(outputPath);
      addCheck(actualHash === declaredHash, "error", `${label}-asset-output-hash`, "Asset outputPath bytes must match outputSha256.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, outputPath, declared: declaredHash, actual: actualHash });
    }
    const embeddedMedia = finalMediaForRegion(region);
    addCheck(embeddedMedia.some((media) => String(media.sha256 ?? "").toLowerCase() === declaredHash), "error", `${label}-asset-final-binding`, "Asset outputSha256 must match the raster bytes actually embedded in the final PPTX at the planned region.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId, declared: declaredHash, finalHashes: embeddedMedia.map((media) => media.sha256).filter(Boolean) });
  };
  if (visualSchema) {
    for (const { parentRegion, item } of compositeVisualAssets) {
      const assetId = typeof item?.assetId === "string" ? item.assetId : "";
      const disposition = String(item?.disposition ?? "");
      const virtualRegion = {
        slideNumber: parentRegion.slideNumber,
        regionId: `visual-${assetId}`,
        bbox: item?.bbox,
        sourceObjectIds: Array.isArray(item?.sourceObjectIds) ? item.sourceObjectIds : parentRegion.sourceObjectIds,
        action: disposition,
      };
      const asset = planAssets.find((candidate) => candidate?.assetId === assetId);
      const requiresRasterOutput = ["retain-raster", "extract-raster", "regenerate-icon"].includes(disposition);
      addCheck(Boolean(asset) || !requiresRasterOutput, "error", "composite-visual-asset-record", "Every retained, extracted, or regenerated visual asset embedded in a composite screenshot requires a matching provenance record.", { slideNumber: parentRegion.slideNumber, parentRegionId: parentRegion.regionId, assetId, disposition });
      if (!asset || !requiresRasterOutput) continue;
      addCheck(asset.action === disposition, "error", "composite-visual-asset-action", "Composite visual-asset disposition must match its provenance action.", { slideNumber: parentRegion.slideNumber, parentRegionId: parentRegion.regionId, assetId, disposition, action: asset.action });
      await verifyAssetSource(asset, virtualRegion, "composite-visual");
      await verifyAssetOutput(asset, virtualRegion, "composite-visual", true);
      const media = finalMediaForRegion(virtualRegion);
      addCheck(media.length > 0, "error", "composite-visual-asset-present", "A planned visible visual asset from a composite screenshot is missing from the final PPTX.", { slideNumber: parentRegion.slideNumber, parentRegionId: parentRegion.regionId, assetId, disposition });
    }
    for (const parentRegion of regions.filter((region) => region.sourceConsumption !== "semantic-child" && ["rebuild-text", "rebuild-shape", "rebuild-table", "rebuild-chart"].includes(region.action))) {
      const sourceLargeRaster = parentRegion.sourceObjectIds.some((id) => sourceObjectEvidence(id)?.coverageRatio >= LARGE_RASTER_RATIO);
      if (sourceLargeRaster) addCheck(Boolean(parentRegion.visualAssets?.complete === true), "error", "composite-visual-inventory", "Deleting a large source raster requires a complete visual-assets inventory, even when the inventory is explicitly empty.", { slideNumber: parentRegion.slideNumber, regionId: parentRegion.regionId });
    }
  }
  for (const region of regions.filter((item) => item.action === "retain-raster" || item.action === "extract-raster")) {
    const asset = assetForRegion(region);
    addCheck(Boolean(asset), "error", "raster-asset-record", "Every retained or extracted raster requires a matching asset/provenance record.", { slideNumber: region.slideNumber, regionId: region.regionId, action: region.action });
    if (!asset) continue;
    await verifyAssetSource(asset, region, "raster");
    await verifyAssetOutput(asset, region, "raster", region.action === "extract-raster");
    if (region.action === "retain-raster" && typeof asset.outputSha256 === "string") {
      addCheck(asset.outputSha256.toLowerCase() === String(asset.sourceSha256 ?? "").toLowerCase(), "error", "retained-asset-output-hash", "A retain-raster output hash, when recorded, must equal its source hash; intentional pixel changes require extract-raster.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId: asset.assetId, sourceSha256: asset.sourceSha256, outputSha256: asset.outputSha256 });
    }
  }
  for (const region of regions.filter((item) => item.action === "regenerate-icon")) {
    const asset = assetForRegion(region);
    addCheck(Boolean(asset), "error", "generated-asset-record", "Every regenerated icon requires a matching asset/provenance record.", { slideNumber: region.slideNumber, regionId: region.regionId });
    if (asset) {
      await verifyAssetSource(asset, region, "generated");
      await verifyAssetOutput(asset, region, "generated", true);
      addCheck(typeof asset.prompt === "string" && asset.prompt.trim() !== "" && typeof asset.provenance === "string" && asset.provenance.trim() !== "", "error", "generated-asset-provenance", "Regenerated icon records require prompt and provenance.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId: asset.assetId });
      if (visualSchema) {
        addCheck(asset.assetClass === "generic-icon", "error", "generated-icon-class", "Only generic icons may use regenerate-icon in schema 1.1.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId: asset.assetId, assetClass: asset.assetClass });
        addCheck(typeof asset.semanticConcept === "string" && asset.semanticConcept.trim() !== "", "error", "generated-icon-concept", "A generated generic icon requires a semantic concept.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId: asset.assetId });
        addCheck(String(asset.iconFamilyId ?? "") === String(iconProfile.familyId ?? iconProfile.id ?? ""), "error", "generated-icon-family", "A generated icon must belong to the frozen icon family.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId: asset.assetId, expected: iconProfile.familyId ?? iconProfile.id, actual: asset.iconFamilyId });
      }
    }
    const iconMedia = finalMediaForRegion(region)[0];
    addCheck(Boolean(iconMedia), "error", "generated-icon-present", "A regenerated icon must resolve to an image asset in the planned region.", { slideNumber: region.slideNumber, regionId: region.regionId });
    if (iconMedia) {
      addCheck(Boolean(iconMedia.hasAlpha && iconMedia.alpha?.transparentPixelRatio > 0.01), "error", "generated-icon-alpha", "A regenerated icon must contain genuine transparent alpha, not an opaque rectangular tile.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId: iconMedia.assetId, alpha: iconMedia.alpha });
      addCheck((iconMedia.alpha?.edgeVisiblePixelRatio ?? 0) <= 0.02, visualSchema ? "error" : "warning", "generated-icon-edge", "Visible icon pixels touch the raster edge; this indicates a seam or clipping risk.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId: iconMedia.assetId, alpha: iconMedia.alpha });
      if (visualSchema) {
        addCheck((iconMedia.alpha?.opaqueTileScore ?? 1) < 0.75, "error", "generated-icon-opaque-tile", "Generated icons must not contain an opaque rectangular tile or white square backing.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId: iconMedia.assetId, alpha: iconMedia.alpha });
        const centroid = iconMedia.alpha?.alphaVisibleCentroid;
        const centroidTolerance = Number(iconProfile.centroidTolerance ?? DEFAULT_ICON_CENTROID_TOLERANCE);
        addCheck(Boolean(centroid) && Math.abs(centroid.normalizedX - 0.5) <= centroidTolerance && Math.abs(centroid.normalizedY - 0.5) <= centroidTolerance, "warning", "generated-icon-centroid", "Generated icon optical mass is off-center relative to the frozen icon family.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId: iconMedia.assetId, centroid, centroidTolerance });
        const opticalTarget = Number(iconProfile.targetOpticalCoverage ?? iconProfile.opticalCoverage ?? DEFAULT_ICON_OPTICAL_TARGET);
        const opticalTolerance = Number(iconProfile.opticalTolerance ?? DEFAULT_ICON_OPTICAL_TOLERANCE);
        addCheck(Number.isFinite(iconMedia.alpha?.opticalCoverage) && Math.abs(iconMedia.alpha.opticalCoverage - opticalTarget) <= opticalTolerance, "warning", "generated-icon-optical-size", "Generated icon optical coverage differs from the frozen family target.", { slideNumber: region.slideNumber, regionId: region.regionId, assetId: iconMedia.assetId, actual: iconMedia.alpha?.opticalCoverage, target: opticalTarget, tolerance: opticalTolerance });
      }
      addCheck((iconMedia.areaRatio ?? 0) < LARGE_RASTER_RATIO, "error", "generated-icon-size", "A regenerated icon must not become a large slide raster.", { slideNumber: region.slideNumber, regionId: region.regionId, areaRatio: iconMedia.areaRatio });
    }
    if (visualSchema) {
      const frame = finalFrames.get(region.slideNumber);
      let regionBox;
      try { regionBox = frame ? bboxToPixels(region.bbox, frame) : undefined; } catch { regionBox = undefined; }
      const iconObjects = semanticElements.filter((item) => item.slideNumber === region.slideNumber && item.semantic?.part === "icon" && objectContainedInRegion(regionBox, item.bbox, 0.8));
      addCheck(iconObjects.length === 1, "error", "generated-icon-semantic-object", "Each regenerated icon region must contain exactly one semantically named part=icon image.", { slideNumber: region.slideNumber, regionId: region.regionId, matches: iconObjects.map((item) => item.element?.aid ?? item.element?.id) });
      if (iconObjects.length === 1) {
        const iconSemantic = iconObjects[0].semantic;
        const iconBox = iconObjects[0].bbox;
        addCheck(Boolean(iconSemantic.family && iconSemantic.instance), "error", "generated-icon-semantic-family", "A generated icon object must declare family and instance in its semantic name.", { slideNumber: region.slideNumber, regionId: region.regionId, name: iconObjects[0].element?.name });
        const supports = semanticElements.filter((item) => item.slideNumber === region.slideNumber && item.semantic?.family === iconSemantic.family && item.semantic?.instance === iconSemantic.instance && ["support", "support-circle"].includes(item.semantic?.part));
        addCheck(supports.length <= 1, "error", "generated-icon-double-support", "An icon may have at most one native support shape; double circles are forbidden.", { slideNumber: region.slideNumber, regionId: region.regionId, supportObjects: supports.map((item) => item.element?.aid ?? item.element?.id) });
        if (iconProfile.supportRequired === true) addCheck(supports.length === 1, "error", "generated-icon-support-required", "The frozen icon family requires exactly one native support shape.", { slideNumber: region.slideNumber, regionId: region.regionId, supportObjects: supports.map((item) => item.element?.aid ?? item.element?.id) });
        if (supports.length === 1) {
          addCheck(objectContainedInRegion(supports[0].bbox, iconBox, 0.95), "error", "generated-icon-support-containment", "The transparent icon image must remain inside its single native support shape.", { slideNumber: region.slideNumber, regionId: region.regionId, iconBBox: iconBox, supportBBox: supports[0].bbox });
        }
        const unrelatedText = semanticElements.filter((item) => {
          if (item.slideNumber !== region.slideNumber || !item.bbox || !String(item.element?.text ?? item.element?.textPreview ?? "").trim()) return false;
          if (item.semantic?.family === iconSemantic.family && item.semantic?.instance === iconSemantic.instance) return false;
          const overlap = intersectionArea(iconBox, item.bbox);
          const centerX = item.bbox.left + item.bbox.width / 2;
          const centerY = item.bbox.top + item.bbox.height / 2;
          const centerCovered = centerX >= iconBox.left && centerX <= iconBox.left + iconBox.width && centerY >= iconBox.top && centerY <= iconBox.top + iconBox.height;
          return centerCovered || overlap / Math.max(1, area(item.bbox)) > 0.12;
        });
        addCheck(unrelatedText.length === 0, "error", "generated-icon-occlusion", "A generated icon must not cover unrelated titles, body text, labels, or badges.", { slideNumber: region.slideNumber, regionId: region.regionId, iconBBox: iconBox, coveredObjects: unrelatedText.map((item) => ({ objectId: item.element?.aid ?? item.element?.id, name: item.element?.name, bbox: item.bbox })) });
      }
    }
    const notes = finalRecords.find((record) => String(record.kind).toLowerCase() === "notes" && Number(record.slide) === region.slideNumber);
    const notesText = recordText(notes);
    addCheck(/\[Sources\]/i.test(notesText) && /(ImageGen|OpenAI)/i.test(notesText), "error", "generated-icon-sources", "Slides with regenerated icons require a [Sources] note naming ImageGen/OpenAI provenance.", { slideNumber: region.slideNumber, regionId: region.regionId });
  }

  for (const diff of sourceVsFinalDiffs) {
    addCheck(diff.dimensionsMatch, "error", "render-dimensions-source-final", "Source and final rendered slide dimensions must match.", { slideNumber: diff.slideNumber });
    addCheck(diff.outsideMaskMeanAbsoluteDifference <= 50 && diff.outsideMaskChangedPixelRatio <= 0.65, "warning", "visual-fidelity", "Source-to-final visual difference outside authorized redesign masks is high; inspect the full-size diff and both slides.", { slideNumber: diff.slideNumber, meanAbsoluteDifference: diff.meanAbsoluteDifference, changedPixelRatio: diff.changedPixelRatio, outsideMaskMeanAbsoluteDifference: diff.outsideMaskMeanAbsoluteDifference, outsideMaskChangedPixelRatio: diff.outsideMaskChangedPixelRatio, maskedPixelRatio: diff.maskedPixelRatio, masks: diff.masks });
  }
  for (const diff of finalVsRoundtripDiffs) {
    addCheck(diff.dimensionsMatch, "error", "render-dimensions-roundtrip", "Final and round-trip rendered slide dimensions must match.", { slideNumber: diff.slideNumber });
    addCheck(diff.meanAbsoluteDifference <= ROUNDTRIP_MEAN_DIFF_LIMIT && diff.changedPixelRatio <= ROUNDTRIP_PIXEL_DIFF_LIMIT, "error", "roundtrip-visual", "Round-trip render drift exceeds the allowed tolerance.", { slideNumber: diff.slideNumber, meanAbsoluteDifference: diff.meanAbsoluteDifference, changedPixelRatio: diff.changedPixelRatio });
  }

  const finalText = finalRecords.filter((record) => kindMatches(record, "text")).map((record) => `${record.slide}|${normalizeText(recordText(record))}`).sort();
  const roundtripText = roundtripRecords.filter((record) => kindMatches(record, "text")).map((record) => `${record.slide}|${normalizeText(recordText(record))}`).sort();
  addCheck(JSON.stringify(finalText) === JSON.stringify(roundtripText), "error", "roundtrip-text", "Round-trip must preserve all final native text records.");
  for (const kind of ["image", "table", "chart"]) {
    const finalCount = finalRecords.filter((record) => String(record.kind).toLowerCase() === kind).length;
    const roundtripCount = roundtripRecords.filter((record) => String(record.kind).toLowerCase() === kind).length;
    addCheck(finalCount === roundtripCount, "error", `roundtrip-${kind}-count`, `Round-trip must preserve the final ${kind} count.`, { finalCount, roundtripCount });
  }

  const sourceHashAfter = await sha256File(source);
  const finalHashAfter = await sha256File(finalPptx);
  addCheck(sourceHashAfter === sourceHashBefore, "error", "source-unchanged", "Source SHA-256 changed during QA; source mutation is forbidden.");
  addCheck(finalHashAfter === finalHashBefore, "error", "final-hash-bound", "Final SHA-256 changed during QA; report no longer binds the delivered file.");

  const exceptionEntry = (region) => ({
    slideNumber: region.slideNumber,
    regionId: region.regionId,
    parentRegionId: region.parentRegionId,
    rootRegionId: region.rootRegionId,
    isSemanticSubregion: region.isSemanticSubregion,
    action: region.action,
    reason: region.reason,
  });
  const exceptions = {
    retainedRaster: regions.filter((region) => region.action === "retain-raster").map(exceptionEntry),
    regeneratedIcons: regions.filter((region) => region.action === "regenerate-icon").map(exceptionEntry),
    manualReview: regions.filter((region) => region.action === "manual-review").map(exceptionEntry),
  };
  const editability = {
    textEditable: regions.filter((region) => region.action === "rebuild-text" || String(region.targetType).toLowerCase().includes("text")).length,
    structureEditable: regions.filter((region) => ["rebuild-shape", "rebuild-table", "rebuild-chart"].includes(region.action) || String(region.targetType).toLowerCase().includes("shape")).length,
    dataEditable: regions.filter((region) => ["rebuild-table", "rebuild-chart"].includes(region.action)).length,
    rasterReplaceable: regions.filter((region) => RASTER_ACTIONS.has(region.action)).length,
  };
  const errors = checks.filter((check) => check.status === "fail" && check.severity === "error").length;
  const warnings = checks.filter((check) => check.status === "fail" && check.severity === "warning").length;
  const status = errors > 0 ? "fail" : warnings > 0 ? "warning" : "pass";
  const reportPath = path.join(workspace, "qa-report.json");
  const ledgerPath = path.join(workspace, "qa-ledger.txt");
  const report = {
    schema: visualSchema ? "make-pptx-editable/qa-report/v1.1" : "make-pptx-editable/qa-report/v1",
    status,
    generatedAt: new Date().toISOString(),
    inputs: { source, final: finalPptx, plan: planPath, workspace, sourceManifest: sourceManifestPath },
    hashes: {
      sourceBefore: sourceHashBefore,
      sourceAfter: sourceHashAfter,
      finalBefore: finalHashBefore,
      finalAfter: finalHashAfter,
      roundtrip: roundtripHash,
    },
    summary: {
      sourceSlides: sourceEvidence.slides.length,
      finalSlides: finalEvidence.slides.length,
      roundtripSlides: roundtripEvidence.slides.length,
      planRegions: regions.length,
      semanticSubregions: regions.filter((region) => region.isSemanticSubregion).length,
      errors,
      warnings,
      finalLargeRasters: finalLargeRasters.length,
    },
    editability,
    visualConsistency,
    exceptions,
    packageFeatures: {
      source: { ...sourcePackageFeatures, ...sourceXmlFeatures, emptyPlaceholders: sourceXmlFeatures.emptyPlaceholders.length, placeholderText: sourceXmlFeatures.placeholderText.length },
      final: { ...finalPackageFeatures, ...finalXmlFeatures, emptyPlaceholders: finalXmlFeatures.emptyPlaceholders.length, placeholderText: finalXmlFeatures.placeholderText.length },
      roundtrip: { ...roundtripPackageFeatures, ...roundtripXmlFeatures, emptyPlaceholders: roundtripXmlFeatures.emptyPlaceholders.length, placeholderText: roundtripXmlFeatures.placeholderText.length },
    },
    rasterEvidence: {
      finalLargeRasters,
      media: finalRasterMetadata,
    },
    visualDiffs: {
      sourceVsFinal: sourceVsFinalDiffs.map((item) => ({ ...item, diffPath: relativeTo(workspace, item.diffPath) })),
      finalVsRoundtrip: finalVsRoundtripDiffs.map((item) => ({ ...item, diffPath: relativeTo(workspace, item.diffPath) })),
    },
    artifacts: {
      root: relativeTo(workspace, artifactRoot),
      sourceInspect: relativeTo(workspace, sourceEvidence.inspectPath),
      finalInspect: relativeTo(workspace, finalEvidence.inspectPath),
      roundtripInspect: relativeTo(workspace, roundtripEvidence.inspectPath),
      roundtripPptx: relativeTo(workspace, roundtripPptx),
      renders: relativeTo(workspace, path.join(artifactRoot, "renders")),
      layouts: relativeTo(workspace, path.join(artifactRoot, "layouts")),
      diffs: relativeTo(workspace, diffDir),
    },
    externalChecks: {
      templateFidelity: "required-separately-by-presentations-skill",
      slideOverflow: "required-separately-by-presentations-skill",
      fullSizeVisualReview: "required",
      desktopPowerPointSpotCheck: "recommended-for-fonts-or-advanced-effects",
    },
    checks,
  };
  await atomicWrite(reportPath, Buffer.from(`${JSON.stringify(report, null, 2)}\n`, "utf8"));
  await atomicWrite(ledgerPath, Buffer.from(`${ledgerText(report)}\n`, "utf8"));
  console.log(JSON.stringify({ status, report: reportPath, ledger: ledgerPath, errors, warnings, finalSha256: finalHashBefore }, null, 2));
  // On Windows, Artifact Tool/renderer teardown can occasionally surface a
  // late 0xC0000409 after all awaited outputs have been flushed. Exit
  // explicitly once the hash-bound reports are durable so a successful run is
  // not misreported as a crash. A failed QA run still returns a non-zero code.
  process.exit(status === "fail" ? 1 : 0);
}

main().catch(async (error) => {
  console.error(`qa_conversion: ${error.message || String(error)}`);
  console.error(usage());
  process.exit(1);
});
