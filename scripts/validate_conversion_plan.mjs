#!/usr/bin/env node

import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

export const ALLOWED_ACTIONS = Object.freeze([
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

const ACTION_SET = new Set(ALLOWED_ACTIONS);
const SUPPORTED_SCHEMA_VERSIONS = new Set(["1.0", "1.1"]);
const PLAN_MODES = new Set(["balanced", "maximum-editability", "fidelity-first"]);
const VISUAL_INTENTS = new Set(["faithful-rebuild", "style-normalization", "user-approved-redesign"]);
const ASSET_CLASSES = new Set([
  "generic-icon",
  "logo",
  "photo",
  "product-ui",
  "evidence-screenshot",
  "official-diagram",
  "complex-illustration",
  "other",
]);
const ISSUE_SCOPES = new Set(["local", "component-family", "deck"]);
const ISSUE_STATUSES = new Set(["open", "fixed", "waived"]);
const SEMANTIC_ID = /^[a-z0-9][a-z0-9_-]*$/;
const SUPPORTED_BBOX_UNITS = new Set(["px", "in"]);
const PX_PER_INCH = 96;
const LOW_CONFIDENCE_THRESHOLD = 0.8;
const MIN_DESTRUCTIVE_OBJECT_COVERAGE = 0.95;
const MIN_DESTRUCTIVE_IOU = 0.8;
const MAX_DESTRUCTIVE_REGION_EXPANSION = 1.25;
const ALLOWED_TARGET_TYPE_COMPONENTS = new Set([
  "text",
  "shape",
  "group",
  "connector",
  "image",
  "icon",
  "table",
  "chart",
  "object",
]);
const TARGET_TYPE_ALIASES = new Map([
  ["shapes", "shape"],
  ["groups", "group"],
  ["connectors", "connector"],
  ["images", "image"],
  ["icons", "icon"],
  ["native table", "table"],
  ["native-table", "table"],
  ["native chart", "chart"],
  ["native-chart", "chart"],
  ["native object", "object"],
  ["native-object", "object"],
  ["preserved object", "object"],
  ["preserved-object", "object"],
]);
const ACTION_TARGET_CONTRACTS = new Map([
  ["keep-native", { anyOf: ["text", "shape", "group", "connector", "image", "icon", "table", "chart", "object"] }],
  ["rebuild-text", { require: ["text"], allow: ["text", "shape", "group", "image", "icon"] }],
  ["rebuild-shape", { anyOf: ["shape", "group", "connector"], allow: ["text", "shape", "group", "connector", "image", "icon"] }],
  ["rebuild-table", { require: ["table"], allow: ["table", "text", "shape"] }],
  ["rebuild-chart", { require: ["chart"], allow: ["chart", "text", "shape"] }],
  ["extract-raster", { anyOf: ["image", "icon"], allow: ["image", "icon"] }],
  ["regenerate-icon", { anyOf: ["image", "icon"], allow: ["image", "icon"] }],
  ["retain-raster", { require: ["image"], allow: ["image", "icon"] }],
  ["manual-review", { anyOf: ["text", "shape", "group", "connector", "image", "icon", "table", "chart", "object"] }],
]);
const EDITABILITY_LABELS = new Set([
  "text-editable",
  "structure-editable",
  "data-editable",
  "raster-replaceable",
]);
const SLIDE_CLASSIFICATIONS = new Set([
  "native-editable",
  "mixed",
  "flattened",
  "low-quality-scan",
]);
const REQUIRED_REGION_FIELDS = [
  "slideNumber",
  "bbox",
  "sourceObjectIds",
  "action",
  "targetType",
  "expectedText",
  "confidence",
  "reason",
  "editability",
];
const DESTRUCTIVE_ACTIONS = new Set([
  "rebuild-text",
  "rebuild-shape",
  "rebuild-table",
  "rebuild-chart",
  "extract-raster",
  "regenerate-icon",
]);
const RASTER_ASSET_ACTIONS = new Set([
  "extract-raster",
  "regenerate-icon",
  "retain-raster",
]);
const COMPOSITE_VISUAL_DISPOSITIONS = new Set([
  "retain-raster",
  "extract-raster",
  "regenerate-icon",
  "rebuild-native",
  "manual-review",
]);
const SEMANTIC_SUBREGION_ACTIONS = new Set([
  "rebuild-text",
  "rebuild-shape",
  "rebuild-table",
  "rebuild-chart",
  "extract-raster",
  "regenerate-icon",
  "manual-review",
]);
const UNSAFE_ID_TOKENS = new Set([
  "*",
  "all",
  "all-images",
  "all_images",
  "images/*",
  "image/*",
  "image:*",
  "//*",
]);
const UNSAFE_KEYS = new Set([
  "clearallimages",
  "clearrasterbody",
  "deleteall",
  "deleteallimages",
  "matchall",
  "removeall",
  "removeallimages",
  "selectallimages",
  "wildcard",
]);

const HELP = `Validate an editable-PPTX conversion plan against an analysis manifest.

Usage:
  $RUNTIME_NODE validate_conversion_plan.mjs \\
    --manifest <source-manifest.json> \\
    --plan <conversion-plan.json> [--stage calibration|final] \\
    --out-map <template-frame-map.json> [--report <validation-report.json>]

Required plan region fields:
  slideNumber, bbox, sourceObjectIds, action, targetType,
  expectedText, confidence, reason, editability

Allowed actions:
  ${ALLOWED_ACTIONS.join(", ")}

Rules:
  * Source object identifiers must match manifest objects exactly and be on the
    same slide as the plan region.
  * Every slide must be represented and every manifest large raster must have
    an explicit disposition.
  * Region boxes must be finite, positive, within the slide, and non-overlapping
    unless allowOverlap:true or a shared non-empty overlapGroup is supplied.
  * Region and manifest boxes may use px or in; comparisons and output zones are
    normalized to the manifest slide's pixel canvas.
  * Destructive regions must safely cover their bound source-object footprint;
    an incidental intersection is never enough to authorize deletion.
  * A composite source image is consumed once by a full-footprint parent;
    optional subregions describe nested semantic output zones without re-deleting it.
  * schemaVersion, mode, source manifest/source PPTX/source SHA-256, target type,
    editability, and the 0.80 destructive-confidence gate are enforced.
  * Schema 1.1 also enforces role/component profiles, staged calibration,
    annotated feedback accounting, visual intent, and extract-or-regenerate icon policy.
  * --stage calibration emits write access only for representative sample slides;
    --stage final (the default) requires frozen calibration evidence.
  * Wildcards and delete-all/clear-all selectors are rejected.
`;

function parseArgs(argv) {
  const parsed = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (token === "--help" || token === "-h") {
      parsed.help = true;
      continue;
    }
    if (!token.startsWith("--")) {
      throw new Error(`Unexpected positional argument: ${token}`);
    }
    const key = token.slice(2);
    if (!["manifest", "plan", "out-map", "report", "stage"].includes(key)) {
      throw new Error(`Unknown option: ${token}`);
    }
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) {
      throw new Error(`Missing value for ${token}`);
    }
    parsed[key] = value;
    index += 1;
  }
  return parsed;
}

function sha256(buffer) {
  return crypto.createHash("sha256").update(buffer).digest("hex");
}

function canonicalJson(value) {
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(",")}]`;
  if (value && typeof value === "object") {
    return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${canonicalJson(value[key])}`).join(",")}}`;
  }
  return JSON.stringify(value);
}

function canonicalJsonSha256(value) {
  return sha256(Buffer.from(canonicalJson(value), "utf8"));
}

async function readJson(filePath, label) {
  let bytes;
  try {
    bytes = await fs.readFile(filePath);
  } catch (error) {
    throw new Error(`Cannot read ${label} ${JSON.stringify(filePath)}: ${error.message}`);
  }
  try {
    return { value: JSON.parse(bytes.toString("utf8")), bytes };
  } catch (error) {
    throw new Error(`Invalid JSON in ${label} ${JSON.stringify(filePath)}: ${error.message}`);
  }
}

function hasOwn(value, key) {
  return Object.prototype.hasOwnProperty.call(value, key);
}

function sameStringSet(left, right) {
  const a = [...new Set((Array.isArray(left) ? left : []).map(String))].sort();
  const b = [...new Set((Array.isArray(right) ? right : []).map(String))].sort();
  return a.length === b.length && a.every((value, index) => value === b[index]);
}

function asPositiveInteger(value) {
  const number = Number(value);
  return Number.isInteger(number) && number > 0 ? number : null;
}

function normalizeUnit(value, context, { required = false, defaultUnit = null } = {}) {
  const normalized = String(value ?? defaultUnit ?? "").trim().toLowerCase();
  if (!normalized) {
    if (required) throw new Error(`${context}.unit must declare px or in.`);
    return null;
  }
  const canonical = normalized === "inch" || normalized === "inches" ? "in" : normalized;
  if (!SUPPORTED_BBOX_UNITS.has(canonical)) {
    throw new Error(`${context}.unit must be one of: ${[...SUPPORTED_BBOX_UNITS].join(", ")}.`);
  }
  return canonical;
}

function normalizeBbox(raw, context, { requireUnit = false, defaultUnit = null } = {}) {
  let left;
  let top;
  let width;
  let height;
  let unit;
  if (Array.isArray(raw) && raw.length >= 4) {
    [left, top, width, height] = raw;
  } else if (raw && typeof raw === "object") {
    left = raw.left ?? raw.x;
    top = raw.top ?? raw.y;
    width = raw.width ?? raw.w;
    height = raw.height ?? raw.h;
    unit = raw.unit;
  } else {
    throw new Error(`${context} must be an object or [left, top, width, height].`);
  }
  const bbox = {
    left: Number(left),
    top: Number(top),
    width: Number(width),
    height: Number(height),
    unit: normalizeUnit(unit, context, { required: requireUnit, defaultUnit }),
  };
  for (const key of ["left", "top", "width", "height"]) {
    if (!Number.isFinite(bbox[key])) {
      throw new Error(`${context}.${key} must be a finite number.`);
    }
  }
  if (bbox.width <= 0 || bbox.height <= 0) {
    throw new Error(`${context} must have positive width and height.`);
  }
  return bbox;
}

function normalizeDimensions(raw, context) {
  if (!raw || typeof raw !== "object") return null;
  const width = Number(raw.width);
  const height = Number(raw.height);
  if (!(width > 0) || !(height > 0)) return null;
  try {
    return {
      width,
      height,
      unit: normalizeUnit(raw.unit, context, { required: false, defaultUnit: "px" }),
    };
  } catch {
    return null;
  }
}

function bboxIntersection(a, b) {
  const left = Math.max(a.left, b.left);
  const top = Math.max(a.top, b.top);
  const right = Math.min(a.left + a.width, b.left + b.width);
  const bottom = Math.min(a.top + a.height, b.top + b.height);
  const width = Math.max(0, right - left);
  const height = Math.max(0, bottom - top);
  return { left, top, width, height, area: width * height };
}

function slideCoordinateSpace(manifest, slide, slideNumber) {
  const candidates = [
    normalizeDimensions(slide?.dimensions, `slide ${slideNumber}.dimensions`),
    normalizeDimensions(slide?.slideSize, `slide ${slideNumber}.slideSize`),
    normalizeDimensions(slide?.renderDimensionsPx, `slide ${slideNumber}.renderDimensionsPx`),
    normalizeDimensions(slide?.canvasPx, `slide ${slideNumber}.canvasPx`),
    normalizeDimensions(manifest?.deck?.slideSize, "manifest.deck.slideSize"),
    normalizeDimensions(manifest?.deck?.renderSlideSizePx, "manifest.deck.renderSlideSizePx"),
    normalizeDimensions(manifest?.slideSize, "manifest.slideSize"),
    normalizeDimensions(manifest?.renderSlideSizePx, "manifest.renderSlideSizePx"),
  ].filter(Boolean);
  const px = candidates.find((entry) => entry.unit === "px");
  const inches = candidates.find((entry) => entry.unit === "in");
  if (!px && !inches) return null;
  const width = px?.width ?? inches.width * PX_PER_INCH;
  const height = px?.height ?? inches.height * PX_PER_INCH;
  const widthIn = inches?.width ?? width / PX_PER_INCH;
  const heightIn = inches?.height ?? height / PX_PER_INCH;
  if (!(width > 0) || !(height > 0) || !(widthIn > 0) || !(heightIn > 0)) return null;
  return {
    width,
    height,
    unit: "px",
    widthIn,
    heightIn,
    pxPerInX: width / widthIn,
    pxPerInY: height / heightIn,
  };
}

function bboxToPixels(bbox, space, context) {
  if (!bbox?.unit) throw new Error(`${context}.unit must declare px or in.`);
  if (bbox.unit === "px") return { ...bbox, unit: "px" };
  if (bbox.unit === "in") {
    return {
      left: bbox.left * space.pxPerInX,
      top: bbox.top * space.pxPerInY,
      width: bbox.width * space.pxPerInX,
      height: bbox.height * space.pxPerInY,
      unit: "px",
    };
  }
  throw new Error(`${context}.unit ${JSON.stringify(bbox.unit)} is unsupported.`);
}

function collectManifestObjects(manifest, manifestSlides) {
  const collected = [];
  const seenTopLevelIds = new Set();
  for (const object of Array.isArray(manifest.objects) ? manifest.objects : []) {
    const objectId = object?.objectId ?? object?.id;
    collected.push({ ...object, objectId });
    if (typeof objectId === "string" && objectId) seenTopLevelIds.add(objectId);
  }
  for (const slide of manifestSlides) {
    for (const object of Array.isArray(slide?.objects) ? slide.objects : []) {
      const objectId = object?.objectId ?? object?.id;
      if (typeof objectId === "string" && seenTopLevelIds.has(objectId)) continue;
      collected.push({ ...object, objectId, slideNumber: object.slideNumber ?? slide.slideNumber });
    }
  }
  return collected;
}

function expectedTextIsValid(value) {
  if (typeof value === "string") return true;
  if (!Array.isArray(value)) return false;
  return value.every((entry) => (
    typeof entry === "string"
    || (entry && typeof entry === "object" && typeof entry.text === "string")
  ));
}

function expectedTextIsEmpty(value) {
  if (typeof value === "string") return !value.trim();
  if (!Array.isArray(value)) return true;
  return value.every((entry) => {
    if (typeof entry === "string") return !entry.trim();
    return !entry?.text?.trim();
  });
}

function validateExpectedTextLedger(value, context, errors) {
  if (!expectedTextIsValid(value)) {
    errors.push(
      `${context}.expectedText must be a string or an array of strings/text-ledger entries `
        + "(use an empty string or array when not applicable).",
    );
    return false;
  }
  if (!Array.isArray(value)) return true;
  for (const [index, entry] of value.entries()) {
    if (!entry || typeof entry !== "object") continue;
    if (
      hasOwn(entry, "confidence")
      && (
        typeof entry.confidence !== "number"
        || !Number.isFinite(entry.confidence)
        || entry.confidence < 0
        || entry.confidence > 1
      )
    ) {
      errors.push(`${context}.expectedText[${index}].confidence must be a number from 0 through 1.`);
    }
    if (hasOwn(entry, "needsReview") && typeof entry.needsReview !== "boolean") {
      errors.push(`${context}.expectedText[${index}].needsReview must be boolean when supplied.`);
    }
  }
  return true;
}

function targetTypes(value) {
  const rawComponents = (Array.isArray(value) ? value.map(String) : String(value ?? "")
    .split(/[+,/|]/u)
    .map((entry) => entry.trim())
    .filter(Boolean));
  return [...new Set(rawComponents.map((entry) => {
    const lower = entry.toLowerCase().replace(/\s+/gu, " ");
    return TARGET_TYPE_ALIASES.get(lower) ?? lower;
  }))];
}

function validateTargetType(action, rawValue, context, errors) {
  if (typeof rawValue !== "string" || !rawValue.trim()) {
    errors.push(`${context}.targetType must be a non-empty string.`);
    return [];
  }
  const components = targetTypes(rawValue);
  if (!components.length) {
    errors.push(`${context}.targetType must contain at least one target component.`);
    return [];
  }
  const unknown = components.filter((entry) => !ALLOWED_TARGET_TYPE_COMPONENTS.has(entry));
  if (unknown.length) {
    errors.push(
      `${context}.targetType contains unsupported component(s): ${unknown.join(", ")}. `
        + `Allowed components: ${[...ALLOWED_TARGET_TYPE_COMPONENTS].join(", ")}.`,
    );
    return components;
  }
  const contract = ACTION_TARGET_CONTRACTS.get(action);
  if (!contract) return components;
  if (contract.require && contract.require.some((entry) => !components.includes(entry))) {
    errors.push(`${context}.targetType is incompatible with ${action}; it must include ${contract.require.join(" + ")}.`);
  }
  if (contract.anyOf && !contract.anyOf.some((entry) => components.includes(entry))) {
    errors.push(`${context}.targetType is incompatible with ${action}; include one of: ${contract.anyOf.join(", ")}.`);
  }
  if (contract.allow) {
    const disallowed = components.filter((entry) => !contract.allow.includes(entry));
    if (disallowed.length) {
      errors.push(`${context}.targetType component(s) ${disallowed.join(", ")} are not allowed for ${action}.`);
    }
  }
  return components;
}

function validateEditability(rawValue, targetComponents, action, context, errors) {
  if (!Array.isArray(rawValue) || rawValue.length === 0) {
    errors.push(`${context}.editability must be a non-empty array of editability labels.`);
    return [];
  }
  const labels = [];
  const seen = new Set();
  for (const [index, value] of rawValue.entries()) {
    if (typeof value !== "string" || !value.trim()) {
      errors.push(`${context}.editability[${index}] must be a non-empty string.`);
      continue;
    }
    const label = value.trim();
    if (!EDITABILITY_LABELS.has(label)) {
      errors.push(
        `${context}.editability[${index}] is ${JSON.stringify(label)}; `
          + `allowed labels are ${[...EDITABILITY_LABELS].join(", ")}.`,
      );
      continue;
    }
    if (seen.has(label)) {
      errors.push(`${context}.editability duplicates ${JSON.stringify(label)}.`);
      continue;
    }
    seen.add(label);
    labels.push(label);
  }
  if (action === "manual-review") return labels;

  const required = new Set();
  if (targetComponents.includes("text")) required.add("text-editable");
  if (targetComponents.some((entry) => ["shape", "group", "connector"].includes(entry))) {
    required.add("structure-editable");
  }
  if (targetComponents.some((entry) => ["table", "chart"].includes(entry))) {
    required.add("structure-editable");
    required.add("data-editable");
  }
  if (targetComponents.some((entry) => ["image", "icon"].includes(entry))) {
    required.add("raster-replaceable");
  }
  for (const label of required) {
    if (!seen.has(label)) {
      errors.push(`${context}.editability must include ${JSON.stringify(label)} for targetType ${targetComponents.join("+")}.`);
    }
  }
  return labels;
}

function textLedgerNeedsManualReview(value) {
  if (!Array.isArray(value)) return false;
  return value.some((entry) => (
    entry
    && typeof entry === "object"
    && (
      entry.needsReview === true
      || (typeof entry.confidence === "number" && entry.confidence < LOW_CONFIDENCE_THRESHOLD)
    )
  ));
}

function bboxArea(bbox) {
  return bbox.width * bbox.height;
}

function bboxUnion(boxes) {
  const left = Math.min(...boxes.map((box) => box.left));
  const top = Math.min(...boxes.map((box) => box.top));
  const right = Math.max(...boxes.map((box) => box.left + box.width));
  const bottom = Math.max(...boxes.map((box) => box.top + box.height));
  return { left, top, width: right - left, height: bottom - top, unit: "px" };
}

function bboxContains(outer, inner, tolerance = 0.5) {
  return outer.left <= inner.left + tolerance
    && outer.top <= inner.top + tolerance
    && outer.left + outer.width >= inner.left + inner.width - tolerance
    && outer.top + outer.height >= inner.top + inner.height - tolerance;
}

function bboxMetrics(region, footprint) {
  const intersection = bboxIntersection(region, footprint);
  const regionArea = bboxArea(region);
  const footprintArea = bboxArea(footprint);
  const unionArea = regionArea + footprintArea - intersection.area;
  return {
    intersectionArea: intersection.area,
    regionCoverage: regionArea > 0 ? intersection.area / regionArea : 0,
    sourceCoverage: footprintArea > 0 ? intersection.area / footprintArea : 0,
    iou: unionArea > 0 ? intersection.area / unionArea : 0,
    expansion: footprintArea > 0 ? regionArea / footprintArea : Number.POSITIVE_INFINITY,
  };
}

function normalizedPathKey(filePath) {
  const resolved = path.resolve(filePath);
  return process.platform === "win32" ? resolved.toLowerCase() : resolved;
}

function samePath(first, second) {
  return normalizedPathKey(first) === normalizedPathKey(second);
}

function resolveOwnedPath(rawValue, ownerPath) {
  if (typeof rawValue !== "string" || !rawValue.trim()) return null;
  return path.resolve(path.dirname(path.resolve(ownerPath)), rawValue.trim());
}

async function fileSha256IfPresent(filePath) {
  try {
    return sha256(await fs.readFile(filePath));
  } catch (error) {
    if (error.code === "ENOENT") return null;
    throw error;
  }
}

function isBareFileName(rawValue) {
  return typeof rawValue === "string"
    && rawValue.trim()
    && path.basename(rawValue.trim()) === rawValue.trim();
}

async function validatePlanIdentity({ manifest, plan, manifestPath, planPath, errors, warnings }) {
  if (!SUPPORTED_SCHEMA_VERSIONS.has(String(manifest.schemaVersion ?? ""))) {
    errors.push(
      `manifest.schemaVersion must be one of: ${[...SUPPORTED_SCHEMA_VERSIONS].join(", ")}; `
        + `received ${JSON.stringify(manifest.schemaVersion)}.`,
    );
  }
  if (!SUPPORTED_SCHEMA_VERSIONS.has(String(plan.schemaVersion ?? ""))) {
    errors.push(
      `plan.schemaVersion must be one of: ${[...SUPPORTED_SCHEMA_VERSIONS].join(", ")}; `
        + `received ${JSON.stringify(plan.schemaVersion)}.`,
    );
  }
  if (!PLAN_MODES.has(plan.mode)) {
    errors.push(`plan.mode must be one of: ${[...PLAN_MODES].join(", ")}; received ${JSON.stringify(plan.mode)}.`);
  }

  if (plan.sourceManifest == null || plan.sourceManifest === "") {
    warnings.push("plan.sourceManifest is omitted; the explicit --manifest argument is the authoritative binding.");
  } else if (typeof plan.sourceManifest !== "string") {
    errors.push("plan.sourceManifest must be a non-empty path string when supplied.");
  } else {
    const declaredManifest = resolveOwnedPath(plan.sourceManifest, planPath);
    if (!samePath(declaredManifest, manifestPath)) {
      errors.push(
        `plan.sourceManifest resolves to ${JSON.stringify(declaredManifest)}, `
          + `not the validated manifest ${JSON.stringify(path.resolve(manifestPath))}.`,
      );
    }
  }

  const manifestHashes = [
    manifest.sourceSha256,
    manifest.source?.sha256,
    manifest.source?.copySha256,
  ].filter((value) => value != null && value !== "");
  for (const [index, value] of manifestHashes.entries()) {
    if (typeof value !== "string" || !/^[a-f0-9]{64}$/iu.test(value)) {
      errors.push(`Manifest source SHA-256 value ${index + 1} is not a 64-character hexadecimal digest.`);
    }
  }
  const normalizedManifestHashes = [...new Set(
    manifestHashes
      .filter((value) => typeof value === "string" && /^[a-f0-9]{64}$/iu.test(value))
      .map((value) => value.toLowerCase()),
  )];
  if (normalizedManifestHashes.length === 0) {
    errors.push("Manifest must declare sourceSha256 or source.sha256.");
  } else if (normalizedManifestHashes.length > 1) {
    errors.push("Manifest source SHA-256 fields disagree; source identity is ambiguous.");
  }
  const expectedSha = normalizedManifestHashes[0] ?? null;
  if (typeof plan.sourceSha256 !== "string" || !/^[a-f0-9]{64}$/iu.test(plan.sourceSha256)) {
    errors.push("plan.sourceSha256 must be a 64-character hexadecimal digest.");
  } else if (expectedSha && plan.sourceSha256.toLowerCase() !== expectedSha) {
    errors.push(`plan.sourceSha256 does not match the source manifest (${plan.sourceSha256} != ${expectedSha}).`);
  }

  const manifestSourceCandidates = [];
  const addManifestCandidate = (rawValue) => {
    if (typeof rawValue !== "string" || !rawValue.trim()) return;
    const candidate = path.isAbsolute(rawValue)
      ? path.resolve(rawValue)
      : resolveOwnedPath(rawValue, manifestPath);
    if (!manifestSourceCandidates.some((entry) => samePath(entry, candidate))) {
      manifestSourceCandidates.push(candidate);
    }
  };
  addManifestCandidate(manifest.source?.inputPath);
  addManifestCandidate(manifest.source?.copiedPath);
  addManifestCandidate(manifest.sourcePptx);

  if (typeof plan.sourcePptx !== "string" || !plan.sourcePptx.trim()) {
    errors.push("plan.sourcePptx must be a non-empty path string.");
  } else if (path.extname(plan.sourcePptx).toLowerCase() !== ".pptx") {
    errors.push("plan.sourcePptx must identify a .pptx file.");
  }
  const declaredSourcePath = resolveOwnedPath(plan.sourcePptx, planPath);
  let matchedSourcePath = null;
  if (declaredSourcePath) {
    matchedSourcePath = manifestSourceCandidates.find((candidate) => samePath(candidate, declaredSourcePath)) ?? null;
    if (!matchedSourcePath && isBareFileName(plan.sourcePptx)) {
      const declaredBase = path.basename(plan.sourcePptx).toLowerCase();
      matchedSourcePath = manifestSourceCandidates.find(
        (candidate) => path.basename(candidate).toLowerCase() === declaredBase,
      ) ?? null;
    }
    const declaredSha = await fileSha256IfPresent(declaredSourcePath);
    if (declaredSha) {
      if (expectedSha && declaredSha !== expectedSha) {
        errors.push(`plan.sourcePptx exists but its SHA-256 does not match the manifest: ${declaredSourcePath}.`);
      } else if (!matchedSourcePath) {
        matchedSourcePath = declaredSourcePath;
      }
    }
  }
  if (!matchedSourcePath) {
    errors.push(
      `plan.sourcePptx ${JSON.stringify(plan.sourcePptx)} does not match any source path recorded by the manifest.`,
    );
  }

  let verifiedManifestSource = false;
  for (const candidate of manifestSourceCandidates) {
    const digest = await fileSha256IfPresent(candidate);
    if (!digest) continue;
    if (expectedSha && digest !== expectedSha) {
      errors.push(`Manifest source candidate has an unexpected SHA-256: ${candidate}.`);
      continue;
    }
    verifiedManifestSource = true;
  }
  if (!verifiedManifestSource) {
    errors.push("No manifest-recorded source PPTX exists with the declared source SHA-256.");
  }

  let outputPath = null;
  if (plan.outputPptx == null || plan.outputPptx === "") {
    errors.push("plan.outputPptx is required and must name a new sibling _editable.pptx file.");
  } else {
    if (typeof plan.outputPptx !== "string") {
      errors.push("plan.outputPptx must be a path string when supplied.");
    } else {
      outputPath = resolveOwnedPath(plan.outputPptx, planPath);
      if (path.extname(plan.outputPptx).toLowerCase() !== ".pptx") {
        errors.push("plan.outputPptx must use the .pptx extension.");
      }
      if (
        (declaredSourcePath && samePath(outputPath, declaredSourcePath))
        || manifestSourceCandidates.some((candidate) => samePath(outputPath, candidate))
        || String(plan.outputPptx).trim().toLowerCase() === String(plan.sourcePptx).trim().toLowerCase()
      ) {
        errors.push("plan.outputPptx must not overwrite plan.sourcePptx or any source path recorded by the manifest.");
      }
      if (!/_editable\.pptx$/iu.test(path.basename(outputPath))) {
        errors.push("plan.outputPptx must use the _editable.pptx suffix.");
      }
    }
  }

  return {
    expectedSha,
    declaredSourcePath,
    matchedSourcePath,
    manifestSourceCandidates,
    outputPath,
  };
}

function flattenPlanRegions(plan) {
  const regions = [];
  const explicitSlideNumbers = new Set();
  const pushList = (list, inheritedSlideNumber, source) => {
    if (!Array.isArray(list)) return;
    list.forEach((item, index) => {
      if (!item || typeof item !== "object" || Array.isArray(item)) {
        regions.push({ __invalid: true, __source: `${source}[${index}]`, value: item });
        return;
      }
      regions.push({
        ...item,
        ...(item.slideNumber == null && inheritedSlideNumber != null
          ? { slideNumber: inheritedSlideNumber }
          : {}),
        __source: `${source}[${index}]`,
      });
    });
  };

  pushList(plan.regions, null, "plan.regions");
  pushList(plan.items, null, "plan.items");
  pushList(plan.actions, null, "plan.actions");
  if (Array.isArray(plan.slides)) {
    plan.slides.forEach((slideEntry, slideIndex) => {
      if (!slideEntry || typeof slideEntry !== "object" || Array.isArray(slideEntry)) {
        regions.push({ __invalid: true, __source: `plan.slides[${slideIndex}]`, value: slideEntry });
        return;
      }
      const number = asPositiveInteger(slideEntry.slideNumber ?? slideEntry.number ?? slideEntry.index);
      if (number != null) explicitSlideNumbers.add(number);
      const before = regions.length;
      pushList(slideEntry.regions, number, `plan.slides[${slideIndex}].regions`);
      pushList(slideEntry.items, number, `plan.slides[${slideIndex}].items`);
      pushList(slideEntry.actions, number, `plan.slides[${slideIndex}].actions`);
      if (regions.length === before && hasOwn(slideEntry, "action")) {
        regions.push({ ...slideEntry, slideNumber: number, __source: `plan.slides[${slideIndex}]` });
      }
    });
  }
  return { regions, explicitSlideNumbers };
}

function truthyUnsafeValue(value) {
  if (value === false || value === 0 || value == null || value === "") return false;
  return true;
}

function scanForUnsafeSemantics(value, errors, location = "plan") {
  if (!value || typeof value !== "object") return;
  if (Array.isArray(value)) {
    value.forEach((entry, index) => scanForUnsafeSemantics(entry, errors, `${location}[${index}]`));
    return;
  }
  for (const [key, child] of Object.entries(value)) {
    const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, "");
    const childPath = `${location}.${key}`;
    if (UNSAFE_KEYS.has(normalizedKey) && truthyUnsafeValue(child)) {
      errors.push(`${childPath} expresses forbidden wildcard/delete-all semantics.`);
    }
    if (["selector", "query", "objectselector", "imageselector"].includes(normalizedKey)
      && typeof child === "string"
      && /(^|[/:])(?:\*|all(?:[-_ ]?images?)?)(?:$|[/:])/i.test(child.trim())) {
      errors.push(`${childPath} contains a forbidden wildcard/all-images selector.`);
    }
    scanForUnsafeSemantics(child, errors, childPath);
  }
}

function ensureParent(filePath) {
  return fs.mkdir(path.dirname(path.resolve(filePath)), { recursive: true });
}

async function writeJson(filePath, value) {
  await ensureParent(filePath);
  await fs.writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

async function removeStaleMap(filePath) {
  if (!filePath) return;
  try {
    await fs.unlink(filePath);
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
  }
}

function makeReportBase(manifestPath, planPath) {
  return {
    schemaVersion: "1.1",
    generatedAt: new Date().toISOString(),
    manifestPath: path.resolve(manifestPath),
    planPath: path.resolve(planPath),
    valid: false,
    errors: [],
    warnings: [],
    statistics: {},
  };
}

function validateSemanticMap(value, label, errors) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    errors.push(`${label} must be an object keyed by semantic IDs.`);
    return new Set();
  }
  const ids = new Set();
  for (const [id, record] of Object.entries(value)) {
    if (!SEMANTIC_ID.test(id)) errors.push(`${label} key ${JSON.stringify(id)} is not a lowercase semantic ID.`);
    if (!record || typeof record !== "object" || Array.isArray(record)) errors.push(`${label}.${id} must be an object.`);
    ids.add(id);
  }
  return ids;
}

function validateVisualContract(plan, slidesByNumber, objectsById, errors, warnings, stage = "final") {
  const isV11 = String(plan.schemaVersion) === "1.1";
  if (!isV11) {
    warnings.push("Legacy schema 1.0 plan does not claim role, component-family, calibration, feedback-closure, or extract-or-regenerate generic-icon guarantees.");
    return { isV11: false, roleIds: new Set(), familyIds: new Set(), issueIds: new Set() };
  }
  const policy = plan.visualPolicy;
  if (!policy || typeof policy !== "object" || Array.isArray(policy)) {
    errors.push("plan.visualPolicy is required for schema 1.1.");
  } else {
    if (policy.normalizationScope !== "full-deck-role-based") errors.push("plan.visualPolicy.normalizationScope must be full-deck-role-based.");
    if (policy.iconMode !== "extract-or-regenerate-generic") errors.push("plan.visualPolicy.iconMode must be extract-or-regenerate-generic.");
    if (policy.qaStrictness !== "tiered") errors.push("plan.visualPolicy.qaStrictness must be tiered.");
    if (!new Set(["automatic", "user-gated"]).has(policy.calibrationMode)) errors.push("plan.visualPolicy.calibrationMode must be automatic or user-gated.");
  }

  const profile = plan.styleProfile;
  if (!profile || typeof profile !== "object" || Array.isArray(profile)) errors.push("plan.styleProfile is required for schema 1.1.");
  const roleIds = validateSemanticMap(profile?.roles, "plan.styleProfile.roles", errors);
  const familyIds = validateSemanticMap(profile?.componentFamilies ?? {}, "plan.styleProfile.componentFamilies", errors);
  if (!roleIds.size) errors.push("plan.styleProfile.roles must define at least one canonical text role.");
  for (const [roleId, role] of Object.entries(profile?.roles ?? {})) {
    if (!Number.isFinite(Number(role.fontSize)) || Number(role.fontSize) <= 0) errors.push(`plan.styleProfile.roles.${roleId}.fontSize must be positive pixels.`);
    if (!Number.isInteger(Number(role.maxLines)) || Number(role.maxLines) < 1) errors.push(`plan.styleProfile.roles.${roleId}.maxLines must be a positive integer.`);
    if (role.allowedEmphasis != null) {
      const allowedEmphasis = new Set(["bold", "italic", "underline", "color"]);
      if (!Array.isArray(role.allowedEmphasis) || role.allowedEmphasis.some((item) => !allowedEmphasis.has(item))) {
        errors.push(`plan.styleProfile.roles.${roleId}.allowedEmphasis must contain only bold, italic, underline, or color.`);
      }
    }
  }
  if (!profile?.iconFamily || typeof profile.iconFamily !== "object" || Array.isArray(profile.iconFamily)) {
    errors.push("plan.styleProfile.iconFamily is required for schema 1.1.");
  } else {
    if (!SEMANTIC_ID.test(String(profile.iconFamily.id ?? ""))) errors.push("plan.styleProfile.iconFamily.id must be a lowercase semantic ID.");
    if (typeof profile.iconFamily.sharedPromptPrefix !== "string" || !profile.iconFamily.sharedPromptPrefix.trim()) errors.push("plan.styleProfile.iconFamily.sharedPromptPrefix is required.");
    for (const [key, fallback] of [["opticalCoverage", 0.72], ["centroidTolerance", 0.04], ["opticalTolerance", 0.08]]) {
      const value = Number(profile.iconFamily[key] ?? fallback);
      if (!Number.isFinite(value) || value <= 0 || value > 1) errors.push(`plan.styleProfile.iconFamily.${key} must be greater than 0 and at most 1.`);
    }
  }

  const calibration = plan.calibration;
  if (!calibration || typeof calibration !== "object" || Array.isArray(calibration)) {
    errors.push("plan.calibration is required for schema 1.1.");
  } else {
    if (!new Set(["automatic", "user-gated"]).has(calibration.mode)) errors.push("plan.calibration.mode must be automatic or user-gated.");
    if (policy?.calibrationMode && calibration.mode !== policy.calibrationMode) errors.push("plan.calibration.mode must match visualPolicy.calibrationMode.");
    const requiredStatus = calibration.mode === "user-gated" ? "approved" : "complete";
    const allowedStatuses = stage === "calibration" ? new Set(["pending", requiredStatus]) : new Set([requiredStatus]);
    if (!allowedStatuses.has(calibration.status)) errors.push(`plan.calibration.status must be ${[...allowedStatuses].join(" or ")} for ${calibration.mode} mode during ${stage} validation.`);
    if (!Array.isArray(calibration.representativeSlides) || calibration.representativeSlides.length < 1 || calibration.representativeSlides.length > 3) {
      errors.push("plan.calibration.representativeSlides must contain one to three slide numbers.");
    } else {
      for (const slideNumber of calibration.representativeSlides) {
        if (!slidesByNumber.has(Number(slideNumber))) errors.push(`plan.calibration references unknown slide ${slideNumber}.`);
      }
    }
    if (stage === "final" || calibration.status !== "pending") {
      if (!Array.isArray(calibration.evidence) || calibration.evidence.length === 0) errors.push("plan.calibration.evidence must contain rendered calibration evidence after calibration completes.");
      if (!/^[0-9a-f]{64}$/i.test(String(calibration.frozenProfileSha256 ?? ""))) {
        errors.push("plan.calibration.frozenProfileSha256 must be a SHA-256 digest after calibration completes.");
      } else if (profile && calibration.frozenProfileSha256.toLowerCase() !== canonicalJsonSha256(profile)) {
        errors.push("plan.calibration.frozenProfileSha256 must equal the canonical SHA-256 of plan.styleProfile.");
      }
    }
  }

  const issues = plan.feedbackIssues ?? [];
  if (!Array.isArray(issues)) errors.push("plan.feedbackIssues must be an array.");
  const issueIds = new Set();
  for (const [index, issue] of (Array.isArray(issues) ? issues : []).entries()) {
    const context = `plan.feedbackIssues[${index}]`;
    if (!issue || typeof issue !== "object" || Array.isArray(issue)) {
      errors.push(`${context} must be an object.`);
      continue;
    }
    if (!SEMANTIC_ID.test(String(issue.issueId ?? ""))) errors.push(`${context}.issueId must be a lowercase semantic ID.`);
    else if (issueIds.has(issue.issueId)) errors.push(`${context}.issueId duplicates ${JSON.stringify(issue.issueId)}.`);
    else issueIds.add(issue.issueId);
    if (!slidesByNumber.has(Number(issue.slideNumber))) errors.push(`${context}.slideNumber references an unknown slide.`);
    if (!ISSUE_SCOPES.has(issue.scope)) errors.push(`${context}.scope must be local, component-family, or deck.`);
    if (!ISSUE_STATUSES.has(issue.status)) errors.push(`${context}.status must be open, fixed, or waived.`);
    if (typeof issue.category !== "string" || !issue.category.trim()) errors.push(`${context}.category is required.`);
    if (typeof issue.expectedFix !== "string" || !issue.expectedFix.trim()) errors.push(`${context}.expectedFix is required.`);
    if (!issue.annotationSource || typeof issue.annotationSource !== "string") errors.push(`${context}.annotationSource is required.`);
    if (!issue.annotationBbox) errors.push(`${context}.annotationBbox is required.`);
    else if (slidesByNumber.has(Number(issue.slideNumber))) {
      try {
        bboxToPixels(
          normalizeBbox(issue.annotationBbox, `${context}.annotationBbox`, { requireUnit: true }),
          slidesByNumber.get(Number(issue.slideNumber)).dimensions,
          `${context}.annotationBbox`,
        );
      } catch (error) {
        errors.push(error.message);
      }
    }
    if (!Array.isArray(issue.relatedObjectIds) || issue.relatedObjectIds.length === 0) errors.push(`${context}.relatedObjectIds must contain exact IDs.`);
    else for (const objectId of issue.relatedObjectIds) {
      const sourceObject = objectsById.get(String(objectId));
      if (!sourceObject) errors.push(`${context}.relatedObjectIds contains unknown object ${JSON.stringify(objectId)}.`);
      else if (Number(sourceObject.slideNumber) !== Number(issue.slideNumber)) errors.push(`${context}.relatedObjectIds contains object ${JSON.stringify(objectId)} from slide ${sourceObject.slideNumber}.`);
    }
    if (issue.scope === "component-family" && !familyIds.has(String(issue.componentFamily ?? ""))) errors.push(`${context}.componentFamily must reference styleProfile.componentFamilies.`);
    if (["component-family", "deck"].includes(issue.scope) && !Array.isArray(issue.inspectedInstanceIds)) errors.push(`${context}.inspectedInstanceIds is required for propagated issues.`);
    if (issue.status === "waived" && (typeof issue.waiverReason !== "string" || !issue.waiverReason.trim())) errors.push(`${context}.waiverReason is required when waived.`);
  }
  return { isV11, roleIds, familyIds, issueIds };
}

function validateVisualRegionFields(raw, context, visualContract, errors) {
  if (!visualContract.isV11) return;
  if (!VISUAL_INTENTS.has(raw.intent)) errors.push(`${context}.intent must be faithful-rebuild, style-normalization, or user-approved-redesign.`);
  if (raw.targetBbox != null && !["style-normalization", "user-approved-redesign"].includes(raw.intent)) {
    errors.push(`${context}.targetBbox requires style-normalization or user-approved-redesign intent.`);
  }
  if (raw.styleRole != null && !visualContract.roleIds.has(String(raw.styleRole))) errors.push(`${context}.styleRole must reference plan.styleProfile.roles.`);
  if (raw.componentFamily != null && !visualContract.familyIds.has(String(raw.componentFamily))) errors.push(`${context}.componentFamily must reference plan.styleProfile.componentFamilies.`);
  if (raw.issueRefs != null) {
    if (!Array.isArray(raw.issueRefs)) errors.push(`${context}.issueRefs must be an array.`);
    else for (const issueId of raw.issueRefs) if (!visualContract.issueIds.has(String(issueId))) errors.push(`${context}.issueRefs contains unknown issue ${JSON.stringify(issueId)}.`);
  }
  if (raw.intent === "user-approved-redesign") {
    if (typeof raw.authorization !== "string" || !raw.authorization.trim()) errors.push(`${context}.authorization is required for user-approved-redesign.`);
    if (raw.preserveText !== true) errors.push(`${context}.preserveText must be true for user-approved-redesign.`);
  }
}

export async function validateConversionPlan({ manifestPath, planPath, outMapPath, reportPath, stage = "final" }) {
  if (!new Set(["calibration", "final"]).has(stage)) throw new Error("stage must be calibration or final.");
  const report = makeReportBase(manifestPath, planPath);
  report.stage = stage;
  await removeStaleMap(outMapPath);
  const { value: manifest, bytes: manifestBytes } = await readJson(manifestPath, "manifest");
  const { value: plan, bytes: planBytes } = await readJson(planPath, "plan");
  const { errors, warnings } = report;

  if (!manifest || typeof manifest !== "object" || Array.isArray(manifest)) {
    errors.push("Manifest root must be a JSON object.");
  }
  if (!plan || typeof plan !== "object" || Array.isArray(plan)) {
    errors.push("Plan root must be a JSON object.");
  }
  if (errors.length) {
    await removeStaleMap(outMapPath);
    if (reportPath) await writeJson(reportPath, report);
    return { report, frameMap: null };
  }

  scanForUnsafeSemantics(plan, errors);
  const sourceIdentity = await validatePlanIdentity({
    manifest,
    plan,
    manifestPath,
    planPath,
    errors,
    warnings,
  });

  const manifestSlides = Array.isArray(manifest.slides) ? manifest.slides : [];
  const manifestObjects = collectManifestObjects(manifest, manifestSlides);
  const manifestLargeRasters = Array.isArray(manifest.largeRasters) ? manifest.largeRasters : [];
  if (!manifestSlides.length) errors.push("Manifest must contain a non-empty slides array.");
  if (!manifestObjects.length) warnings.push("Manifest objects array is empty; exact object binding cannot cover raster regions.");

  const slidesByNumber = new Map();
  for (const [index, slide] of manifestSlides.entries()) {
    const number = asPositiveInteger(slide?.slideNumber);
    if (number == null) {
      errors.push(`manifest.slides[${index}].slideNumber must be a positive integer.`);
      continue;
    }
    if (slidesByNumber.has(number)) {
      errors.push(`Duplicate manifest slideNumber ${number}.`);
      continue;
    }
    if (!SLIDE_CLASSIFICATIONS.has(slide.classification)) {
      errors.push(
        `Slide ${number} is not classified. Expected one of: ${[...SLIDE_CLASSIFICATIONS].join(", ")}.`,
      );
    }
    const dimensions = slideCoordinateSpace(manifest, slide, number);
    if (!dimensions) errors.push(`Slide ${number} has no valid dimensions.`);
    slidesByNumber.set(number, { slide, dimensions });
  }
  const declaredSlideCount = asPositiveInteger(manifest.deck?.slideCount ?? manifest.slideCount);
  if (declaredSlideCount != null && declaredSlideCount !== slidesByNumber.size) {
    errors.push(
      `Manifest deck.slideCount is ${declaredSlideCount}, but ${slidesByNumber.size} slide record(s) were found.`,
    );
  }

  const objectsById = new Map();
  const objectBoxesById = new Map();
  for (const [index, object] of manifestObjects.entries()) {
    const objectId = object?.objectId;
    if (typeof objectId !== "string" || !objectId.trim()) {
      errors.push(`manifest.objects[${index}].objectId must be a non-empty string.`);
      continue;
    }
    if (objectsById.has(objectId)) {
      errors.push(`Duplicate manifest objectId ${JSON.stringify(objectId)}.`);
      continue;
    }
    const number = asPositiveInteger(object.slideNumber);
    if (number == null || !slidesByNumber.has(number)) {
      errors.push(`Manifest object ${JSON.stringify(objectId)} references unknown slide ${object.slideNumber}.`);
    }
    objectsById.set(objectId, object);
    if (object.bbox && number != null && slidesByNumber.has(number)) {
      try {
        const rawBox = normalizeBbox(
          object.bbox,
          `manifest object ${objectId}.bbox`,
          { requireUnit: true },
        );
        objectBoxesById.set(
          objectId,
          bboxToPixels(rawBox, slidesByNumber.get(number).dimensions, `manifest object ${objectId}.bbox`),
        );
      } catch (error) {
        warnings.push(error.message);
      }
    }
  }
  const visualContract = validateVisualContract(plan, slidesByNumber, objectsById, errors, warnings, stage);
  const calibrationSlideNumbers = new Set((plan.calibration?.representativeSlides ?? []).map(Number));

  const largeRasterIds = new Set();
  for (const [index, raster] of manifestLargeRasters.entries()) {
    const objectId = raster?.objectId;
    if (typeof objectId !== "string" || !objectId.trim()) {
      errors.push(`manifest.largeRasters[${index}].objectId must be a non-empty string.`);
      continue;
    }
    if (!objectsById.has(objectId)) {
      errors.push(`Large raster ${JSON.stringify(objectId)} is absent from manifest.objects.`);
    } else if (Number(objectsById.get(objectId).slideNumber) !== Number(raster.slideNumber)) {
      errors.push(
        `Large raster ${JSON.stringify(objectId)} is listed on slide ${raster.slideNumber}, `
          + `but its manifest object belongs to slide ${objectsById.get(objectId).slideNumber}.`,
      );
    }
    largeRasterIds.add(objectId);
  }
  for (const slide of manifestSlides) {
    for (const objectId of slide?.largeRasterObjectIds ?? []) {
      if (typeof objectId === "string" && objectId) {
        if (!objectsById.has(objectId)) {
          errors.push(`Slide ${slide.slideNumber} largeRasterObjectIds references unknown object ${JSON.stringify(objectId)}.`);
        }
        largeRasterIds.add(objectId);
      }
    }
  }
  for (const object of manifestObjects) {
    const ratio = Number(object.areaRatio ?? object.coverageRatio ?? 0);
    if (
      (object.isLargeRaster === true || object.largeRaster === true || ratio >= 0.25)
      && /^(?:image|picture|raster)$/i.test(String(object.kind ?? object.type ?? ""))
      && object.slideLocal !== false
      && typeof object.objectId === "string"
      && object.objectId
    ) {
      largeRasterIds.add(object.objectId);
    }
  }

  const { regions: rawRegions, explicitSlideNumbers } = flattenPlanRegions(plan);
  if (!rawRegions.length) errors.push("Conversion plan contains no regions/actions.");
  if (Array.isArray(plan.slides)) {
    plan.slides.forEach((slideEntry, index) => {
      const number = asPositiveInteger(slideEntry?.slideNumber ?? slideEntry?.number ?? slideEntry?.index);
      if (number == null) {
        errors.push(`plan.slides[${index}] must declare a positive slideNumber.`);
      } else if (!slidesByNumber.has(number)) {
        errors.push(`plan.slides[${index}] references unknown slide ${number}.`);
      }
    });
  }
  const normalizedRegions = [];
  const representedSlides = new Set(explicitSlideNumbers);
  const explicitRegionIds = new Set();

  for (const [index, raw] of rawRegions.entries()) {
    const context = raw?.__source ?? `plan region ${index + 1}`;
    if (raw?.__invalid) {
      errors.push(`${context} must be an object.`);
      continue;
    }
    const missing = REQUIRED_REGION_FIELDS.filter((field) => !hasOwn(raw, field));
    if (missing.length) {
      errors.push(`${context} is missing required field(s): ${missing.join(", ")}.`);
    }

    const slideNumber = asPositiveInteger(raw.slideNumber);
    if (slideNumber == null || !slidesByNumber.has(slideNumber)) {
      errors.push(`${context}.slideNumber references an unknown slide: ${raw.slideNumber}.`);
      continue;
    }
    representedSlides.add(slideNumber);
    let bbox;
    try {
      const declaredBbox = normalizeBbox(
        raw.bbox,
        `${context}.bbox`,
        { requireUnit: true },
      );
      bbox = bboxToPixels(
        declaredBbox,
        slidesByNumber.get(slideNumber).dimensions,
        `${context}.bbox`,
      );
    } catch (error) {
      errors.push(error.message);
      continue;
    }
    const dimensions = slidesByNumber.get(slideNumber).dimensions;
    if (dimensions) {
      const tolerance = 0.5;
      if (
        bbox.left < -tolerance
        || bbox.top < -tolerance
        || bbox.left + bbox.width > dimensions.width + tolerance
        || bbox.top + bbox.height > dimensions.height + tolerance
      ) {
        errors.push(
          `${context}.bbox is outside slide ${slideNumber} bounds ${dimensions.width}x${dimensions.height}.`,
        );
      }
    }
    let targetBbox = bbox;
    if (raw.targetBbox != null) {
      try {
        targetBbox = bboxToPixels(
          normalizeBbox(raw.targetBbox, `${context}.targetBbox`, { requireUnit: true }),
          dimensions,
          `${context}.targetBbox`,
        );
        const tolerance = 0.5;
        if (
          targetBbox.left < -tolerance
          || targetBbox.top < -tolerance
          || targetBbox.left + targetBbox.width > dimensions.width + tolerance
          || targetBbox.top + targetBbox.height > dimensions.height + tolerance
        ) errors.push(`${context}.targetBbox is outside slide ${slideNumber} bounds ${dimensions.width}x${dimensions.height}.`);
      } catch (error) {
        errors.push(error.message);
      }
    }

    if (!Array.isArray(raw.sourceObjectIds) || raw.sourceObjectIds.length === 0) {
      errors.push(`${context}.sourceObjectIds must be a non-empty array of exact manifest object IDs.`);
      continue;
    }
    const sourceObjectIds = [];
    const seenIds = new Set();
    for (const [sourceIndex, rawId] of raw.sourceObjectIds.entries()) {
      const sourcePath = `${context}.sourceObjectIds[${sourceIndex}]`;
      if (typeof rawId !== "string" || !rawId.trim()) {
        errors.push(`${sourcePath} must be a non-empty string.`);
        continue;
      }
      const objectId = rawId.trim();
      if (UNSAFE_ID_TOKENS.has(objectId.toLowerCase()) || objectId.includes("*")) {
        errors.push(`${sourcePath} must be an exact object ID, not ${JSON.stringify(rawId)}.`);
        continue;
      }
      if (seenIds.has(objectId)) {
        errors.push(`${sourcePath} duplicates object ID ${JSON.stringify(objectId)} in the same region.`);
        continue;
      }
      seenIds.add(objectId);
      const object = objectsById.get(objectId);
      if (!object) {
        errors.push(`${sourcePath} does not exist in manifest.objects: ${JSON.stringify(objectId)}.`);
        continue;
      }
      if (Number(object.slideNumber) !== slideNumber) {
        errors.push(
          `${sourcePath} belongs to slide ${object.slideNumber}, not slide ${slideNumber}.`,
        );
        continue;
      }
      if (object.bbox) {
        const objectBbox = objectBoxesById.get(objectId);
        if (objectBbox && bboxIntersection(bbox, objectBbox).area === 0) {
          errors.push(`${context}.bbox does not intersect bound source object ${JSON.stringify(objectId)}.`);
        }
      }
      sourceObjectIds.push(objectId);
    }

    if (!ACTION_SET.has(raw.action)) {
      errors.push(`${context}.action must be one of: ${ALLOWED_ACTIONS.join(", ")}.`);
    }
    const targetTypeComponents = validateTargetType(raw.action, raw.targetType, context, errors);
    const expectedTextValid = validateExpectedTextLedger(raw.expectedText, context, errors);
    if (raw.action === "rebuild-text" && expectedTextValid && expectedTextIsEmpty(raw.expectedText)) {
      errors.push(`${context}.expectedText cannot be empty for rebuild-text; use manual-review if uncertain.`);
    }
    const confidence = raw.confidence;
    if (typeof confidence !== "number" || !Number.isFinite(confidence) || confidence < 0 || confidence > 1) {
      errors.push(`${context}.confidence must be a number from 0 through 1.`);
    }
    if (typeof raw.reason !== "string" || !raw.reason.trim()) {
      errors.push(`${context}.reason must be a non-empty evidence-based rationale.`);
    }
    const editability = validateEditability(
      raw.editability,
      targetTypeComponents,
      raw.action,
      context,
      errors,
    );
    validateVisualRegionFields(raw, context, visualContract, errors);
    if (
      DESTRUCTIVE_ACTIONS.has(raw.action)
      && typeof confidence === "number"
      && confidence < LOW_CONFIDENCE_THRESHOLD
    ) {
      errors.push(
        `${context}.confidence ${confidence} is below ${LOW_CONFIDENCE_THRESHOLD}; `
          + "a destructive action is forbidden. Change the region to manual-review.",
      );
    }
    if (DESTRUCTIVE_ACTIONS.has(raw.action) && textLedgerNeedsManualReview(raw.expectedText)) {
      errors.push(
        `${context}.expectedText contains an uncertain ledger entry; `
          + "a destructive action is forbidden until it is resolved or changed to manual-review.",
      );
    }
    if (raw.regionId != null) {
      if (typeof raw.regionId !== "string" || !raw.regionId.trim()) {
        errors.push(`${context}.regionId must be a non-empty string when supplied.`);
      } else if (explicitRegionIds.has(raw.regionId.trim())) {
        errors.push(`${context}.regionId duplicates ${JSON.stringify(raw.regionId.trim())}.`);
      } else {
        explicitRegionIds.add(raw.regionId.trim());
      }
    }

    const semanticSubregions = [];
    if (raw.subregions != null && !Array.isArray(raw.subregions)) {
      errors.push(`${context}.subregions must be an array when supplied.`);
    } else if (Array.isArray(raw.subregions)) {
      for (const [subIndex, subregion] of raw.subregions.entries()) {
        const subContext = `${context}.subregions[${subIndex}]`;
        if (!subregion || typeof subregion !== "object" || Array.isArray(subregion)) {
          errors.push(`${subContext} must be an object.`);
          continue;
        }
        const missingSubregionFields = REQUIRED_REGION_FIELDS.filter((field) => !hasOwn(subregion, field));
        if (missingSubregionFields.length) {
          errors.push(`${subContext} is missing required field(s): ${missingSubregionFields.join(", ")}.`);
        }
        const subSlideNumber = asPositiveInteger(subregion.slideNumber);
        if (subSlideNumber !== slideNumber) {
          errors.push(`${subContext}.slideNumber must equal its parent slide ${slideNumber}.`);
        }
        let subBbox;
        try {
          subBbox = bboxToPixels(
            normalizeBbox(subregion.bbox, `${subContext}.bbox`, { requireUnit: true }),
            slidesByNumber.get(slideNumber).dimensions,
            `${subContext}.bbox`,
          );
        } catch (error) {
          errors.push(error.message);
          continue;
        }
        const subArea = subBbox.width * subBbox.height;
        const insideArea = bboxIntersection(bbox, subBbox).area;
        if (subArea <= 0 || insideArea / subArea < 0.995) {
          errors.push(`${subContext}.bbox must be fully contained by its parent reconstruction region.`);
        }
        if (!Array.isArray(subregion.sourceObjectIds) || subregion.sourceObjectIds.length === 0) {
          errors.push(`${subContext}.sourceObjectIds must repeat exact source IDs owned by the parent region.`);
        }
        const subSourceObjectIds = [];
        for (const rawId of Array.isArray(subregion.sourceObjectIds) ? subregion.sourceObjectIds : []) {
          const objectId = typeof rawId === "string" ? rawId.trim() : "";
          if (!sourceObjectIds.includes(objectId)) {
            errors.push(`${subContext}.sourceObjectIds contains ${JSON.stringify(rawId)}, which is not owned by the parent region.`);
            continue;
          }
          if (!subSourceObjectIds.includes(objectId)) subSourceObjectIds.push(objectId);
        }
        if (!SEMANTIC_SUBREGION_ACTIONS.has(subregion.action)) {
          errors.push(`${subContext}.action must be one of: ${[...SEMANTIC_SUBREGION_ACTIONS].join(", ")}.`);
        }
        const subTargetTypeComponents = validateTargetType(subregion.action, subregion.targetType, subContext, errors);
        const subExpectedTextValid = validateExpectedTextLedger(subregion.expectedText, subContext, errors);
        if (subregion.action === "rebuild-text" && subExpectedTextValid && expectedTextIsEmpty(subregion.expectedText)) {
          errors.push(`${subContext}.expectedText cannot be empty for rebuild-text; use manual-review if uncertain.`);
        }
        const subConfidence = subregion.confidence;
        if (typeof subConfidence !== "number" || !Number.isFinite(subConfidence) || subConfidence < 0 || subConfidence > 1) {
          errors.push(`${subContext}.confidence must be a number from 0 through 1.`);
        }
        if (typeof subregion.reason !== "string" || !subregion.reason.trim()) {
          errors.push(`${subContext}.reason must be a non-empty evidence-based rationale.`);
        }
        const subEditability = validateEditability(
          subregion.editability,
          subTargetTypeComponents,
          subregion.action,
          subContext,
          errors,
        );
        validateVisualRegionFields(subregion, subContext, visualContract, errors);
        if (
          DESTRUCTIVE_ACTIONS.has(subregion.action)
          && typeof subConfidence === "number"
          && subConfidence < LOW_CONFIDENCE_THRESHOLD
        ) {
          errors.push(`${subContext}.confidence ${subConfidence} is below ${LOW_CONFIDENCE_THRESHOLD}; use manual-review.`);
        }
        if (DESTRUCTIVE_ACTIONS.has(subregion.action) && textLedgerNeedsManualReview(subregion.expectedText)) {
          errors.push(`${subContext}.expectedText contains an uncertain ledger entry; use manual-review.`);
        }
        const subregionIdRaw = subregion.subregionId ?? subregion.regionId;
        const subregionId = typeof subregionIdRaw === "string" && subregionIdRaw.trim()
          ? subregionIdRaw.trim()
          : `${typeof raw.regionId === "string" && raw.regionId.trim() ? raw.regionId.trim() : `slide-${slideNumber}-region-${index + 1}`}-sub-${subIndex + 1}`;
        if (explicitRegionIds.has(subregionId)) {
          errors.push(`${subContext} identifier duplicates ${JSON.stringify(subregionId)}.`);
        } else {
          explicitRegionIds.add(subregionId);
        }
        semanticSubregions.push({
          ...subregion,
          regionId: subregionId,
          slideNumber,
          bbox: subBbox,
          sourceObjectIds: subSourceObjectIds,
          confidence: subConfidence,
          targetTypeComponents: subTargetTypeComponents,
          editability: subEditability,
          sourceConsumption: "semantic-child",
          __source: subContext,
        });
      }
      for (let leftIndex = 0; leftIndex < semanticSubregions.length; leftIndex += 1) {
        for (let rightIndex = leftIndex + 1; rightIndex < semanticSubregions.length; rightIndex += 1) {
          const first = semanticSubregions[leftIndex];
          const second = semanticSubregions[rightIndex];
          if (first.allowOverlap === true || second.allowOverlap === true) continue;
          if (first.overlapGroup && first.overlapGroup === second.overlapGroup) continue;
          const overlap = bboxIntersection(first.bbox, second.bbox);
          const smallerArea = Math.min(first.bbox.width * first.bbox.height, second.bbox.width * second.bbox.height);
          if (overlap.area > 1 && smallerArea > 0 && overlap.area / smallerArea > 0.005) {
            errors.push(`${first.__source} and ${second.__source} overlap; declare allowOverlap or a shared overlapGroup only when intentional.`);
          }
        }
      }
      if (semanticSubregions.length > 0 && !DESTRUCTIVE_ACTIONS.has(raw.action)) {
        errors.push(`${context} must use one destructive rebuild action to own source deletion before declaring semantic subregions.`);
      }
    }

    normalizedRegions.push({
      ...raw,
      slideNumber,
      sourceBbox: bbox,
      bbox: targetBbox,
      sourceObjectIds,
      confidence,
      targetTypeComponents,
      editability,
      semanticSubregions,
      sourceConsumption: "owner",
      __source: context,
    });
  }

  const assetRecords = Array.isArray(plan.assets) ? plan.assets : [];
  if (plan.assets != null && !Array.isArray(plan.assets)) {
    errors.push("plan.assets must be an array when supplied.");
  }
  const assetIds = new Set();
  for (const [assetIndex, asset] of assetRecords.entries()) {
    const context = `plan.assets[${assetIndex}]`;
    if (!asset || typeof asset !== "object" || Array.isArray(asset)) {
      errors.push(`${context} must be an object.`);
      continue;
    }
    if (typeof asset.assetId !== "string" || !asset.assetId.trim()) {
      errors.push(`${context}.assetId must be a non-empty string.`);
    } else if (assetIds.has(asset.assetId.trim())) {
      errors.push(`${context}.assetId duplicates ${JSON.stringify(asset.assetId.trim())}.`);
    } else {
      assetIds.add(asset.assetId.trim());
    }
    const assetSlide = asPositiveInteger(asset.slideNumber);
    if (assetSlide == null || !slidesByNumber.has(assetSlide)) {
      errors.push(`${context}.slideNumber must reference a manifest slide.`);
    }
    if (!Array.isArray(asset.sourceObjectIds) || asset.sourceObjectIds.length === 0) {
      errors.push(`${context}.sourceObjectIds must contain exact source object IDs.`);
    } else {
      for (const rawId of asset.sourceObjectIds) {
        const objectId = typeof rawId === "string" ? rawId.trim() : "";
        const object = objectsById.get(objectId);
        if (!object) {
          errors.push(`${context}.sourceObjectIds contains unknown object ${JSON.stringify(rawId)}.`);
        } else if (assetSlide != null && Number(object.slideNumber) !== assetSlide) {
          errors.push(`${context}.sourceObjectIds contains object ${JSON.stringify(objectId)} from slide ${object.slideNumber}, not ${assetSlide}.`);
        }
      }
    }
    if (typeof asset.reason !== "string" || !asset.reason.trim()) {
      errors.push(`${context}.reason must be a non-empty rationale.`);
    }
    const assetEditability = Array.isArray(asset.editability) ? asset.editability : [asset.editability];
    if (!assetEditability.filter(Boolean).map(String).includes("raster-replaceable")) {
      errors.push(`${context}.editability must include "raster-replaceable".`);
    }
    if (typeof asset.sourceSha256 !== "string" || !/^[0-9a-f]{64}$/iu.test(asset.sourceSha256)) {
      errors.push(`${context}.sourceSha256 must be a 64-character SHA-256 digest.`);
    } else {
      const expectedObjectHashes = new Set((Array.isArray(asset.sourceObjectIds) ? asset.sourceObjectIds : [])
        .map((id) => objectsById.get(String(id)))
        .flatMap((object) => [object?.mediaSha256, object?.sha256])
        .filter(Boolean)
        .map((value) => String(value).toLowerCase()));
      if (expectedObjectHashes.size > 0 && !expectedObjectHashes.has(asset.sourceSha256.toLowerCase())) {
        errors.push(`${context}.sourceSha256 does not match any bound source object's media hash.`);
      }
    }
    if (!["extract-raster", "regenerate-icon", "retain-raster"].includes(asset.action)) {
      errors.push(`${context}.action must be extract-raster, regenerate-icon, or retain-raster.`);
    }
    if (visualContract.isV11) {
      if (!ASSET_CLASSES.has(asset.assetClass)) {
        errors.push(`${context}.assetClass must be one of: ${[...ASSET_CLASSES].join(", ")}.`);
      }
      if (asset.assetClass === "generic-icon") {
        if (!["extract-raster", "regenerate-icon"].includes(asset.action)) errors.push(`${context} generic-icon assets must use extract-raster or regenerate-icon under the default visual policy.`);
        if (typeof asset.semanticConcept !== "string" || !asset.semanticConcept.trim()) errors.push(`${context}.semanticConcept is required for generic-icon assets.`);
        const familyId = String(plan.styleProfile?.iconFamily?.id ?? "");
        if (asset.action === "regenerate-icon" && asset.iconFamilyId !== familyId) errors.push(`${context}.iconFamilyId must match plan.styleProfile.iconFamily.id.`);
        if (asset.action === "regenerate-icon" && (typeof asset.familyStylePrompt !== "string" || !asset.familyStylePrompt.trim())) {
          errors.push(`${context}.familyStylePrompt is required for generic-icon assets.`);
        } else if (asset.action === "regenerate-icon" && asset.familyStylePrompt.trim() !== String(plan.styleProfile?.iconFamily?.sharedPromptPrefix ?? "").trim()) {
          errors.push(`${context}.familyStylePrompt must exactly match the frozen icon-family sharedPromptPrefix.`);
        }
        if (asset.action === "extract-raster" && asset.extractionQuality !== "clean") errors.push(`${context}.extractionQuality must be \"clean\" when retaining a generic icon crop.`);
      } else if (asset.action === "regenerate-icon") {
        errors.push(`${context} may use regenerate-icon only when assetClass is generic-icon.`);
      }
    }
    if (asset.action === "regenerate-icon") {
      if (typeof asset.prompt !== "string" || !asset.prompt.trim() || typeof asset.provenance !== "string" || !asset.provenance.trim()) {
        errors.push(`${context} for regenerate-icon requires non-empty prompt and provenance.`);
      }
    } else if (!asset.inputPath && !asset.mediaRef) {
      errors.push(`${context} requires inputPath or mediaRef source provenance.`);
    }
  }

  // A single flattened screenshot can contain many semantic visual assets that
  // are invisible to OOXML inspection. Do not let a valid full-footprint delete
  // silently discard logos, illustrations, or icons embedded in that raster.
  const compositeVisualAssetIds = new Set();
  for (const region of normalizedRegions) {
    const sourceSlide = slidesByNumber.get(region.slideNumber);
    const ownsLargeRaster = region.sourceObjectIds.some((id) => largeRasterIds.has(id));
    const needsInventory = visualContract.isV11
      && ownsLargeRaster
      && DESTRUCTIVE_ACTIONS.has(region.action)
      && ["flattened", "low-quality-scan"].includes(String(sourceSlide?.slide?.classification ?? ""));
    if (!needsInventory) continue;
    const inventory = region.visualAssets;
    const inventoryContext = `${region.__source}.visualAssets`;
    if (!inventory || typeof inventory !== "object" || Array.isArray(inventory)) {
      errors.push(`${inventoryContext} is required when deleting a flattened large raster; inventory every embedded visual asset or explicitly declare an empty complete inventory.`);
      continue;
    }
    if (inventory.complete !== true) {
      errors.push(`${inventoryContext}.complete must be true before a flattened large raster may be deleted.`);
    }
    if (!Array.isArray(inventory.items)) {
      errors.push(`${inventoryContext}.items must be an array.`);
      continue;
    }
    if (inventory.items.length === 0 && (typeof inventory.emptyReason !== "string" || !inventory.emptyReason.trim())) {
      errors.push(`${inventoryContext}.emptyReason is required when the complete visual-asset inventory is empty.`);
    }
    for (const [itemIndex, item] of inventory.items.entries()) {
      const itemContext = `${inventoryContext}.items[${itemIndex}]`;
      if (!item || typeof item !== "object" || Array.isArray(item)) {
        errors.push(`${itemContext} must be an object.`);
        continue;
      }
      const assetId = typeof item.assetId === "string" ? item.assetId.trim() : "";
      if (!assetId || !SEMANTIC_ID.test(assetId)) errors.push(`${itemContext}.assetId must be a unique lowercase ASCII identifier.`);
      else if (compositeVisualAssetIds.has(assetId)) errors.push(`${itemContext}.assetId duplicates ${JSON.stringify(assetId)}.`);
      else compositeVisualAssetIds.add(assetId);
      let itemBox;
      try {
        itemBox = bboxToPixels(normalizeBbox(item.bbox, `${itemContext}.bbox`, { requireUnit: true }), sourceSlide.dimensions, `${itemContext}.bbox`);
        const inside = bboxIntersection(region.sourceBbox, itemBox).area;
        if (inside / Math.max(1, itemBox.width * itemBox.height) < 0.995) errors.push(`${itemContext}.bbox must be fully contained by its parent reconstruction region.`);
      } catch (error) { errors.push(error.message); }
      if (!ASSET_CLASSES.has(item.assetClass)) errors.push(`${itemContext}.assetClass must be one of: ${[...ASSET_CLASSES].join(", ")}.`);
      if (!COMPOSITE_VISUAL_DISPOSITIONS.has(item.disposition)) errors.push(`${itemContext}.disposition must be one of: ${[...COMPOSITE_VISUAL_DISPOSITIONS].join(", ")}.`);
      if (typeof item.reason !== "string" || !item.reason.trim()) errors.push(`${itemContext}.reason must be non-empty.`);
      const itemIds = Array.isArray(item.sourceObjectIds) ? item.sourceObjectIds.map(String) : [];
      if (!sameStringSet(itemIds, region.sourceObjectIds)) errors.push(`${itemContext}.sourceObjectIds must exactly repeat the parent composite source object IDs.`);
      if (item.assetClass === "generic-icon" && !["extract-raster", "regenerate-icon"].includes(item.disposition)) errors.push(`${itemContext} generic-icon assets from a composite raster must use extract-raster or regenerate-icon; Unicode glyph substitution is not a valid asset disposition.`);
      if (["logo", "photo", "product-ui", "evidence-screenshot", "official-diagram", "complex-illustration"].includes(item.assetClass) && !["retain-raster", "extract-raster", "manual-review"].includes(item.disposition)) errors.push(`${itemContext} authenticity-sensitive assets must be retained, extracted, or sent to manual-review; do not regenerate or replace them natively.`);
      if (["retain-raster", "extract-raster", "regenerate-icon"].includes(item.disposition) && !assetRecords.some((asset) => asset?.assetId === assetId && asset.action === item.disposition)) errors.push(`${itemContext} requires a matching plan.assets record with the same assetId and action.`);
      if (item.disposition === "rebuild-native" && item.assetClass !== "other") errors.push(`${itemContext}.rebuild-native is allowed only for a genuinely simple non-icon visual asset classified as other.`);
      if (itemBox) item.__bbox = itemBox;
    }
  }

  const allPlannedOutputRegions = normalizedRegions.flatMap((region) => [region, ...region.semanticSubregions]);
  for (const region of allPlannedOutputRegions.filter((entry) => RASTER_ASSET_ACTIONS.has(entry.action))) {
    const matchingAssets = assetRecords.filter((asset) => {
      if (!asset || typeof asset !== "object") return false;
      if (Number(asset.slideNumber) !== region.slideNumber || asset.action !== region.action) return false;
      if (region.assetId && asset.assetId === region.assetId) return true;
      if (region.regionId && asset.regionId === region.regionId) return true;
      const ids = Array.isArray(asset.sourceObjectIds) ? asset.sourceObjectIds.map(String) : [];
      return ids.some((id) => region.sourceObjectIds.includes(id));
    });
    if (matchingAssets.length === 0) {
      errors.push(`${region.__source} (${region.action}) requires a matching plan.assets provenance record.`);
      continue;
    }
    for (const objectId of region.sourceObjectIds) {
      if (!matchingAssets.some((asset) => Array.isArray(asset.sourceObjectIds) && asset.sourceObjectIds.map(String).includes(objectId))) {
        errors.push(`${region.__source} has no matching asset record for source object ${JSON.stringify(objectId)}.`);
      }
    }
  }

  for (const slideNumber of [...slidesByNumber.keys()].sort((a, b) => a - b)) {
    if (!representedSlides.has(slideNumber)) {
      errors.push(
        `Slide ${slideNumber} is unclassified in the conversion plan; add a plan slide entry or a keep-native/rebuild region.`,
      );
    }
  }

  const coverageByObjectId = new Map();
  for (const region of normalizedRegions) {
    for (const objectId of region.sourceObjectIds) {
      const list = coverageByObjectId.get(objectId) ?? [];
      list.push(region);
      coverageByObjectId.set(objectId, list);
    }
  }
  for (const objectId of [...largeRasterIds].sort()) {
    const dispositions = coverageByObjectId.get(objectId) ?? [];
    if (!dispositions.length) {
      errors.push(`Large raster ${JSON.stringify(objectId)} is unclassified in the conversion plan.`);
      continue;
    }
    if (dispositions.every((region) => region.action === "keep-native")) {
      errors.push(
        `Large raster ${JSON.stringify(objectId)} uses keep-native only; use retain-raster with a reason or choose a reconstruction action.`,
      );
    }
    if (dispositions.length > 1) {
      errors.push(
        `Large raster ${JSON.stringify(objectId)} has ${dispositions.length} plan dispositions; `
          + "model it as one bounded region with one disposition.",
      );
    }
  }

  for (const [objectId, dispositions] of coverageByObjectId.entries()) {
    const destructive = dispositions.filter((region) => DESTRUCTIVE_ACTIONS.has(region.action));
    if (destructive.length > 1) {
      errors.push(
        `Source object ${JSON.stringify(objectId)} is consumed by multiple destructive actions: `
          + destructive.map((region) => `${region.action} (${region.__source})`).join(", ") + ".",
      );
    }
  }

  for (const region of normalizedRegions.filter((entry) => DESTRUCTIVE_ACTIONS.has(entry.action))) {
    const sourceBoxes = [];
    for (const objectId of region.sourceObjectIds) {
      const sourceBox = objectBoxesById.get(objectId);
      if (!sourceBox) {
        errors.push(
          `${region.__source} cannot authorize deletion of ${JSON.stringify(objectId)} `
            + "because the source manifest has no valid, unit-declared bbox for that object.",
        );
        continue;
      }
      const sourceCoverage = bboxMetrics(region.sourceBbox, sourceBox).sourceCoverage;
      if (sourceCoverage < MIN_DESTRUCTIVE_OBJECT_COVERAGE) {
        errors.push(
          `${region.__source}.bbox covers only ${(sourceCoverage * 100).toFixed(2)}% of bound source object `
            + `${JSON.stringify(objectId)}; destructive actions require at least `
            + `${(MIN_DESTRUCTIVE_OBJECT_COVERAGE * 100).toFixed(0)}% coverage.`,
        );
      }
      sourceBoxes.push(sourceBox);
    }
    if (!sourceBoxes.length) continue;
    const footprint = bboxUnion(sourceBoxes);
    const metrics = bboxMetrics(region.sourceBbox, footprint);
    const safeFootprint = metrics.sourceCoverage >= MIN_DESTRUCTIVE_OBJECT_COVERAGE
      && (
        metrics.iou >= MIN_DESTRUCTIVE_IOU
        || (
          bboxContains(region.sourceBbox, footprint)
          && metrics.expansion >= 1
          && metrics.expansion <= MAX_DESTRUCTIVE_REGION_EXPANSION
        )
      );
    if (!safeFootprint) {
      errors.push(
        `${region.__source}.bbox is not a safe replacement footprint for its destructive source objects `
          + `(sourceCoverage=${metrics.sourceCoverage.toFixed(3)}, IoU=${metrics.iou.toFixed(3)}, `
          + `areaRatio=${metrics.expansion.toFixed(3)}). Require >=${MIN_DESTRUCTIVE_OBJECT_COVERAGE} source coverage `
          + `and either IoU >=${MIN_DESTRUCTIVE_IOU} or a containing region no larger than `
          + `${MAX_DESTRUCTIVE_REGION_EXPANSION}x the source footprint.`,
      );
    }
    region.destructiveFootprint = {
      bbox: footprint,
      sourceCoverage: Number(metrics.sourceCoverage.toFixed(6)),
      regionCoverage: Number(metrics.regionCoverage.toFixed(6)),
      iou: Number(metrics.iou.toFixed(6)),
      areaRatio: Number(metrics.expansion.toFixed(6)),
    };
  }

  const regionsBySlide = new Map();
  for (const region of normalizedRegions) {
    const list = regionsBySlide.get(region.slideNumber) ?? [];
    list.push(region);
    regionsBySlide.set(region.slideNumber, list);
  }
  for (const [slideNumber, regions] of regionsBySlide.entries()) {
    for (let leftIndex = 0; leftIndex < regions.length; leftIndex += 1) {
      for (let rightIndex = leftIndex + 1; rightIndex < regions.length; rightIndex += 1) {
        const first = regions[leftIndex];
        const second = regions[rightIndex];
        if (first.allowOverlap === true || second.allowOverlap === true) continue;
        if (
          typeof first.overlapGroup === "string"
          && first.overlapGroup
          && first.overlapGroup === second.overlapGroup
        ) continue;
        const overlap = bboxIntersection(first.bbox, second.bbox);
        if (overlap.area <= 0) continue;
        const smallerArea = Math.min(
          first.bbox.width * first.bbox.height,
          second.bbox.width * second.bbox.height,
        );
        if (overlap.area > 1 && overlap.area / smallerArea > 0.005) {
          errors.push(
            `Plan regions overlap on slide ${slideNumber}: ${first.__source} and ${second.__source}. `
              + "Set allowOverlap:true only when the overlap is intentional.",
          );
        }
      }
    }
  }

  report.statistics = {
    stage,
    slideCount: slidesByNumber.size,
    manifestObjectCount: objectsById.size,
    largeRasterCount: largeRasterIds.size,
    planRegionCount: normalizedRegions.length,
    semanticSubregionCount: normalizedRegions.reduce((sum, region) => sum + region.semanticSubregions.length, 0),
    representedSlideCount: representedSlides.size,
    boundSourceObjectCount: coverageByObjectId.size,
    activeSlideCount: stage === "calibration" ? calibrationSlideNumbers.size : slidesByNumber.size,
  };
  report.valid = errors.length === 0;

  if (!report.valid) {
    await removeStaleMap(outMapPath);
    if (reportPath) await writeJson(reportPath, report);
    return { report, frameMap: null };
  }

  const manifestHash = sha256(manifestBytes);
  const planHash = sha256(planBytes);
  const planSlidesByNumber = new Map(
    (Array.isArray(plan.slides) ? plan.slides : [])
      .map((entry) => [asPositiveInteger(entry?.slideNumber ?? entry?.number ?? entry?.index), entry])
      .filter(([number]) => number != null),
  );
  const frameIdByRegion = new Map();
  const frameMapSlides = [...slidesByNumber.entries()]
    .sort(([a], [b]) => a - b)
    .map(([slideNumber, { slide, dimensions }]) => {
      const calibrationActive = stage !== "calibration" || calibrationSlideNumbers.has(slideNumber);
      const sourceRegions = calibrationActive ? (regionsBySlide.get(slideNumber) ?? []) : [];
      const regions = sourceRegions.map((region, index) => {
        const frameId = typeof region.regionId === "string" && region.regionId.trim()
          ? region.regionId.trim()
          : `slide-${String(slideNumber).padStart(3, "0")}-frame-${String(index + 1).padStart(3, "0")}`;
        frameIdByRegion.set(region, frameId);
        const semanticSubregions = region.semanticSubregions.map((subregion, subIndex) => {
          const subregionFrameId = subregion.regionId || `${frameId}-sub-${subIndex + 1}`;
          frameIdByRegion.set(subregion, subregionFrameId);
          return {
            frameId: subregionFrameId,
            parentFrameId: frameId,
            sourceConsumption: "semantic-child",
            bbox: subregion.bbox,
            sourceObjectIds: [...subregion.sourceObjectIds],
            action: subregion.action,
            targetType: subregion.targetType,
            targetTypeComponents: [...subregion.targetTypeComponents],
            expectedText: subregion.expectedText,
            confidence: subregion.confidence,
            reason: subregion.reason,
            editability: [...subregion.editability],
            ...(subregion.intent ? { intent: subregion.intent } : {}),
            ...(subregion.styleRole ? { styleRole: subregion.styleRole } : {}),
            ...(subregion.componentFamily ? { componentFamily: subregion.componentFamily } : {}),
            ...(Array.isArray(subregion.issueRefs) ? { issueRefs: [...subregion.issueRefs] } : {}),
            ...(subregion.allowOverlap === true ? { allowOverlap: true } : {}),
            ...(subregion.overlapGroup ? { overlapGroup: subregion.overlapGroup } : {}),
          };
        });
        return {
          frameId,
          bbox: region.bbox,
          sourceBbox: region.sourceBbox,
          ...(region.targetBbox != null ? { targetBbox: region.bbox } : {}),
          sourceObjectIds: [...region.sourceObjectIds],
          sourceObjects: region.sourceObjectIds.map((objectId) => {
            const object = objectsById.get(objectId);
            return {
              objectId,
              kind: object.kind ?? null,
              name: object.name ?? null,
              bbox: objectBoxesById.get(objectId) ?? null,
              mediaRef: object.mediaRef ?? null,
              mediaSha256: object.mediaSha256 ?? object.sha256 ?? null,
            };
          }),
          action: region.action,
          targetType: region.targetType,
          targetTypeComponents: [...region.targetTypeComponents],
          expectedText: region.expectedText,
          confidence: region.confidence,
          reason: region.reason,
          editability: [...region.editability],
          ...(region.intent ? { intent: region.intent } : {}),
          ...(region.styleRole ? { styleRole: region.styleRole } : {}),
          ...(region.componentFamily ? { componentFamily: region.componentFamily } : {}),
          ...(Array.isArray(region.issueRefs) ? { issueRefs: [...region.issueRefs] } : {}),
          ...(region.intent === "user-approved-redesign" ? { authorization: region.authorization, preserveText: true } : {}),
          sourceConsumption: "owner",
          semanticSubregions,
          ...(region.destructiveFootprint
            ? { destructiveFootprint: region.destructiveFootprint }
            : {}),
          ...(region.allowOverlap === true ? { allowOverlap: true } : {}),
          ...(region.overlapGroup ? { overlapGroup: region.overlapGroup } : {}),
        };
      });
      const planSlide = planSlidesByNumber.get(slideNumber) ?? {};
      const deleteObjectIds = [...new Set(
        regions
          .filter((region) => DESTRUCTIVE_ACTIONS.has(region.action))
          .flatMap((region) => region.sourceObjectIds),
      )].sort();
      const allSlideObjectIds = [...objectsById.values()]
        .filter((object) => Number(object.slideNumber) === slideNumber)
        .map((object) => object.objectId)
        .sort();
      const deleteSet = new Set(deleteObjectIds);
      const addZones = regions
        .filter((region) => DESTRUCTIVE_ACTIONS.has(region.action))
        .map((region) => ({
          regionId: region.frameId,
          bbox: region.bbox,
          allowedTargetTypes: [...new Set([
            ...region.targetTypeComponents,
            ...region.semanticSubregions.flatMap((subregion) => subregion.targetTypeComponents),
          ])],
          ...(region.semanticSubregions.length > 0
            ? { semanticSubregions: region.semanticSubregions }
            : {}),
        }));
      return {
        slideNumber,
        sourceSlide: asPositiveInteger(planSlide.sourceSlide) ?? slideNumber,
        slideId: slide.slideId ?? null,
        classification: slide.classification,
        dimensions,
        deleteObjectIds,
        preserveObjectIds: allSlideObjectIds.filter((objectId) => !deleteSet.has(objectId)),
        addZones,
        regions,
        calibrationActive,
      };
    });

  // Emit the native $presentations template-following view as well as the
  // conversion-specific view above. This lets prepare_template_starter_deck
  // duplicate the source slides without translating IDs or weakening the
  // exact-object deletion contract.
  const outputSlides = frameMapSlides.map((slideEntry) => {
    if (stage === "calibration" && !slideEntry.calibrationActive) {
      return {
        outputSlide: slideEntry.slideNumber,
        sourceSlide: slideEntry.sourceSlide,
        narrativeRole: "source preservation outside calibration sample",
        reuseMode: "duplicate-slide",
        editTargets: slideEntry.preserveObjectIds.map((sourceElementId) => ({
          action: objectsById.get(sourceElementId)?.isPlaceholder ? "rewrite" : "keep",
          sourceElementId,
          reason: "Preserve this source object while calibrating representative slides.",
        })),
      };
    }
    const editTargets = [];
    const regionObjectIds = new Set();
    for (const region of slideEntry.regions) {
      for (const objectId of region.sourceObjectIds) regionObjectIds.add(objectId);
      if (DESTRUCTIVE_ACTIONS.has(region.action)) {
        const allowedTargetTypes = [...new Set([
          ...region.targetTypeComponents,
          ...region.semanticSubregions.flatMap((subregion) => subregion.targetTypeComponents),
        ])];
        editTargets.push({
          action: "delete",
          sourceElementIds: [...region.sourceObjectIds],
          regionId: region.frameId,
          reason: region.reason || `Remove exact source object(s) before ${region.action}.`,
        });
        editTargets.push({
          action: "add",
          newPrimitiveAllowed: true,
          mustNotOverlapInherited: true,
          zone: { ...region.bbox },
          regionId: region.frameId,
          targetType: allowedTargetTypes.join("+"),
          ...(region.semanticSubregions.length > 0
            ? { semanticSubregions: region.semanticSubregions }
            : {}),
          reason: region.reason || `Add planned ${region.targetType} content inside the deleted composite-image zone.`,
        });
        continue;
      }
      for (const objectId of region.sourceObjectIds) {
        const object = objectsById.get(objectId);
        editTargets.push({
          action: region.targetBbox != null ? "rewrite-and-reposition" : object?.isPlaceholder ? "rewrite" : "keep",
          sourceElementId: objectId,
          regionId: region.frameId,
          ...(region.targetBbox != null ? { targetBbox: region.bbox } : {}),
          reason: region.reason || `Preserve the exact source object for ${region.action}.`,
        });
      }
    }

    // The upstream template validator requires every inherited placeholder to
    // be explicitly handled. Preserve non-empty source placeholders through a
    // same-content rewrite declaration and delete empty ones so PowerPoint does
    // not reveal authoring prompts even when PNG renders look clean.
    for (const objectId of slideEntry.preserveObjectIds) {
      if (regionObjectIds.has(objectId)) continue;
      const object = objectsById.get(objectId);
      if (!object?.isPlaceholder) continue;
      const sourceText = String(object.text ?? object.textPreview ?? "").trim();
      editTargets.push({
        action: sourceText ? "rewrite" : "delete",
        sourceElementId: objectId,
        reason: sourceText
          ? "Preserve the verified source placeholder content while rebuilding slide-local body regions."
          : "Delete an empty structural placeholder to prevent an unresolved PowerPoint authoring prompt.",
      });
    }

    return {
      outputSlide: slideEntry.slideNumber,
      sourceSlide: slideEntry.sourceSlide,
      narrativeRole: slideEntry.regions.length > 0 ? "editable content reconstruction" : "source preservation",
      reuseMode: "duplicate-slide",
      editTargets,
    };
  });

  const frameMap = {
    schemaVersion: String(plan.schemaVersion ?? "1.0"),
    stage,
    generatedAt: manifest.generatedAt ?? null,
    source: {
      manifestPath: path.resolve(manifestPath),
      manifestSha256: manifestHash,
      inputPath: manifest.source?.inputPath ?? manifest.sourcePptx ?? null,
      inputSha256: manifest.source?.sha256 ?? manifest.sourceSha256 ?? null,
      validatedInputPath: sourceIdentity.matchedSourcePath ?? null,
    },
    plan: {
      path: path.resolve(planPath),
      sha256: planHash,
      mode: plan.mode,
      stage,
      sourceManifest: path.resolve(manifestPath),
      sourcePptx: sourceIdentity.declaredSourcePath ?? null,
      outputPptx: sourceIdentity.outputPath ?? null,
      ...(visualContract.isV11 ? {
        visualPolicy: plan.visualPolicy,
        styleProfile: plan.styleProfile,
        calibration: plan.calibration,
        feedbackIssues: plan.feedbackIssues ?? [],
      } : {}),
    },
    allowedActions: [...ALLOWED_ACTIONS],
    deck: {
      slideCount: manifest.deck?.slideCount ?? manifest.slideCount ?? slidesByNumber.size,
      slideSize: manifest.deck?.slideSize ?? manifest.slideSize ?? null,
    },
    outputSlides,
    omittedSourceSlides: [],
    slides: frameMapSlides,
    sourceObjectFrames: [...coverageByObjectId.entries()]
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([objectId, regions]) => ({
        objectId,
        slideNumber: Number(objectsById.get(objectId).slideNumber),
        frameIds: regions.flatMap((region) => [
          frameIdByRegion.get(region),
          ...region.semanticSubregions
            .filter((subregion) => subregion.sourceObjectIds.includes(objectId))
            .map((subregion) => frameIdByRegion.get(subregion)),
        ]).filter(Boolean),
      }))
      .filter((entry) => entry.frameIds.length > 0),
  };

  await writeJson(outMapPath, frameMap);
  if (reportPath) await writeJson(reportPath, report);
  return { report, frameMap };
}

async function main() {
  let args;
  try {
    args = parseArgs(process.argv.slice(2));
  } catch (error) {
    console.error(`ERROR: ${error.message}\n\n${HELP}`);
    process.exitCode = 2;
    return;
  }
  if (args.help) {
    console.log(HELP);
    return;
  }
  const missingOptions = ["manifest", "plan", "out-map"].filter((key) => !args[key]);
  if (missingOptions.length) {
    console.error(`ERROR: Missing required option(s): ${missingOptions.map((key) => `--${key}`).join(", ")}\n\n${HELP}`);
    process.exitCode = 2;
    return;
  }

  try {
    const { report } = await validateConversionPlan({
      manifestPath: args.manifest,
      planPath: args.plan,
      outMapPath: args["out-map"],
      reportPath: args.report,
      stage: args.stage ?? "final",
    });
    if (!report.valid) {
      console.error(`Conversion plan is invalid (${report.errors.length} error(s), ${report.warnings.length} warning(s)).`);
      for (const error of report.errors) console.error(`  - ${error}`);
      if (args.report) console.error(`Validation report: ${path.resolve(args.report)}`);
      process.exitCode = 1;
      return;
    }
    console.log(`Conversion plan is valid: ${report.statistics.planRegionCount} region(s) across ${report.statistics.slideCount} slide(s).`);
    console.log(`Template frame map: ${path.resolve(args["out-map"])}`);
    if (args.report) console.log(`Validation report: ${path.resolve(args.report)}`);
  } catch (error) {
    console.error(`ERROR: ${error.stack ?? error.message}`);
    process.exitCode = 1;
  }
}

const invokedAsScript = process.argv[1]
  && path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url));
if (invokedAsScript) await main();
