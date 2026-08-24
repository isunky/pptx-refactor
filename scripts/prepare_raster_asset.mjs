#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";
import { parseArgs as parseNodeArgs } from "node:util";

const MODES = new Set(["auto", "chroma", "edge"]);
const DEFAULT_TOLERANCE = 28;
const ALPHA_VISIBLE = 8;

function usage() {
  return [
    "Usage:",
    "  $RUNTIME_NODE prepare_raster_asset.mjs --input <image> --output <png> --mode <auto|chroma|edge> --report <json> [options]",
    "",
    "Required:",
    "  --input <path>          Source raster image.",
    "  --output <path>         Prepared PNG output. Must differ from --input.",
    "  --mode <mode>           auto, chroma, or edge.",
    "  --report <path>         Machine-readable JSON report.",
    "",
    "Options:",
    "  --chroma <#RRGGBB>      Explicit background key. Required by chroma mode.",
    "  --tolerance <0-255>     Per-channel color tolerance (default 28).",
    "  --optical-size <0-1>    Put the trimmed asset on a square transparent canvas",
    "                          so its longest visible edge occupies this fraction.",
    "  --padding <px>          Transparent padding around trimmed content (default 0).",
    "  --help                  Show this help.",
    "",
    "Modes:",
    "  auto    Preserve useful source alpha; otherwise infer an edge background.",
    "  chroma  Remove the explicit chroma color throughout the image.",
    "  edge    Remove only matching background pixels connected to an outer edge.",
  ].join("\n");
}

function parseCli(argv) {
  return parseNodeArgs({
    args: argv,
    options: {
      help: { type: "boolean" },
      input: { type: "string" },
      output: { type: "string" },
      mode: { type: "string" },
      report: { type: "string" },
      chroma: { type: "string" },
      tolerance: { type: "string" },
      "optical-size": { type: "string" },
      padding: { type: "string" },
    },
    allowPositionals: false,
    strict: true,
  }).values;
}

function requireString(args, key) {
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`Missing required --${key}.`);
  }
  return value;
}

function finiteNumber(value, label, { min, max, integer = false } = {}) {
  const number = Number(value);
  if (!Number.isFinite(number) || (integer && !Number.isInteger(number))) {
    throw new Error(`${label} must be ${integer ? "an integer" : "a number"}.`);
  }
  if (min !== undefined && number < min) throw new Error(`${label} must be >= ${min}.`);
  if (max !== undefined && number > max) throw new Error(`${label} must be <= ${max}.`);
  return number;
}

function parseColor(value) {
  if (typeof value !== "string") return undefined;
  const normalized = value.trim();
  const hex = normalized.match(/^#?([0-9a-f]{6})$/i);
  if (hex) {
    return {
      r: Number.parseInt(hex[1].slice(0, 2), 16),
      g: Number.parseInt(hex[1].slice(2, 4), 16),
      b: Number.parseInt(hex[1].slice(4, 6), 16),
      hex: `#${hex[1].toUpperCase()}`,
    };
  }
  const rgb = normalized.match(/^\s*(\d{1,3})\s*[,/]\s*(\d{1,3})\s*[,/]\s*(\d{1,3})\s*$/);
  if (rgb) {
    const channels = rgb.slice(1).map(Number);
    if (channels.every((channel) => channel >= 0 && channel <= 255)) {
      const [r, g, b] = channels;
      return { r, g, b, hex: `#${[r, g, b].map((channel) => channel.toString(16).padStart(2, "0")).join("").toUpperCase()}` };
    }
  }
  throw new Error(`Invalid --chroma color "${value}". Use #RRGGBB or r,g,b.`);
}

function runtimeRequire() {
  const modulesPath = process.env.RUNTIME_NODE_MODULES;
  if (!modulesPath || !path.isAbsolute(modulesPath)) {
    throw new Error("RUNTIME_NODE_MODULES must be set to the absolute bundled Node modules path.");
  }
  return createRequire(path.join(modulesPath, "__make_pptx_editable_runtime__.cjs"));
}

function channelDistance(data, offset, color) {
  return Math.max(
    Math.abs(data[offset] - color.r),
    Math.abs(data[offset + 1] - color.g),
    Math.abs(data[offset + 2] - color.b),
  );
}

function median(values) {
  if (values.length === 0) return 0;
  values.sort((a, b) => a - b);
  return values[Math.floor(values.length / 2)];
}

function samplePatchColor(data, info, left, top, width, height) {
  const reds = [];
  const greens = [];
  const blues = [];
  for (let y = top; y < top + height; y += 1) {
    for (let x = left; x < left + width; x += 1) {
      const offset = (y * info.width + x) * info.channels;
      if (data[offset + 3] <= ALPHA_VISIBLE) continue;
      reds.push(data[offset]);
      greens.push(data[offset + 1]);
      blues.push(data[offset + 2]);
    }
  }
  if (reds.length === 0) return undefined;
  return { r: median(reds), g: median(greens), b: median(blues) };
}

function inferEdgeBackground(data, info, tolerance) {
  const patch = Math.max(1, Math.min(9, Math.floor(Math.min(info.width, info.height) / 12)));
  const samples = [
    samplePatchColor(data, info, 0, 0, patch, patch),
    samplePatchColor(data, info, info.width - patch, 0, patch, patch),
    samplePatchColor(data, info, 0, info.height - patch, patch, patch),
    samplePatchColor(data, info, info.width - patch, info.height - patch, patch, patch),
  ].filter(Boolean);
  if (samples.length < 2) return { color: undefined, confidence: 0, samples };

  let maxDistance = 0;
  for (let i = 0; i < samples.length; i += 1) {
    for (let j = i + 1; j < samples.length; j += 1) {
      maxDistance = Math.max(
        maxDistance,
        Math.max(
          Math.abs(samples[i].r - samples[j].r),
          Math.abs(samples[i].g - samples[j].g),
          Math.abs(samples[i].b - samples[j].b),
        ),
      );
    }
  }
  const allowedSpread = Math.max(12, Math.min(72, tolerance * 1.75));
  if (maxDistance > allowedSpread) return { color: undefined, confidence: 0, samples, maxDistance };
  const color = {
    r: median(samples.map((sample) => sample.r)),
    g: median(samples.map((sample) => sample.g)),
    b: median(samples.map((sample) => sample.b)),
  };
  return {
    color,
    confidence: Math.max(0, 1 - maxDistance / Math.max(1, allowedSpread)),
    samples,
    maxDistance,
  };
}

function clearAlreadyTransparent(data, info) {
  let transparent = 0;
  for (let index = 0; index < info.width * info.height; index += 1) {
    const alphaOffset = index * info.channels + 3;
    if (data[alphaOffset] <= ALPHA_VISIBLE) {
      data[alphaOffset] = 0;
      transparent += 1;
    }
  }
  return transparent;
}

function removeGlobalChroma(data, info, color, tolerance) {
  let removed = 0;
  for (let index = 0; index < info.width * info.height; index += 1) {
    const offset = index * info.channels;
    if (data[offset + 3] === 0) continue;
    if (channelDistance(data, offset, color) <= tolerance) {
      data[offset + 3] = 0;
      removed += 1;
    }
  }
  return removed;
}

function removeEdgeConnectedBackground(data, info, color, tolerance) {
  const pixelCount = info.width * info.height;
  const visited = new Uint8Array(pixelCount);
  const queue = new Int32Array(pixelCount);
  let head = 0;
  let tail = 0;
  let removed = 0;

  const enqueue = (x, y) => {
    if (x < 0 || y < 0 || x >= info.width || y >= info.height) return;
    const index = y * info.width + x;
    if (visited[index]) return;
    const offset = index * info.channels;
    if (data[offset + 3] === 0 || channelDistance(data, offset, color) > tolerance) return;
    visited[index] = 1;
    queue[tail] = index;
    tail += 1;
  };

  for (let x = 0; x < info.width; x += 1) {
    enqueue(x, 0);
    enqueue(x, info.height - 1);
  }
  for (let y = 1; y < info.height - 1; y += 1) {
    enqueue(0, y);
    enqueue(info.width - 1, y);
  }

  while (head < tail) {
    const index = queue[head];
    head += 1;
    const x = index % info.width;
    const y = Math.floor(index / info.width);
    data[index * info.channels + 3] = 0;
    removed += 1;
    enqueue(x - 1, y);
    enqueue(x + 1, y);
    enqueue(x, y - 1);
    enqueue(x, y + 1);
  }
  return removed;
}

function removeSmallEdgeComponents(data, info) {
  const pixelCount = info.width * info.height;
  const visited = new Uint8Array(pixelCount);
  const queue = new Int32Array(pixelCount);
  const component = new Int32Array(pixelCount);
  const maximumArea = Math.max(32, Math.floor(pixelCount * 0.006));
  let removedComponents = 0;
  let removedPixels = 0;

  const visitComponent = (start) => {
    if (visited[start] || data[start * info.channels + 3] <= ALPHA_VISIBLE) return;
    let head = 0;
    let tail = 0;
    let count = 0;
    let minX = info.width;
    let maxX = 0;
    let minY = info.height;
    let maxY = 0;
    visited[start] = 1;
    queue[tail] = start;
    tail += 1;
    while (head < tail) {
      const current = queue[head];
      head += 1;
      component[count] = current;
      count += 1;
      const x = current % info.width;
      const y = Math.floor(current / info.width);
      minX = Math.min(minX, x);
      maxX = Math.max(maxX, x);
      minY = Math.min(minY, y);
      maxY = Math.max(maxY, y);
      const neighbors = [
        x > 0 ? current - 1 : -1,
        x + 1 < info.width ? current + 1 : -1,
        y > 0 ? current - info.width : -1,
        y + 1 < info.height ? current + info.width : -1,
      ];
      for (const neighbor of neighbors) {
        if (neighbor < 0 || visited[neighbor] || data[neighbor * info.channels + 3] <= ALPHA_VISIBLE) continue;
        visited[neighbor] = 1;
        queue[tail] = neighbor;
        tail += 1;
      }
    }
    const componentWidth = maxX - minX + 1;
    const componentHeight = maxY - minY + 1;
    const seamLike = componentWidth <= 3 || componentHeight <= 3;
    if (count <= maximumArea || seamLike) {
      for (let index = 0; index < count; index += 1) data[component[index] * info.channels + 3] = 0;
      removedComponents += 1;
      removedPixels += count;
    }
  };

  for (let x = 0; x < info.width; x += 1) {
    visitComponent(x);
    visitComponent((info.height - 1) * info.width + x);
  }
  for (let y = 1; y < info.height - 1; y += 1) {
    visitComponent(y * info.width);
    visitComponent(y * info.width + info.width - 1);
  }
  return { removedComponents, removedPixels };
}

function alphaBounds(data, info) {
  let left = info.width;
  let top = info.height;
  let right = -1;
  let bottom = -1;
  let visiblePixels = 0;
  let transparentPixels = 0;
  let edgeVisiblePixels = 0;
  let visibleXSum = 0;
  let visibleYSum = 0;
  let opaquePixels = 0;
  let borderPixels = 0;
  let borderOpaquePixels = 0;
  const borderSamples = [];
  const pixelCount = info.width * info.height;

  for (let y = 0; y < info.height; y += 1) {
    for (let x = 0; x < info.width; x += 1) {
      const alpha = data[(y * info.width + x) * info.channels + 3];
      const isBorder = x === 0 || y === 0 || x === info.width - 1 || y === info.height - 1;
      if (isBorder) {
        borderPixels += 1;
        if (alpha >= 250) borderOpaquePixels += 1;
        if (alpha > ALPHA_VISIBLE) {
          const offset = (y * info.width + x) * info.channels;
          borderSamples.push([data[offset], data[offset + 1], data[offset + 2]]);
        }
      }
      if (alpha > ALPHA_VISIBLE) {
        visiblePixels += 1;
        visibleXSum += x;
        visibleYSum += y;
        if (alpha >= 250) opaquePixels += 1;
        left = Math.min(left, x);
        top = Math.min(top, y);
        right = Math.max(right, x);
        bottom = Math.max(bottom, y);
        if (x === 0 || y === 0 || x === info.width - 1 || y === info.height - 1) edgeVisiblePixels += 1;
      } else {
        transparentPixels += 1;
      }
    }
  }
  if (right < left || bottom < top) return undefined;
  const channelStdDev = [0, 1, 2].map((channel) => {
    if (!borderSamples.length) return 0;
    const mean = borderSamples.reduce((sum, sample) => sum + sample[channel], 0) / borderSamples.length;
    const variance = borderSamples.reduce((sum, sample) => sum + (sample[channel] - mean) ** 2, 0) / borderSamples.length;
    return Math.sqrt(variance);
  });
  const borderColorStdDev = channelStdDev.reduce((sum, value) => sum + value, 0) / channelStdDev.length;
  const opaqueTileScore = Math.min(1,
    (transparentPixels / pixelCount < 0.001 ? 0.5 : 0)
      + (borderOpaquePixels / Math.max(1, borderPixels) >= 0.98 ? 0.25 : 0)
      + (borderSamples.length > 0 && borderColorStdDev <= 6 ? 0.25 : 0));
  return {
    left,
    top,
    width: right - left + 1,
    height: bottom - top + 1,
    visiblePixels,
    transparentPixels,
    transparencyRatio: transparentPixels / pixelCount,
    edgeVisiblePixels,
    edgeVisibleRatio: edgeVisiblePixels / Math.max(1, info.width * 2 + info.height * 2 - 4),
    opaquePixelRatio: opaquePixels / pixelCount,
    alphaVisibleCentroid: {
      x: visibleXSum / visiblePixels,
      y: visibleYSum / visiblePixels,
      normalizedX: (visibleXSum / visiblePixels) / Math.max(1, info.width - 1),
      normalizedY: (visibleYSum / visiblePixels) / Math.max(1, info.height - 1),
    },
    borderOpaqueRatio: borderOpaquePixels / Math.max(1, borderPixels),
    borderColorStdDev,
    opaqueTileScore,
  };
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

async function main() {
  const args = parseCli(process.argv.slice(2));
  if (args.help) {
    console.log(usage());
    return;
  }

  const input = path.resolve(requireString(args, "input"));
  const output = path.resolve(requireString(args, "output"));
  const reportPath = path.resolve(requireString(args, "report"));
  const mode = requireString(args, "mode").toLowerCase();
  if (!MODES.has(mode)) throw new Error(`--mode must be one of: ${[...MODES].join(", ")}.`);
  if (path.extname(output).toLowerCase() !== ".png") throw new Error("--output must use a .png extension.");
  if (input === output) throw new Error("--output must differ from --input; source assets are never overwritten.");
  if (reportPath === output || reportPath === input) throw new Error("--report must differ from the input and output paths.");

  const inputStat = await fs.stat(input).catch(() => undefined);
  if (!inputStat?.isFile()) throw new Error(`Input image does not exist: ${input}`);
  const tolerance = args.tolerance === undefined
    ? DEFAULT_TOLERANCE
    : finiteNumber(args.tolerance, "--tolerance", { min: 0, max: 255 });
  const padding = args.padding === undefined
    ? 0
    : finiteNumber(args.padding, "--padding", { min: 0, max: 4096, integer: true });
  const opticalSize = args["optical-size"] === undefined
    ? undefined
    : finiteNumber(args["optical-size"], "--optical-size", { min: Number.EPSILON, max: 1 });
  const explicitChroma = args.chroma === undefined ? undefined : parseColor(args.chroma);
  if (mode === "chroma" && !explicitChroma) throw new Error("--chroma is required when --mode chroma is used.");

  const requireFromRuntime = runtimeRequire();
  const sharpModule = requireFromRuntime("sharp");
  const sharp = sharpModule.default ?? sharpModule;
  const decoded = await sharp(input, { failOn: "error", limitInputPixels: 100_000_000 })
    .ensureAlpha()
    .raw()
    .toBuffer({ resolveWithObject: true });
  const data = Buffer.from(decoded.data);
  const info = decoded.info;
  if (info.channels !== 4) throw new Error(`Expected RGBA pixels after decoding; got ${info.channels} channels.`);

  const pixelCount = info.width * info.height;
  const transparentBefore = clearAlreadyTransparent(data, info);
  const inferred = explicitChroma
    ? { color: explicitChroma, confidence: 1, samples: [], maxDistance: 0 }
    : inferEdgeBackground(data, info, tolerance);
  let backgroundColor;
  let backgroundRemoval = 0;
  let cleanupMethod = "alpha-only";

  if (mode === "chroma") {
    backgroundColor = explicitChroma;
    backgroundRemoval = removeGlobalChroma(data, info, backgroundColor, tolerance);
    cleanupMethod = "global-chroma";
  } else if (mode === "edge") {
    backgroundColor = explicitChroma ?? inferred.color;
    if (!backgroundColor) throw new Error("Could not infer a consistent edge background; provide --chroma or use --mode auto.");
    backgroundRemoval = removeEdgeConnectedBackground(data, info, backgroundColor, tolerance);
    cleanupMethod = "edge-connected";
  } else {
    const usefulAlpha = transparentBefore / pixelCount >= 0.01;
    backgroundColor = explicitChroma ?? (!usefulAlpha ? inferred.color : undefined);
    if (explicitChroma) {
      backgroundRemoval = removeGlobalChroma(data, info, explicitChroma, tolerance);
      cleanupMethod = "global-chroma";
    } else if (backgroundColor) {
      backgroundRemoval = removeEdgeConnectedBackground(data, info, backgroundColor, tolerance);
      cleanupMethod = "edge-connected-auto";
    }
  }

  const edgeCleanup = removeSmallEdgeComponents(data, info);
  const cleanedBounds = alphaBounds(data, info);
  if (!cleanedBounds) throw new Error("Raster cleanup removed every visible pixel; reduce --tolerance or choose another mode.");
  if (mode === "chroma" && backgroundRemoval === 0) {
    throw new Error("The chroma color matched no pixels; verify --chroma and --tolerance.");
  }

  let pipeline = sharp(data, {
    raw: { width: info.width, height: info.height, channels: info.channels },
  }).extract({
    left: cleanedBounds.left,
    top: cleanedBounds.top,
    width: cleanedBounds.width,
    height: cleanedBounds.height,
  });

  if (padding > 0) {
    pipeline = pipeline.extend({
      top: padding,
      bottom: padding,
      left: padding,
      right: padding,
      background: { r: 0, g: 0, b: 0, alpha: 0 },
    });
  }

  let paddedWidth = cleanedBounds.width + padding * 2;
  let paddedHeight = cleanedBounds.height + padding * 2;
  if (opticalSize !== undefined) {
    const side = Math.max(1, Math.ceil(Math.max(paddedWidth, paddedHeight) / opticalSize));
    const left = Math.floor((side - paddedWidth) / 2);
    const right = side - paddedWidth - left;
    const top = Math.floor((side - paddedHeight) / 2);
    const bottom = side - paddedHeight - top;
    pipeline = pipeline.extend({
      top,
      bottom,
      left,
      right,
      background: { r: 0, g: 0, b: 0, alpha: 0 },
    });
    paddedWidth = side;
    paddedHeight = side;
  }

  const outputBuffer = await pipeline.png({ compressionLevel: 9, adaptiveFiltering: true }).toBuffer();
  const finalDecoded = await sharp(outputBuffer).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
  const finalBounds = alphaBounds(finalDecoded.data, finalDecoded.info);
  if (!finalBounds) throw new Error("Prepared output has no visible pixels.");

  const warnings = [];
  if (!backgroundColor && mode === "auto" && transparentBefore / pixelCount < 0.01) {
    warnings.push("No consistent edge background was detected; only alpha and small edge artifacts were cleaned.");
  }
  if (finalBounds.edgeVisibleRatio > 0.02) {
    warnings.push("Visible pixels still touch the output edge; use --padding or --optical-size for safer slide placement.");
  }
  if (finalBounds.transparencyRatio < 0.001) {
    warnings.push("Prepared output is effectively opaque; verify that this asset is meant to retain a rectangular background.");
  }
  if (finalBounds.opaqueTileScore >= 0.75) {
    warnings.push("Output has strong opaque-tile evidence; do not use it as a transparent generic icon.");
  }
  const opticalCoverage = Math.max(finalBounds.width, finalBounds.height) / Math.max(finalDecoded.info.width, finalDecoded.info.height);
  if (opticalSize !== undefined && Math.abs(opticalCoverage - opticalSize) > 0.02) {
    warnings.push(`Requested optical size ${opticalSize} produced ${opticalCoverage.toFixed(4)} after integer rounding.`);
  }

  await atomicWrite(output, outputBuffer);
  const outputStat = await fs.stat(output);
  const report = {
    schema: "pptx-refactor/raster-asset-report/v1.1",
    status: warnings.length > 0 ? "warning" : "pass",
    generatedAt: new Date().toISOString(),
    input,
    output,
    report: reportPath,
    mode,
    options: {
      chroma: explicitChroma?.hex,
      tolerance,
      opticalSize,
      padding,
    },
    source: {
      bytes: inputStat.size,
      width: info.width,
      height: info.height,
      channels: info.channels,
      transparentPixelRatio: transparentBefore / pixelCount,
    },
    cleanup: {
      method: cleanupMethod,
      backgroundColor: backgroundColor
        ? `#${[backgroundColor.r, backgroundColor.g, backgroundColor.b].map((channel) => channel.toString(16).padStart(2, "0")).join("").toUpperCase()}`
        : undefined,
      backgroundInferenceConfidence: backgroundColor && !explicitChroma ? inferred.confidence : undefined,
      backgroundPixelsRemoved: backgroundRemoval,
      edgeComponentsRemoved: edgeCleanup.removedComponents,
      edgeComponentPixelsRemoved: edgeCleanup.removedPixels,
      sourceVisibleBounds: {
        left: cleanedBounds.left,
        top: cleanedBounds.top,
        width: cleanedBounds.width,
        height: cleanedBounds.height,
      },
    },
    result: {
      bytes: outputStat.size,
      width: finalDecoded.info.width,
      height: finalDecoded.info.height,
      visibleBounds: {
        left: finalBounds.left,
        top: finalBounds.top,
        width: finalBounds.width,
        height: finalBounds.height,
      },
      transparentPixelRatio: finalBounds.transparencyRatio,
      edgeVisiblePixelRatio: finalBounds.edgeVisibleRatio,
      opaquePixelRatio: finalBounds.opaquePixelRatio,
      alphaVisibleCentroid: finalBounds.alphaVisibleCentroid,
      borderOpaqueRatio: finalBounds.borderOpaqueRatio,
      borderColorStdDev: finalBounds.borderColorStdDev,
      opaqueTileScore: finalBounds.opaqueTileScore,
      opticalCoverage,
    },
    warnings,
  };
  await atomicWrite(reportPath, Buffer.from(`${JSON.stringify(report, null, 2)}\n`, "utf8"));
  console.log(JSON.stringify(report, null, 2));
}

main().catch((error) => {
  console.error(`prepare_raster_asset: ${error.message || String(error)}`);
  console.error(usage());
  process.exitCode = 1;
});
