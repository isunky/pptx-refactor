import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";

const require = createRequire(import.meta.url);
let cachedSharp = null;
const SEMANTIC_ID = /^[a-z0-9][a-z0-9_-]*$/;

function normalizeSemanticId(value, label) {
  if (value == null || value === "") return undefined;
  const text = String(value);
  if (!SEMANTIC_ID.test(text)) {
    throw new TypeError(`${label} must use lowercase ASCII letters, digits, _ or - and start with a letter or digit.`);
  }
  return text;
}

export function semanticName(metadata = {}) {
  if (!metadata || typeof metadata !== "object" || Array.isArray(metadata)) {
    throw new TypeError("semanticName metadata must be an object.");
  }
  const ordered = [
    ["role", normalizeSemanticId(metadata.role, "semantic role")],
    ["family", normalizeSemanticId(metadata.family, "semantic family")],
    ["instance", normalizeSemanticId(metadata.instance, "semantic instance")],
    ["part", normalizeSemanticId(metadata.part, "semantic part")],
    ["item", metadata.item == null ? undefined : normalizeSemanticId(String(metadata.item), "semantic item")],
  ].filter(([, value]) => value !== undefined);
  if (!ordered.length) throw new TypeError("semanticName requires at least one semantic field.");
  return `mppe|${ordered.map(([key, value]) => `${key}=${value}`).join("|")}`;
}

function resolvedObjectName(options = {}, semanticOverrides = {}) {
  if (options.name) return options.name;
  const hasSemantic = options.semantic && Object.values(options.semantic)
    .some((value) => value !== undefined && value !== null && value !== "");
  if (!hasSemantic) return undefined;
  const semantic = { ...(options.semantic ?? {}), ...semanticOverrides };
  return Object.values(semantic).some((value) => value !== undefined && value !== null && value !== "")
    ? semanticName(semantic)
    : undefined;
}

export function ptToPx(points, dpi = 96) {
  const value = Number(points);
  const resolution = Number(dpi);
  if (!Number.isFinite(value) || !Number.isFinite(resolution) || resolution <= 0) {
    throw new TypeError("ptToPx(points, dpi) requires finite points and a positive dpi.");
  }
  return value * resolution / 72;
}

export function pxToPt(pixels, dpi = 96) {
  const value = Number(pixels);
  const resolution = Number(dpi);
  if (!Number.isFinite(value) || !Number.isFinite(resolution) || resolution <= 0) {
    throw new TypeError("pxToPt(pixels, dpi) requires finite pixels and a positive dpi.");
  }
  return value * 72 / resolution;
}

function compactObject(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return value;
  return Object.fromEntries(
    Object.entries(value)
      .filter(([, child]) => child !== undefined)
      .map(([key, child]) => [key, compactObject(child)]),
  );
}

function assertSlide(slide) {
  if (!slide?.shapes?.add) throw new TypeError("A slide facade with slide.shapes.add(...) is required.");
}

function assertPosition(position, name = "position") {
  if (!position || typeof position !== "object") throw new TypeError(`${name} must be an object.`);
  for (const key of ["left", "top", "width", "height"]) {
    if (!Number.isFinite(Number(position[key]))) throw new TypeError(`${name}.${key} must be finite.`);
  }
  if (Number(position.width) <= 0 || Number(position.height) <= 0) {
    throw new RangeError(`${name} width and height must be positive.`);
  }
  return {
    ...position,
    left: Number(position.left),
    top: Number(position.top),
    width: Number(position.width),
    height: Number(position.height),
  };
}

function getThemeValue(theme, paths) {
  for (const candidate of paths) {
    let value = theme;
    for (const key of candidate.split(".")) value = value?.[key];
    if (value !== undefined && value !== null) return value;
  }
  return undefined;
}

function themeFont(theme, role) {
  return getThemeValue(theme, [
    `typography.${role}.typeface`,
    `fonts.${role}`,
    role === "title" ? "fonts.heading" : "fonts.body",
    "fontFamily",
  ]);
}

function themeColor(theme, role) {
  return getThemeValue(theme, [`colors.${role}`, role]);
}

function makeLine(fill, width = 1, style = "solid") {
  if (fill == null || fill === "none" || Number(width) <= 0) {
    return { style: "solid", fill: "none", width: 0 };
  }
  return { style, fill, width: Number(width) };
}

export function addShape(slide, geometry, position, options = {}) {
  assertSlide(slide);
  if (typeof geometry !== "string" || !geometry) throw new TypeError("geometry must be a non-empty string.");
  const frame = assertPosition(position);
  return slide.shapes.add(compactObject({
    geometry,
    name: resolvedObjectName(options),
    position: frame,
    fill: options.fill ?? "none",
    line: options.line ?? makeLine(options.lineFill, options.lineWidth ?? 0, options.lineStyle),
    borderRadius: options.borderRadius,
    shadow: options.shadow,
    className: options.className,
    adjustmentList: options.adjustmentList,
    placeholderType: options.placeholderType,
    placeholderIndex: options.placeholderIndex,
  }));
}

export function addTextBox(slide, text, position, options = {}) {
  const theme = options.theme ?? {};
  const shape = addShape(slide, options.geometry ?? "textbox", position, {
    name: resolvedObjectName(options),
    fill: options.fill ?? "none",
    line: options.line ?? makeLine(options.lineFill, options.lineWidth ?? 0, options.lineStyle),
    borderRadius: options.borderRadius,
    shadow: options.shadow,
    className: options.className,
    placeholderType: options.placeholderType,
    placeholderIndex: options.placeholderIndex,
  });
  shape.text = text ?? "";
  const style = compactObject({
    styleName: options.styleName,
    className: options.textClassName,
    typeface: options.typeface ?? themeFont(theme, options.role ?? "body"),
    fontSize: options.fontSize,
    bold: options.bold,
    italic: options.italic,
    underline: options.underline,
    color: options.color ?? themeColor(theme, options.colorRole ?? "text"),
    alignment: options.alignment,
    verticalAlignment: options.verticalAlignment,
    lineSpacing: options.lineSpacing,
    autoFit: options.autoFit ?? "none",
    wrap: options.wrap ?? "square",
    insets: options.insets ?? { top: 0, right: 0, bottom: 0, left: 0 },
  });
  if (Object.keys(style).length) shape.text.style = style;
  return shape;
}

export function addCard(slide, position, options = {}) {
  const theme = options.theme ?? {};
  const border = options.border ?? themeColor(theme, "border");
  return addShape(slide, options.geometry ?? "roundRect", position, {
    name: resolvedObjectName(options, { part: options.semantic?.part ?? "frame" }),
    fill: options.fill ?? themeColor(theme, "surface") ?? "none",
    line: options.line ?? makeLine(border, options.lineWidth ?? (border ? 1 : 0), options.lineStyle),
    borderRadius: options.borderRadius,
    shadow: options.shadow,
    className: options.className,
  });
}

export function addTitle(slide, text, position, options = {}) {
  const theme = options.theme ?? {};
  return addTextBox(slide, text, position, {
    ...options,
    theme,
    role: "title",
    typeface: options.typeface ?? themeFont(theme, "title"),
    color: options.color ?? themeColor(theme, "title") ?? themeColor(theme, "text"),
    bold: options.bold ?? true,
    verticalAlignment: options.verticalAlignment ?? "middle",
  });
}

export function addRoleTextBox(slide, text, position, options = {}) {
  const role = normalizeSemanticId(options.role, "role");
  if (!role) throw new TypeError("addRoleTextBox requires options.role.");
  const theme = options.theme ?? {};
  const roleStyle = getThemeValue(theme, [`styleProfile.roles.${role}`, `roles.${role}`, `typography.${role}`]) ?? {};
  const semantic = {
    ...(options.semantic ?? {}),
    role,
    family: options.family ?? options.semantic?.family,
    instance: options.instance ?? options.semantic?.instance,
    part: options.part ?? options.semantic?.part ?? "text",
  };
  return addTextBox(slide, text, position, {
    ...roleStyle,
    ...options,
    role,
    semantic,
    name: options.name ?? semanticName(semantic),
    typeface: options.typeface ?? roleStyle.typeface ?? themeFont(theme, role),
    fontSize: options.fontSize ?? roleStyle.fontSize,
    bold: options.bold ?? roleStyle.bold,
    color: options.color ?? roleStyle.color ?? themeColor(theme, role) ?? themeColor(theme, "text"),
    alignment: options.alignment ?? roleStyle.alignment,
    verticalAlignment: options.verticalAlignment ?? roleStyle.verticalAlignment,
    lineSpacing: options.lineSpacing ?? roleStyle.lineSpacing,
    insets: options.insets ?? roleStyle.insets,
  });
}

export function addComponentCard(slide, position, options = {}) {
  const family = normalizeSemanticId(options.family, "component family");
  const instance = normalizeSemanticId(options.instance, "component instance");
  if (!family || !instance) throw new TypeError("addComponentCard requires options.family and options.instance.");
  return addCard(slide, position, {
    ...options,
    semantic: { ...(options.semantic ?? {}), family, instance, part: options.part ?? "frame" },
  });
}

export function addDivider(slide, position, options = {}) {
  assertSlide(slide);
  const theme = options.theme ?? {};
  const lineFill = options.color ?? themeColor(theme, "divider") ?? themeColor(theme, "border");
  if (lineFill == null) throw new Error("addDivider requires options.color or theme.colors.divider/border.");
  const normalized = {
    left: Number(position.left),
    top: Number(position.top),
    width: Number(position.width),
    height: Number(position.height ?? 0),
  };
  if (![normalized.left, normalized.top, normalized.width, normalized.height].every(Number.isFinite)) {
    throw new TypeError("Divider position values must be finite.");
  }
  return slide.shapes.add({
    geometry: "line",
    name: options.name,
    position: normalized,
    fill: "none",
    line: makeLine(lineFill, options.width ?? 1, options.style ?? "solid"),
  });
}

export function addConnector(slide, sourceShape, targetShape, options = {}) {
  assertSlide(slide);
  if (!slide.shapes.connect) throw new TypeError("This slide facade does not expose slide.shapes.connect(...).");
  if (!sourceShape || !targetShape) throw new TypeError("Both connector endpoint shapes/IDs are required.");
  const theme = options.theme ?? {};
  const lineFill = options.color ?? options.line?.fill ?? themeColor(theme, "connector") ?? themeColor(theme, "accent");
  if (lineFill == null) throw new Error("addConnector requires options.color/line.fill or a connector/accent theme color.");
  const connector = slide.shapes.connect(sourceShape, targetShape, compactObject({
    kind: options.kind ?? "elbow",
    fromSide: options.fromSide,
    toSide: options.toSide,
    fromIdx: options.fromIdx,
    toIdx: options.toIdx,
    line: options.line ?? makeLine(lineFill, options.width ?? 2, options.style ?? "solid"),
    head: options.head,
    tail: options.tail,
    cap: options.cap,
    join: options.join,
  }));
  // artifact-tool places new connected routes behind shapes. Keep this helper
  // side-effect free with respect to z-order unless the caller asks otherwise.
  if (options.bringToFront === true) connector.bringToFront();
  return connector;
}

export function estimateGlyphWidth(character, fontSize, options = {}) {
  const size = Number(fontSize);
  if (!(size > 0)) throw new RangeError("fontSize must be positive.");
  const char = String(character ?? "");
  if (!char) return 0;
  if (/\s/u.test(char)) return size * (options.spaceFactor ?? 0.32);
  if (/\p{Script=Han}|\p{Script=Hiragana}|\p{Script=Katakana}|\p{Script=Hangul}/u.test(char)) {
    return size * (options.cjkFactor ?? 1);
  }
  if (/\p{Number}/u.test(char)) return size * (options.digitFactor ?? 0.56);
  if (/\p{Punctuation}|\p{Symbol}/u.test(char)) return size * (options.punctuationFactor ?? 0.5);
  return size * (options.latinFactor ?? 0.56);
}

function tokenWidth(token, fontSize, options) {
  return [...token].reduce((sum, character) => sum + estimateGlyphWidth(character, fontSize, options), 0);
}

function tokenizeForWrap(line) {
  return line.match(/[\p{Script=Latin}\p{Number}][\p{Script=Latin}\p{Number}._+/%:@#&'’\-]*(?:\s+)?|\s+|./gu) ?? [];
}

export function wrapText(text, maxWidth, fontSize, options = {}) {
  const width = Number(maxWidth);
  const size = Number(fontSize);
  if (!(width > 0) || !(size > 0)) throw new RangeError("wrapText requires positive maxWidth and fontSize.");
  const output = [];
  for (const sourceLine of String(text ?? "").split(/\r?\n/u)) {
    if (sourceLine === "") {
      output.push("");
      continue;
    }
    let line = "";
    let currentWidth = 0;
    const flush = () => {
      output.push(line.trimEnd());
      line = "";
      currentWidth = 0;
    };
    for (let token of tokenizeForWrap(sourceLine)) {
      let widthOfToken = tokenWidth(token, size, options);
      if (line && currentWidth + widthOfToken > width) {
        flush();
        token = token.trimStart();
        widthOfToken = tokenWidth(token, size, options);
      }
      if (widthOfToken > width && token.length > 1 && options.breakLongTokens !== false) {
        for (const character of token) {
          const charWidth = estimateGlyphWidth(character, size, options);
          if (line && currentWidth + charWidth > width) flush();
          line += character;
          currentWidth += charWidth;
        }
      } else {
        line += token;
        currentWidth += widthOfToken;
      }
    }
    if (line || output.length === 0) flush();
  }
  return output.join("\n");
}

function normalizeBulletItem(item) {
  if (typeof item === "string" || typeof item === "number") {
    return { text: String(item), runs: [String(item)] };
  }
  if (!item || typeof item !== "object") throw new TypeError("Bullet items must be strings, numbers, or objects.");
  if (Array.isArray(item.runs)) {
    const plain = item.runs.map((run) => (typeof run === "object" ? run.run ?? "" : String(run))).join("");
    return { ...item, text: item.text ?? plain, runs: item.runs };
  }
  return { ...item, text: String(item.text ?? ""), runs: [String(item.text ?? "")] };
}

export function addStructuredBullets(slide, items, position, options = {}) {
  if (!Array.isArray(items) || items.length === 0) throw new TypeError("items must be a non-empty array.");
  const theme = options.theme ?? {};
  const frame = assertPosition(position);
  const fontSize = Number(options.fontSize ?? getThemeValue(theme, ["typography.body.fontSize"]));
  if (!(fontSize > 0)) throw new Error("addStructuredBullets requires options.fontSize or theme.typography.body.fontSize.");
  const marginLeft = Number(options.marginLeft ?? fontSize * 1.35);
  const indent = Number(options.indent ?? -fontSize * 0.7);
  const spaceAfter = Number(options.spaceAfter ?? fontSize * 0.5);
  const paragraphs = items.map((rawItem) => {
    const item = normalizeBulletItem(rawItem);
    return compactObject({
      bulletCharacter: options.bulletCharacter ?? "•",
      marginLeft: item.marginLeft ?? marginLeft,
      indent: item.indent ?? indent,
      spaceBefore: item.spaceBefore ?? options.spaceBefore,
      spaceAfter: item.spaceAfter ?? spaceAfter,
      paragraphStyle: item.paragraphStyle ?? options.paragraphStyle,
      styleId: item.styleId ?? options.styleId,
      runs: item.runs,
    });
  });
  const box = addShape(slide, "textbox", frame, {
    name: resolvedObjectName(options, { role: options.semantic?.role ?? "bullet-body", part: options.semantic?.part ?? "list" }),
    fill: options.fill ?? "none",
    line: options.line ?? makeLine(options.lineFill, options.lineWidth ?? 0, options.lineStyle),
  });
  box.text = paragraphs;
  box.text.style = compactObject({
    styleName: options.textStyleName,
    typeface: options.typeface ?? themeFont(theme, "body"),
    fontSize,
    bold: options.bold,
    color: options.color ?? themeColor(theme, "text"),
    alignment: options.alignment ?? "left",
    verticalAlignment: options.verticalAlignment ?? "top",
    lineSpacing: options.lineSpacing ?? 1.1,
    autoFit: options.autoFit ?? "none",
    wrap: options.wrap ?? "square",
    insets: options.insets ?? { top: 0, right: 0, bottom: 0, left: 0 },
  });
  return box;
}

export const addNativeBullets = addStructuredBullets;

export function addHangingBullets(slide, items, position, options = {}) {
  if (!Array.isArray(items) || items.length === 0) throw new TypeError("items must be a non-empty array.");
  const theme = options.theme ?? {};
  const frame = assertPosition(position);
  const fontSize = Number(options.fontSize ?? getThemeValue(theme, ["typography.body.fontSize"]));
  if (!(fontSize > 0)) throw new Error("addHangingBullets requires options.fontSize or theme.typography.body.fontSize.");
  const lineSpacing = Number(options.lineSpacing ?? 1.1);
  const dotSize = Number(options.dotSize ?? Math.max(3, fontSize * 0.42));
  const gap = Number(options.gap ?? fontSize * 0.65);
  const spaceAfter = Number(options.spaceAfter ?? fontSize * 0.5);
  const lineHeight = Number(options.lineHeight ?? fontSize * lineSpacing);
  const textLeft = frame.left + dotSize + gap;
  const textWidth = frame.width - dotSize - gap;
  if (!(textWidth > 0)) throw new RangeError("Bullet dot and gap leave no width for text.");
  const normalizedItems = items.map(normalizeBulletItem);
  const prepared = normalizedItems.map((item) => {
    const wrapped = wrapText(item.text, textWidth - Number(options.wrapSafety ?? 2), fontSize, options.wrapOptions);
    const lineCount = Math.max(1, wrapped.split("\n").length);
    const height = Math.max(lineHeight, Math.ceil(lineCount * lineHeight + Number(options.textHeightPadding ?? 2)));
    return { item, wrapped, height };
  });
  const totalHeight = prepared.reduce((sum, item) => sum + item.height, 0)
    + Math.max(0, prepared.length - 1) * spaceAfter;
  const plannedOverflow = Math.max(0, totalHeight - frame.height);
  if (plannedOverflow > 0 && options.allowOverflow !== true) {
    throw new RangeError(`Hanging bullet list exceeds its frame by ${plannedOverflow.toFixed(2)} px.`);
  }
  let cursorTop = frame.top;
  if (options.verticalAlignment === "middle") cursorTop += Math.max(0, (frame.height - totalHeight) / 2);
  if (options.verticalAlignment === "bottom") cursorTop += Math.max(0, frame.height - totalHeight);

  const bulletColor = options.bulletColor ?? themeColor(theme, "bullet") ?? themeColor(theme, "text");
  if (bulletColor == null) throw new Error("addHangingBullets requires options.bulletColor or theme.colors.bullet/text.");
  const semanticFamily = options.family ?? options.semantic?.family ?? "list";
  const semanticInstance = options.instance ?? options.semantic?.instance;
  const result = {
    bullets: [],
    textBoxes: [],
    elements: [],
    totalHeight,
    metrics: { fontSize, dotSize, gap, lineSpacing, lineHeight, textLeft, textWidth, spaceAfter },
  };
  prepared.forEach(({ item, wrapped, height }, index) => {
    const itemId = String(index + 1);
    const dotName = options.name
      ? `${options.name}-dot-${itemId}`
      : semanticInstance
        ? semanticName({ role: "bullet-body", family: semanticFamily, instance: semanticInstance, part: "dot", item: itemId })
        : undefined;
    const textName = options.name
      ? `${options.name}-text-${itemId}`
      : semanticInstance
        ? semanticName({ role: "bullet-body", family: semanticFamily, instance: semanticInstance, part: "text", item: itemId })
        : undefined;
    const dot = addShape(slide, "ellipse", {
      left: frame.left,
      top: cursorTop + Number(options.dotTopOffset ?? Math.max(0, (lineHeight - dotSize) / 2)),
      width: dotSize,
      height: dotSize,
    }, {
      name: dotName,
      fill: item.bulletColor ?? bulletColor,
      line: makeLine(null, 0),
    });
    const textBox = addTextBox(slide, wrapped, {
      left: textLeft,
      top: cursorTop,
      width: textWidth,
      height,
    }, {
      ...options,
      name: textName,
      theme,
      role: "body",
      fill: "none",
      line: makeLine(null, 0),
      fontSize,
      lineSpacing,
      verticalAlignment: "top",
      color: item.color ?? options.color ?? themeColor(theme, "text"),
      bold: item.bold ?? options.bold,
      insets: options.insets ?? { top: 0, right: 0, bottom: 0, left: 0 },
    });
    result.bullets.push(dot);
    result.textBoxes.push(textBox);
    result.elements.push(dot, textBox);
    cursorTop += height + (index === prepared.length - 1 ? 0 : spaceAfter);
  });
  result.bottom = cursorTop;
  result.overflow = Math.max(0, cursorTop - (frame.top + frame.height));
  return result;
}

function sharpSearchPaths() {
  const values = [
    process.env.RUNTIME_NODE_MODULES,
    process.env.NODE_PATH,
  ].filter(Boolean);
  return [...new Set(values.flatMap((value) => value.split(path.delimiter)).filter(Boolean))];
}

export async function loadSharp() {
  if (cachedSharp) return cachedSharp;
  try {
    const loaded = await import("sharp");
    cachedSharp = loaded.default ?? loaded;
    return cachedSharp;
  } catch (directError) {
    for (const searchPath of sharpSearchPaths()) {
      try {
        const resolved = require.resolve("sharp", { paths: [searchPath] });
        const loaded = require(resolved);
        cachedSharp = loaded.default ?? loaded;
        return cachedSharp;
      } catch {
        // Try the next configured runtime module path.
      }
    }
    throw new Error(
      `Unable to load sharp. Set RUNTIME_NODE_MODULES or NODE_PATH to a runtime containing sharp. `
        + `Original error: ${directError.message}`,
    );
  }
}

async function rasterInputBytes(input) {
  if (typeof input === "string" || input instanceof URL) return fs.readFile(input);
  if (Buffer.isBuffer(input)) return input;
  if (input instanceof Uint8Array) return Buffer.from(input.buffer, input.byteOffset, input.byteLength);
  if (input instanceof ArrayBuffer) return Buffer.from(input);
  throw new TypeError("Icon input must be a path, URL, Buffer, Uint8Array, or ArrayBuffer.");
}

export async function trimTransparentRaster(input, options = {}) {
  const sharp = await loadSharp();
  const bytes = await rasterInputBytes(input);
  const image = sharp(bytes, { failOn: options.failOn ?? "warning" }).ensureAlpha();
  const trimmed = options.trim === false
    ? image
    : image.trim({
      background: options.trimBackground ?? { r: 0, g: 0, b: 0, alpha: 0 },
      threshold: options.trimThreshold ?? 2,
    });
  return trimmed.png().toBuffer();
}

export async function addTransparentIcon(slide, input, position, alt, options = {}) {
  assertSlide(slide);
  const frame = assertPosition(position);
  if (typeof alt !== "string" || !alt.trim()) throw new TypeError("A non-empty alt description is required.");
  const theme = options.theme ?? {};
  const bytes = await trimTransparentRaster(input, options);

  let supportCircle = null;
  if (options.supportCircle === true) {
    const padding = Number(options.supportPadding ?? 0);
    const diameter = Number(options.supportDiameter ?? Math.min(frame.width, frame.height) + padding * 2);
    const fill = options.supportFill ?? themeColor(theme, "iconSupport");
    if (fill == null) throw new Error("supportCircle:true requires options.supportFill or theme.colors.iconSupport.");
    supportCircle = addShape(slide, "ellipse", {
      left: frame.left + (frame.width - diameter) / 2,
      top: frame.top + (frame.height - diameter) / 2,
      width: diameter,
      height: diameter,
    }, {
      name: options.name
        ? `${options.name}-support`
        : resolvedObjectName(options, { part: "support" }),
      fill,
      line: options.supportLine ?? makeLine(options.supportBorder, options.supportBorderWidth ?? 0),
    });
  }

  const image = slide.images.add(compactObject({
    name: options.name ?? resolvedObjectName(options, { part: "icon" }),
    blob: bytes,
    contentType: "image/png",
    alt: alt.trim(),
    fit: options.fit ?? "contain",
    position: frame,
    geometry: options.geometry ?? "rect",
    borderRadius: options.borderRadius,
    prompt: options.prompt,
  }));
  return { image, supportCircle, bytes };
}

export async function addIconCell(slide, input, position, alt, options = {}) {
  const family = normalizeSemanticId(options.family, "icon family");
  const instance = normalizeSemanticId(options.instance, "icon instance");
  if (!family || !instance) throw new TypeError("addIconCell requires options.family and options.instance.");
  return addTransparentIcon(slide, input, position, alt, {
    ...options,
    semantic: { ...(options.semantic ?? {}), role: options.role ?? "icon", family, instance },
  });
}

export function createRebuildHelpers(theme) {
  const inject = (options) => ({ ...(options ?? {}), theme: options?.theme ?? theme });
  return {
    addShape,
    addTextBox: (slide, text, position, options) => addTextBox(slide, text, position, inject(options)),
    addCard: (slide, position, options) => addCard(slide, position, inject(options)),
    addTitle: (slide, text, position, options) => addTitle(slide, text, position, inject(options)),
    addRoleTextBox: (slide, text, position, options) => addRoleTextBox(slide, text, position, inject(options)),
    addComponentCard: (slide, position, options) => addComponentCard(slide, position, inject(options)),
    addDivider: (slide, position, options) => addDivider(slide, position, inject(options)),
    addConnector: (slide, source, target, options) => addConnector(slide, source, target, inject(options)),
    addStructuredBullets: (slide, items, position, options) => addStructuredBullets(slide, items, position, inject(options)),
    addNativeBullets: (slide, items, position, options) => addNativeBullets(slide, items, position, inject(options)),
    addHangingBullets: (slide, items, position, options) => addHangingBullets(slide, items, position, inject(options)),
    addTransparentIcon: (slide, input, position, alt, options) => addTransparentIcon(slide, input, position, alt, inject(options)),
    addIconCell: (slide, input, position, alt, options) => addIconCell(slide, input, position, alt, inject(options)),
    semanticName,
    wrapText,
    estimateGlyphWidth,
    ptToPx,
    pxToPt,
  };
}

export const __internal = Object.freeze({
  compactObject,
  makeLine,
  normalizeBulletItem,
  themeColor,
  themeFont,
  tokenizeForWrap,
});

// Make direct execution useful during authoring without turning this helper
// module into a CLI dependency.
if (process.argv[1] && pathToFileURL(path.resolve(process.argv[1])).href === import.meta.url) {
  console.log("rebuild_helpers.mjs exports parameterized @oai/artifact-tool helpers; import it from a deck rebuild script.");
}
