const ASCII_WORD = /[A-Za-z0-9]/u;
const SENSITIVE_TOKEN = /[+-]?(?:\d{1,3}(?:[,，]\d{3})+|\d+)(?:[.．]\d+)?(?:[%％]|(?:万|亿)?(?:元|人|个|项|页|天|日|月|年|小时|分钟|秒|毫秒|kg|g|km|cm|mm|GB|MB|KB|ms|s))?/giu;

export function normalizeComparableText(value) {
  return String(value ?? "")
    .normalize("NFKC")
    .replace(/[•●▪‣]/gu, "")
    .replace(/\s+/gu, "")
    .trim();
}

function occurrenceAt(haystack, needle, index) {
  const before = index > 0 ? haystack[index - 1] : "";
  const after = index + needle.length < haystack.length ? haystack[index + needle.length] : "";
  const needsLeftBoundary = ASCII_WORD.test(needle[0] ?? "");
  const needsRightBoundary = ASCII_WORD.test(needle.at(-1) ?? "");
  return !(needsLeftBoundary && ASCII_WORD.test(before))
    && !(needsRightBoundary && ASCII_WORD.test(after));
}

export function countComparableOccurrences(text, expected) {
  const haystack = normalizeComparableText(text);
  const needle = normalizeComparableText(expected);
  if (!needle) return 0;
  let count = 0;
  let cursor = 0;
  while (cursor <= haystack.length - needle.length) {
    const index = haystack.indexOf(needle, cursor);
    if (index < 0) break;
    if (occurrenceAt(haystack, needle, index)) count += 1;
    cursor = index + Math.max(1, needle.length);
  }
  return count;
}

export function accountExpectedText(expectedItems, actualText) {
  const requirements = new Map();
  for (const item of expectedItems ?? []) {
    const text = typeof item === "object" && item !== null ? item.text : item;
    const normalized = normalizeComparableText(text);
    if (!normalized) continue;
    const current = requirements.get(normalized) ?? { text: String(text), required: 0 };
    current.required += 1;
    requirements.set(normalized, current);
  }
  return [...requirements.values()].map((requirement) => {
    const found = countComparableOccurrences(actualText, requirement.text);
    return { ...requirement, found, accounted: found >= requirement.required };
  });
}

export function sensitiveTokenCounts(value) {
  const counts = new Map();
  for (const match of String(value ?? "").normalize("NFKC").matchAll(SENSITIVE_TOKEN)) {
    const token = match[0].replaceAll("，", ",").replaceAll("．", ".");
    counts.set(token, (counts.get(token) ?? 0) + 1);
  }
  return counts;
}

export function missingSensitiveTokens(expectedText, actualText) {
  const expected = sensitiveTokenCounts(expectedText);
  const actual = sensitiveTokenCounts(actualText);
  return [...expected.entries()]
    .filter(([token, required]) => (actual.get(token) ?? 0) < required)
    .map(([token, required]) => ({ token, required, found: actual.get(token) ?? 0 }));
}

function styleValue(candidates, key) {
  return candidates.find((candidate) => candidate?.[key] !== undefined)?.[key];
}

export function textStyleSegments(element) {
  const paragraphs = Array.isArray(element?.paragraphs) ? element.paragraphs : [];
  const segments = [];
  for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
    const runs = Array.isArray(paragraph?.runs) ? paragraph.runs : [];
    for (const [runIndex, run] of runs.entries()) {
      const text = String(run?.text ?? run?.run ?? "");
      if (!text.trim()) continue;
      const candidates = [run, paragraph?.resolvedTextStyle, element?.resolvedTextStyle, element];
      segments.push({
        paragraphIndex,
        runIndex,
        text,
        typeface: styleValue(candidates, "typeface"),
        fontSize: Number(styleValue(candidates, "fontSize") ?? element?.resolvedFontSize),
        color: String(styleValue(candidates, "color") ?? "").toLowerCase(),
        bold: styleValue(candidates, "bold"),
        italic: styleValue(candidates, "italic"),
        underline: styleValue(candidates, "underline"),
        alignment: styleValue([paragraph, ...candidates], "alignment"),
        lineSpacing: Number(paragraph?.lineSpacing ?? styleValue(candidates, "lineSpacing")),
      });
    }
  }
  if (segments.length > 0) return segments;
  const text = String(element?.text ?? element?.textPreview ?? "");
  if (!text.trim()) return [];
  return [{
    paragraphIndex: 0,
    runIndex: 0,
    text,
    typeface: element?.resolvedTextStyle?.typeface ?? element?.typeface,
    fontSize: Number(element?.resolvedTextStyle?.fontSize ?? element?.resolvedFontSize ?? element?.fontSize),
    color: String(element?.resolvedTextStyle?.color ?? element?.color ?? "").toLowerCase(),
    bold: element?.resolvedTextStyle?.bold ?? element?.bold,
    italic: element?.resolvedTextStyle?.italic ?? element?.italic,
    underline: element?.resolvedTextStyle?.underline ?? element?.underline,
    alignment: element?.resolvedTextStyle?.alignment ?? element?.alignment,
    lineSpacing: Number(element?.resolvedTextStyle?.lineSpacing ?? element?.lineSpacing),
  }];
}
