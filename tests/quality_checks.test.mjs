import assert from "node:assert/strict";
import test from "node:test";

import {
  accountExpectedText,
  countComparableOccurrences,
  missingSensitiveTokens,
  textStyleSegments,
} from "../scripts/quality_checks.mjs";

test("short numbers do not match inside longer numbers", () => {
  assert.equal(countComparableOccurrences("完成率 100%，目标 10%", "10"), 1);
  assert.equal(countComparableOccurrences("完成率 100%", "10"), 0);
});

test("duplicate expected text requires duplicate output occurrences", () => {
  const result = accountExpectedText(["增长", "增长"], "增长");
  assert.deepEqual(result, [{ text: "增长", required: 2, found: 1, accounted: false }]);
});

test("adjacent text objects retain token boundaries", () => {
  assert.equal(countComparableOccurrences("02\u0000A screenshot", "02"), 1);
  assert.equal(countComparableOccurrences("SYNTHETIC DEMO 2026\u0000Flattened slides", "SYNTHETIC DEMO 2026"), 1);
});

test("numbers, signs, percentages, and CJK units are checked separately", () => {
  const missing = missingSensitiveTokens("收入 +12.5%，金额 3亿元", "收入 12.5%，金额 3亿元");
  assert.deepEqual(missing, [{ token: "+12.5%", required: 1, found: 0 }]);
});

test("text style inspection covers every non-empty run", () => {
  const segments = textStyleSegments({
    resolvedTextStyle: { typeface: "Aptos", fontSize: 24, color: "#111111" },
    paragraphs: [{
      lineSpacing: 1.1,
      runs: [
        { text: "一致", fontSize: 24 },
        { text: "漂移", fontSize: 18, bold: true },
      ],
    }],
  });
  assert.equal(segments.length, 2);
  assert.equal(segments[0].fontSize, 24);
  assert.equal(segments[1].fontSize, 18);
  assert.equal(segments[1].bold, true);
});
