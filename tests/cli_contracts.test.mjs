import assert from "node:assert/strict";
import path from "node:path";
import { spawnSync } from "node:child_process";
import test from "node:test";
import { fileURLToPath } from "node:url";

const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const HELP_SCRIPTS = [
  "analyze_hybrid_deck.mjs",
  "prepare_raster_asset.mjs",
  "qa_conversion.mjs",
  "unzip_compat.mjs",
  "validate_conversion_plan.mjs",
  "validate_skill_bundle.mjs",
];

for (const script of HELP_SCRIPTS) {
  test(`${script} exposes a successful help contract`, () => {
    const result = spawnSync(process.execPath, [path.join(REPO_ROOT, "scripts", script), "--help"], {
      cwd: REPO_ROOT,
      encoding: "utf8",
    });
    assert.equal(result.status, 0, result.stderr);
    assert.match(result.stdout, /Usage:/u);
  });
}

test("conversion-plan validator rejects missing required options with exit code 2", () => {
  const result = spawnSync(process.execPath, [path.join(REPO_ROOT, "scripts", "validate_conversion_plan.mjs")], {
    cwd: REPO_ROOT,
    encoding: "utf8",
  });
  assert.equal(result.status, 2);
  assert.match(result.stderr, /Missing required option/u);
});

test("Skill-bundle validator rejects unknown options with exit code 2", () => {
  const result = spawnSync(process.execPath, [path.join(REPO_ROOT, "scripts", "validate_skill_bundle.mjs"), "--unknown"], {
    cwd: REPO_ROOT,
    encoding: "utf8",
  });
  assert.equal(result.status, 2);
  assert.match(result.stderr, /Unknown option/u);
});
