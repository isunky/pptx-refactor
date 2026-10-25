import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const python = process.env.PYTHON || (process.platform === "win32" ? "python" : "python3");
test("bounded palette editing preserves package contents and rejects unsafe plans", () => {
  const result = spawnSync(python, [path.join(root, "tests", "test_palette.py"), "-v"], { encoding: "utf8" });
  assert.equal(result.status, 0, result.stderr || result.error?.message);
});
test("palette CLI exposes help without workspace packages", () => {
  const result = spawnSync(python, [path.join(root, "scripts", "apply_palette.py"), "--help"], { encoding: "utf8" });
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /--dry-run/u);
});
test("palette CLI rejects incomplete arguments", () => {
  const result = spawnSync(python, [path.join(root, "scripts", "apply_palette.py"), "--source", "missing.pptx"], { encoding: "utf8" });
  assert.equal(result.status, 2);
});
