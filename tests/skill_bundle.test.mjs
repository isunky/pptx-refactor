import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { validateSkillBundle } from "../scripts/validate_skill_bundle.mjs";

const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

async function makePackage() {
  const temp = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-refactor-package-"));
  const root = path.join(temp, "pptx-refactor");
  await fs.mkdir(root);
  for (const entry of ["SKILL.md", "LICENSE", "agents", "references", "scripts"]) {
    await fs.cp(path.join(REPO_ROOT, entry), path.join(root, entry), { recursive: true });
  }
  return root;
}

test("repository Skill metadata and local links are valid", async () => {
  const result = await validateSkillBundle(REPO_ROOT);
  assert.equal(result.valid, true, result.errors.join("\n"));
});

test("release package accepts only the public Skill payload", async (t) => {
  const root = await makePackage();
  t.after(() => fs.rm(path.dirname(root), { recursive: true, force: true }));
  const result = await validateSkillBundle(root, { packageMode: true });
  assert.equal(result.valid, true, result.errors.join("\n"));

  await fs.writeFile(path.join(root, "private-notes.txt"), "not for release\n", "utf8");
  const rejected = await validateSkillBundle(root, { packageMode: true });
  assert.equal(rejected.valid, false);
  assert.match(rejected.errors.join("\n"), /Unexpected release-package entry/u);
});

test("broken Skill references fail validation", async (t) => {
  const root = await makePackage();
  t.after(() => fs.rm(path.dirname(root), { recursive: true, force: true }));
  await fs.appendFile(path.join(root, "SKILL.md"), "\n[missing](references/not-found.md)\n", "utf8");
  const result = await validateSkillBundle(root, { packageMode: true });
  assert.equal(result.valid, false);
  assert.match(result.errors.join("\n"), /Broken local link/u);
});
