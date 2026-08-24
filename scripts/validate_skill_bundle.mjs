#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const SKILL_ID = "pptx-refactor";
const REQUIRED_PATHS = [
  "SKILL.md",
  "LICENSE",
  "agents/openai.yaml",
  "references",
  "scripts",
];
const PACKAGE_ALLOWLIST = new Set(["SKILL.md", "LICENSE", "agents", "references", "scripts"]);
const JUNK_NAMES = new Set([".DS_Store", "Thumbs.db", "desktop.ini"]);
const SOURCE_ONLY_IGNORES = new Set([".git", ".tmp", "node_modules"]);

const HELP = `Validate the PPTX Refactor Skill source or a release package.

Usage:
  node scripts/validate_skill_bundle.mjs [--root <directory>] [--package]

Options:
  --root <directory>  Skill root. Defaults to the repository root.
  --package           Enforce the public release-package allowlist and folder name.
  -h, --help          Show this help.
`;

async function exists(target) {
  try {
    await fs.access(target);
    return true;
  } catch {
    return false;
  }
}

async function walk(root, { packageMode = false } = {}) {
  const entries = [];
  for (const dirent of await fs.readdir(root, { withFileTypes: true })) {
    if (!packageMode && SOURCE_ONLY_IGNORES.has(dirent.name)) continue;
    const absolute = path.join(root, dirent.name);
    entries.push({ absolute, relative: path.relative(root, absolute), dirent });
    if (dirent.isDirectory()) {
      for (const child of await walk(absolute, { packageMode })) {
        entries.push({ ...child, relative: path.join(dirent.name, child.relative) });
      }
    }
  }
  return entries;
}

function unquote(value) {
  const trimmed = String(value ?? "").trim();
  if (
    trimmed.length >= 2
    && ((trimmed.startsWith('"') && trimmed.endsWith('"'))
      || (trimmed.startsWith("'") && trimmed.endsWith("'")))
  ) {
    return trimmed.slice(1, -1);
  }
  return trimmed;
}

function parseFrontmatter(markdown) {
  const normalized = markdown.replace(/^\uFEFF/u, "").replaceAll("\r\n", "\n");
  if (!normalized.startsWith("---\n")) return null;
  const end = normalized.indexOf("\n---\n", 4);
  if (end < 0) return null;
  const values = {};
  for (const line of normalized.slice(4, end).split("\n")) {
    const match = line.match(/^([a-zA-Z0-9_-]+):\s*(.+)$/u);
    if (match) values[match[1]] = unquote(match[2]);
  }
  return values;
}

function parseYamlScalar(source, key) {
  const match = source.match(new RegExp(`^\\s*${key}:\\s*(.+)$`, "mu"));
  return match ? unquote(match[1]) : null;
}

function localMarkdownLinks(markdown) {
  const links = [];
  const pattern = /\[[^\]]*\]\(([^)]+)\)/gu;
  for (const match of markdown.matchAll(pattern)) {
    const raw = match[1].trim().replace(/^<|>$/gu, "");
    if (!raw || raw.startsWith("#") || /^[a-z][a-z0-9+.-]*:/iu.test(raw)) continue;
    links.push(decodeURIComponent(raw.split("#", 1)[0]));
  }
  return links;
}

export async function validateSkillBundle(root, { packageMode = false } = {}) {
  const resolvedRoot = path.resolve(root);
  const errors = [];
  const warnings = [];

  for (const relative of REQUIRED_PATHS) {
    if (!(await exists(path.join(resolvedRoot, relative)))) errors.push(`Missing required path: ${relative}`);
  }
  if (errors.length) return { valid: false, errors, warnings, root: resolvedRoot };

  const skillMarkdown = await fs.readFile(path.join(resolvedRoot, "SKILL.md"), "utf8");
  const frontmatter = parseFrontmatter(skillMarkdown);
  if (!frontmatter) errors.push("SKILL.md must start with valid YAML frontmatter.");
  if (frontmatter?.name !== SKILL_ID) errors.push(`SKILL.md name must be ${SKILL_ID}.`);
  if (!frontmatter?.description || frontmatter.description.length < 40) {
    errors.push("SKILL.md description must be a discriminating description of at least 40 characters.");
  }

  const agentYaml = await fs.readFile(path.join(resolvedRoot, "agents", "openai.yaml"), "utf8");
  for (const key of ["display_name", "short_description", "default_prompt"]) {
    if (!parseYamlScalar(agentYaml, key)) errors.push(`agents/openai.yaml is missing interface.${key}.`);
  }
  if (!String(parseYamlScalar(agentYaml, "default_prompt") ?? "").includes(`$${SKILL_ID}`)) {
    errors.push(`agents/openai.yaml default_prompt must invoke $${SKILL_ID}.`);
  }

  const entries = await walk(resolvedRoot, { packageMode });
  for (const entry of entries) {
    if (entry.dirent.isSymbolicLink()) errors.push(`Symbolic links are not allowed in the Skill bundle: ${entry.relative}`);
    if (JUNK_NAMES.has(entry.dirent.name)) errors.push(`Remove platform junk file: ${entry.relative}`);
    if (/\.(?:ppt|pptx|key)$/iu.test(entry.dirent.name) && packageMode) {
      errors.push(`Presentation fixtures must not be shipped in the Skill package: ${entry.relative}`);
    }
  }

  const markdownFiles = entries
    .filter((entry) => entry.dirent.isFile() && entry.relative.toLowerCase().endsWith(".md"))
    .map((entry) => entry.absolute);
  markdownFiles.push(path.join(resolvedRoot, "SKILL.md"));
  for (const markdownPath of new Set(markdownFiles)) {
    const source = await fs.readFile(markdownPath, "utf8");
    const displayPath = path.relative(resolvedRoot, markdownPath) || "SKILL.md";
    if (/(?:file:\/\/|[A-Za-z]:\\(?:Users|Documents|Desktop)\\|\/(?:Users|home)\/[^\s`]+)/u.test(source)) {
      errors.push(`Absolute user path found in ${displayPath}; Skill paths must be portable.`);
    }
    for (const link of localMarkdownLinks(source)) {
      const target = path.resolve(path.dirname(markdownPath), link);
      if (!(await exists(target))) errors.push(`Broken local link in ${displayPath}: ${link}`);
    }
  }

  if (packageMode) {
    if (path.basename(resolvedRoot) !== SKILL_ID) errors.push(`Release package folder must be named ${SKILL_ID}.`);
    const topLevel = await fs.readdir(resolvedRoot);
    for (const name of topLevel) {
      if (!PACKAGE_ALLOWLIST.has(name)) errors.push(`Unexpected release-package entry: ${name}`);
    }
  }

  return { valid: errors.length === 0, errors, warnings, root: resolvedRoot };
}

function parseArgs(argv) {
  const result = { root: path.resolve(path.dirname(fileURLToPath(import.meta.url)), ".."), packageMode: false };
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (token === "--help" || token === "-h") return { help: true };
    if (token === "--package") {
      result.packageMode = true;
      continue;
    }
    if (token === "--root") {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) throw new Error("--root requires a directory.");
      result.root = path.resolve(value);
      index += 1;
      continue;
    }
    throw new Error(`Unknown option: ${token}`);
  }
  return result;
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
  const result = await validateSkillBundle(args.root, { packageMode: args.packageMode });
  if (!result.valid) {
    console.error(`Skill validation failed with ${result.errors.length} error(s):`);
    for (const error of result.errors) console.error(`  - ${error}`);
    process.exitCode = 1;
    return;
  }
  console.log(`Skill validation passed: ${result.root}`);
}

const invokedAsScript = process.argv[1]
  && path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url));
if (invokedAsScript) await main();
