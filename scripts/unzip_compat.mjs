#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { createRequire, syncBuiltinESMExports } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";

const HELP = `Usage:
  unzip -Z1 <archive>
  unzip -p <archive> <entry>
  unzip_compat.mjs --run <script.mjs> [script arguments...]

Compatibility shim used by pptx-refactor on Windows.

Options:
  -Z1              List archive entries, one UTF-8 path per line.
  -p               Write one entry to stdout without conversion.
  --run             Run a Node script while intercepting its spawnSync("unzip")
                    calls. Use this for bundled presentation inspectors on
                    Windows, where spawnSync does not resolve unzip.cmd.
  -h, --help       Show this help.

The -p form is binary-safe. RUNTIME_NODE_MODULES must point at the bundled
workspace Node.js package directory containing jszip.
`;

class CliError extends Error {
  constructor(message, exitCode = 2) {
    super(message);
    this.name = "CliError";
    this.exitCode = exitCode;
  }
}

async function loadRuntimeModule(packageName) {
  const modulesDir = process.env.RUNTIME_NODE_MODULES;
  if (!modulesDir) {
    throw new CliError(
      "RUNTIME_NODE_MODULES is not set. Load the bundled workspace dependencies before using this shim.",
    );
  }
  const requireFromRuntime = createRequire(
    path.join(path.resolve(modulesDir), "pptx-refactor-loader.cjs"),
  );
  let resolved;
  try {
    resolved = requireFromRuntime.resolve(packageName);
  } catch (error) {
    throw new CliError(
      `Cannot resolve ${packageName} from RUNTIME_NODE_MODULES (${modulesDir}): ${error.message}`,
    );
  }
  return import(pathToFileURL(resolved).href);
}

function parseArgs(argv) {
  if (argv.length === 0 || argv.includes("--help") || argv.includes("-h")) {
    return { help: true };
  }
  const [mode, archive, entry, ...extra] = argv;
  if (mode === "--run") {
    if (!archive) throw new CliError("--run requires a target Node script.");
    return { mode, script: archive, scriptArgs: [entry, ...extra].filter((value) => value !== undefined) };
  }
  if (mode === "-Z1") {
    if (!archive || entry !== undefined || extra.length > 0) {
      throw new CliError("-Z1 requires exactly one archive path.");
    }
    return { mode, archive };
  }
  if (mode === "-p") {
    if (!archive || !entry || extra.length > 0) {
      throw new CliError("-p requires exactly one archive path and one entry path.");
    }
    return { mode, archive, entry };
  }
  throw new CliError(`Unsupported unzip arguments: ${argv.join(" ")}`);
}

async function runWithUnzipCompat(scriptPath, scriptArgs) {
  const resolvedScript = path.resolve(scriptPath);
  let stat;
  try {
    stat = await fs.stat(resolvedScript);
  } catch {
    throw new CliError(`Runner target not found: ${resolvedScript}`);
  }
  if (!stat.isFile()) throw new CliError(`Runner target is not a regular file: ${resolvedScript}`);

  const require = createRequire(import.meta.url);
  const childProcess = require("node:child_process");
  const originalSpawnSync = childProcess.spawnSync;
  const shimPath = fileURLToPath(import.meta.url);
  childProcess.spawnSync = function spawnSyncWithUnzipCompat(command, args = [], options = {}) {
    const executable = path.basename(String(command)).toLowerCase();
    if (executable === "unzip" || executable === "unzip.exe" || executable === "unzip.cmd") {
      return originalSpawnSync(process.execPath, [shimPath, ...args], options);
    }
    return originalSpawnSync(command, args, options);
  };
  syncBuiltinESMExports();

  // Artifact Tool may leave a Windows native teardown hook that surfaces
  // 0xC0000409 after a successful custom-canvas run. The bundled template
  // scripts each emit one recognizable terminal record after their outputs are
  // durable. Flush that record and exit before the late native teardown.
  const targetName = path.basename(resolvedScript).toLowerCase();
  const originalConsoleLog = console.log;
  console.log = (...values) => {
    originalConsoleLog(...values);
    const first = typeof values[0] === "string" ? values[0] : "";
    const terminalRecord =
      (targetName === "inspect_template_deck.mjs" && /template-manifest\.json\s*$/i.test(first))
      || (targetName === "prepare_template_starter_deck.mjs" && /"manifestPath"\s*:/i.test(first))
      || (targetName === "check_template_fidelity.mjs" && /"status"\s*:/i.test(first));
    if (terminalRecord) {
      process.stdout.write("", () => process.exit(process.exitCode || 0));
    }
  };
  process.once("beforeExit", (code) => process.exit(code || process.exitCode || 0));
  process.argv = [process.execPath, resolvedScript, ...scriptArgs];
  await import(`${pathToFileURL(resolvedScript).href}?unzip_compat_run=${Date.now()}`);
}

async function loadArchive(archivePath) {
  const resolvedArchive = path.resolve(archivePath);
  let stat;
  try {
    stat = await fs.stat(resolvedArchive);
  } catch (error) {
    throw new CliError(`Archive not found: ${resolvedArchive}`, 9);
  }
  if (!stat.isFile()) {
    throw new CliError(`Archive is not a regular file: ${resolvedArchive}`, 9);
  }

  const jszipModule = await loadRuntimeModule("jszip");
  const JSZip = jszipModule.default ?? jszipModule;
  try {
    return await JSZip.loadAsync(await fs.readFile(resolvedArchive));
  } catch (error) {
    throw new CliError(
      `Unable to read ZIP archive ${resolvedArchive}: ${error.message}`,
      9,
    );
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    process.stdout.write(HELP);
    return;
  }

  if (args.mode === "--run") {
    await runWithUnzipCompat(args.script, args.scriptArgs);
    return;
  }

  const archive = await loadArchive(args.archive);
  if (args.mode === "-Z1") {
    const names = Object.keys(archive.files);
    if (names.length > 0) {
      process.stdout.write(`${names.join("\n")}\n`);
    }
    return;
  }

  const entry = archive.file(args.entry);
  if (!entry) {
    throw new CliError(
      `Entry not found in ${path.resolve(args.archive)}: ${args.entry}`,
      11,
    );
  }
  const bytes = await entry.async("nodebuffer");
  await new Promise((resolve, reject) => {
    process.stdout.write(bytes, (error) => (error ? reject(error) : resolve()));
  });
}

main().catch((error) => {
  const exitCode = Number.isInteger(error.exitCode) ? error.exitCode : 1;
  process.stderr.write(`unzip: ${error.message}\n`);
  process.exitCode = exitCode;
});
