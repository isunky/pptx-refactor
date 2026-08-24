import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { validateConversionPlan } from "../scripts/validate_conversion_plan.mjs";

function canonicalJson(value) {
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(",")}]`;
  if (value && typeof value === "object") {
    return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${canonicalJson(value[key])}`).join(",")}}`;
  }
  return JSON.stringify(value);
}

function sha256(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
}

async function makeFixture(mutator = () => {}) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "pptx-refactor-plan-"));
  const sourceBytes = Buffer.from("synthetic-pptx-fixture\n", "utf8");
  const sourceHash = sha256(sourceBytes);
  const sourcePath = path.join(root, "source.pptx");
  await fs.writeFile(sourcePath, sourceBytes);

  const styleProfile = {
    roles: {
      title: {
        fontSize: 42,
        maxLines: 2,
        typeface: "Aptos Display",
        color: "#172033",
      },
    },
    componentFamilies: {},
    iconFamily: {
      id: "generic-icons",
      sharedPromptPrefix: "Minimal flat icon with a transparent background",
      opticalCoverage: 0.72,
      centroidTolerance: 0.04,
      opticalTolerance: 0.08,
    },
  };

  const manifest = {
    schemaVersion: "1.1",
    sourcePptx: "source.pptx",
    sourceSha256: sourceHash,
    slideSize: { width: 13.333, height: 7.5, unit: "in" },
    renderSlideSizePx: { width: 1280, height: 720, unit: "px" },
    slideCount: 1,
    largeRasters: [],
    slides: [
      {
        slideNumber: 1,
        classification: "native-editable",
        dimensions: { width: 13.333, height: 7.5, unit: "in" },
        renderDimensionsPx: { width: 1280, height: 720, unit: "px" },
        objects: [
          {
            id: "tx/1",
            slideNumber: 1,
            kind: "text",
            text: "Synthetic title",
            bbox: { x: 1, y: 1, width: 6, height: 1, unit: "in" },
            slideLocal: true,
          },
        ],
      },
    ],
  };

  const plan = {
    schemaVersion: "1.1",
    mode: "balanced",
    sourcePptx: "source.pptx",
    sourceManifest: "source-manifest.json",
    outputPptx: "source_editable.pptx",
    sourceSha256: sourceHash,
    visualPolicy: {
      normalizationScope: "full-deck-role-based",
      iconMode: "extract-or-regenerate-generic",
      qaStrictness: "tiered",
      calibrationMode: "automatic",
    },
    styleProfile,
    calibration: {
      mode: "automatic",
      representativeSlides: [1],
      status: "complete",
      evidence: [{ slideNumber: 1, render: "slide-1.png" }],
      frozenProfileSha256: sha256(Buffer.from(canonicalJson(styleProfile), "utf8")),
    },
    feedbackIssues: [],
    assets: [],
    slides: [
      {
        slideNumber: 1,
        sourceSlide: 1,
        classification: "native-editable",
        regions: [
          {
            regionId: "s01-title",
            slideNumber: 1,
            bbox: { x: 1, y: 1, width: 6, height: 1, unit: "in" },
            sourceObjectIds: ["tx/1"],
            action: "keep-native",
            targetType: "text",
            expectedText: ["Synthetic title"],
            confidence: 1,
            reason: "The title is already a native editable text object.",
            intent: "faithful-rebuild",
            styleRole: "title",
            editability: ["text-editable"],
          },
        ],
      },
    ],
  };

  await mutator({ manifest, plan, styleProfile });
  const manifestPath = path.join(root, "source-manifest.json");
  const planPath = path.join(root, "conversion-plan.json");
  const outMapPath = path.join(root, "template-frame-map.json");
  const reportPath = path.join(root, "validation-report.json");
  await fs.writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
  await fs.writeFile(planPath, `${JSON.stringify(plan, null, 2)}\n`, "utf8");
  return { root, manifestPath, planPath, outMapPath, reportPath };
}

async function validateFixture(t, mutator) {
  const fixture = await makeFixture(mutator);
  t.after(() => fs.rm(fixture.root, { recursive: true, force: true }));
  return validateConversionPlan(fixture);
}

test("valid schema 1.1 conversion plan produces a frame map", async (t) => {
  const { report, frameMap } = await validateFixture(t);
  assert.equal(report.valid, true, report.errors.join("\n"));
  assert.equal(frameMap.schemaVersion, "1.1");
  assert.equal(frameMap.slides.length, 1);
});

test("wildcard source-object selectors are rejected", async (t) => {
  const { report } = await validateFixture(t, ({ plan }) => {
    plan.slides[0].regions[0].sourceObjectIds = ["*"];
  });
  assert.equal(report.valid, false);
  assert.match(report.errors.join("\n"), /exact object ID/u);
});

test("low-confidence destructive actions are rejected", async (t) => {
  const { report } = await validateFixture(t, ({ plan }) => {
    Object.assign(plan.slides[0].regions[0], { action: "rebuild-text", confidence: 0.5 });
  });
  assert.equal(report.valid, false);
  assert.match(report.errors.join("\n"), /below 0\.8/u);
});

test("unitless region bounds are rejected", async (t) => {
  const { report } = await validateFixture(t, ({ plan }) => {
    delete plan.slides[0].regions[0].bbox.unit;
  });
  assert.equal(report.valid, false);
  assert.match(report.errors.join("\n"), /unit must declare/u);
});

test("in-place source overwrite is rejected", async (t) => {
  const { report } = await validateFixture(t, ({ plan }) => {
    plan.outputPptx = "source.pptx";
  });
  assert.equal(report.valid, false);
  assert.match(report.errors.join("\n"), /must not overwrite/u);
});

test("stale style-profile calibration hashes are rejected", async (t) => {
  const { report } = await validateFixture(t, ({ plan }) => {
    plan.calibration.frozenProfileSha256 = "0".repeat(64);
  });
  assert.equal(report.valid, false);
  assert.match(report.errors.join("\n"), /canonical SHA-256/u);
});

test("large rasters require an explicit raster disposition", async (t) => {
  const { report } = await validateFixture(t, ({ manifest }) => {
    Object.assign(manifest.slides[0].objects[0], {
      kind: "image",
      isLargeRaster: true,
      coverageRatio: 0.5,
    });
    manifest.slides[0].classification = "mixed";
  });
  assert.equal(report.valid, false);
  assert.match(report.errors.join("\n"), /large raster|Large raster/iu);
});
