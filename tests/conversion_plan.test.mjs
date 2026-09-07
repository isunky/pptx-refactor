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

async function validateFixture(t, mutator, options = {}) {
  const fixture = await makeFixture(mutator);
  t.after(() => fs.rm(fixture.root, { recursive: true, force: true }));
  return validateConversionPlan({ ...fixture, ...options });
}

test("valid schema 1.1 conversion plan produces a frame map", async (t) => {
  const { report, frameMap } = await validateFixture(t);
  assert.equal(report.valid, true, report.errors.join("\n"));
  assert.equal(frameMap.schemaVersion, "1.1");
  assert.equal(frameMap.slides.length, 1);
});

test("conversion-plan validator rejects unknown stages", async (t) => {
  const fixture = await makeFixture();
  t.after(() => fs.rm(fixture.root, { recursive: true, force: true }));
  await assert.rejects(validateConversionPlan({ ...fixture, stage: "preview" }), /stage must be calibration or final/u);
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

test("pending calibration is accepted only for calibration-stage validation", async (t) => {
  const mutator = ({ plan }) => {
    plan.calibration.status = "pending";
    plan.calibration.evidence = [];
    delete plan.calibration.frozenProfileSha256;
  };
  const calibration = await validateFixture(t, mutator, { stage: "calibration" });
  assert.equal(calibration.report.valid, true, calibration.report.errors.join("\n"));
  assert.equal(calibration.report.stage, "calibration");
  assert.equal(calibration.report.statistics.activeSlideCount, 1);
  assert.equal(calibration.frameMap.stage, "calibration");

  const final = await validateFixture(t, mutator, { stage: "final" });
  assert.equal(final.report.valid, false);
  assert.match(final.report.errors.join("\n"), /status must be complete|frozenProfileSha256/u);
});

test("calibration-stage frame maps preserve non-representative slides", async (t) => {
  const { report, frameMap } = await validateFixture(t, ({ manifest, plan }) => {
    manifest.slideCount = 2;
    manifest.slides.push({
      slideNumber: 2,
      classification: "native-editable",
      dimensions: { width: 13.333, height: 7.5, unit: "in" },
      renderDimensionsPx: { width: 1280, height: 720, unit: "px" },
      objects: [{ id: "tx/2", slideNumber: 2, kind: "text", text: "Second slide", bbox: { x: 1, y: 1, width: 6, height: 1, unit: "in" }, slideLocal: true }],
    });
    plan.slides.push({
      slideNumber: 2,
      sourceSlide: 2,
      classification: "native-editable",
      regions: [{
        regionId: "s02-title",
        slideNumber: 2,
        bbox: { x: 1, y: 1, width: 6, height: 1, unit: "in" },
        sourceObjectIds: ["tx/2"],
        action: "rebuild-text",
        targetType: "text",
        expectedText: ["Second slide"],
        confidence: 1,
        reason: "Normalize the second slide title.",
        intent: "style-normalization",
        styleRole: "title",
        editability: ["text-editable"],
      }],
    });
    plan.calibration.status = "pending";
    plan.calibration.evidence = [];
    delete plan.calibration.frozenProfileSha256;
  }, { stage: "calibration" });
  assert.equal(report.valid, true, report.errors.join("\n"));
  assert.equal(frameMap.slides[1].calibrationActive, false);
  assert.deepEqual(frameMap.slides[1].deleteObjectIds, []);
  assert.deepEqual(frameMap.slides[1].addZones, []);
  assert.ok(frameMap.outputSlides[1].editTargets.every((target) => target.action === "keep"));
});

test("targetBbox moves output without broadening source deletion authority", async (t) => {
  const { report, frameMap } = await validateFixture(t, ({ plan }) => {
    Object.assign(plan.slides[0].regions[0], {
      action: "rebuild-text",
      intent: "style-normalization",
      targetBbox: { x: 2, y: 1.25, width: 6, height: 1, unit: "in" },
    });
  });
  assert.equal(report.valid, true, report.errors.join("\n"));
  assert.ok(Math.abs(frameMap.slides[0].regions[0].sourceBbox.left - 96) < 0.01);
  assert.ok(Math.abs(frameMap.slides[0].regions[0].bbox.left - 192) < 0.01);
  assert.ok(Math.abs(frameMap.slides[0].addZones[0].bbox.left - 192) < 0.01);
  assert.equal(frameMap.outputSlides[0].editTargets[0].action, "delete");
  assert.ok(Math.abs(frameMap.outputSlides[0].editTargets[1].zone.left - 192) < 0.01);
});

test("targetBbox requires an explicit normalization or redesign intent", async (t) => {
  const { report } = await validateFixture(t, ({ plan }) => {
    plan.slides[0].regions[0].targetBbox = { x: 2, y: 1.25, width: 6, height: 1, unit: "in" };
  });
  assert.equal(report.valid, false);
  assert.match(report.errors.join("\n"), /targetBbox requires/u);
});

test("native normalization emits rewrite-and-reposition permission", async (t) => {
  const { report, frameMap } = await validateFixture(t, ({ plan }) => {
    Object.assign(plan.slides[0].regions[0], {
      intent: "style-normalization",
      targetBbox: { x: 2, y: 1.25, width: 6, height: 1, unit: "in" },
    });
  });
  assert.equal(report.valid, true, report.errors.join("\n"));
  assert.equal(frameMap.outputSlides[0].editTargets[0].action, "rewrite-and-reposition");
  assert.ok(Math.abs(frameMap.outputSlides[0].editTargets[0].targetBbox.left - 192) < 0.01);
});

test("role emphasis declarations reject unknown style tokens", async (t) => {
  const { report } = await validateFixture(t, ({ plan }) => {
    plan.styleProfile.roles.title.allowedEmphasis = ["glow"];
  });
  assert.equal(report.valid, false);
  assert.match(report.errors.join("\n"), /allowedEmphasis/u);
});
