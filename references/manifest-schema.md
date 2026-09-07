# Manifest and Plan Schema

Keep manifests deterministic, UTF-8 encoded, and path-portable. Store paths relative to the analysis or delivery workspace when possible. Use slide numbers as one-based integers and bounding boxes in a declared coordinate system.

### Validator coordinate handling

- Schema `1.0` and `1.1` manifest `bbox` values are canonical inches; pixel measurements appear only in suffixed evidence fields such as `bboxPx` and `renderDimensionsPx`.
- A plan `bbox` must explicitly declare `px` or `in`; unitless boxes are invalid. The validator converts plan and canonical manifest boxes to the slide pixel canvas using the explicit render dimensions and physical slide size. If pixel evidence is unavailable, use 96 px/in.
- Do not compare raw numbers from different units. `template-frame-map.json` always emits `dimensions`, region boxes, source-object boxes, and add zones in `px`.

## Source manifest

`analyze_hybrid_deck.mjs` writes `source-manifest.json`. At minimum it should expose:

```json
{
  "schemaVersion": "1.1",
  "sourcePptx": "source.pptx",
  "sourceSha256": "...",
  "slideSize": { "width": 13.333, "height": 7.5, "unit": "in" },
  "slideCount": 12,
  "themeProfile": "theme-profile.json",
  "specialObjects": [],
  "slides": [
    {
      "slideNumber": 1,
      "classification": "mixed",
      "imageCoverageRatio": 0.63,
      "nativeTextCount": 2,
      "objects": [
        {
          "id": "im/7",
          "kind": "image",
          "bbox": { "x": 0.8, "y": 1.4, "width": 11.7, "height": 5.3, "unit": "in" },
          "slideLocal": true,
          "sha256": "...",
          "coverageRatio": 0.62
        }
      ]
    }
  ]
}
```

Object IDs must be the stable inspect IDs used by the editing tool. Do not substitute shape names, ordinal positions, or inferred labels when an inspect ID exists.

### Coordinate contract

Schemas `1.0` and `1.1` use inches as the canonical manifest coordinate system. The following fields must all use `{ "unit": "in" }` and the same top-left origin:

- `slideSize` and `deck.slideSize`;
- `slides[].dimensions`;
- `slides[].objects[].bbox`;
- top-level `objects[].bbox`;
- `largeRasters[].bbox`.

Canonical boxes use `x`, `y`, `width`, and `height`. A consumer may also accept `left`/`top` aliases, but an analyzer must not emit the same logical object in different units at different manifest locations. Pixel measurements inside the manifest are render evidence only and must be placed in explicitly suffixed fields such as `renderDimensionsPx`, `renderSlideSizePx`, `canvasPx`, or `bboxPx`. A conversion plan may declare a supported unit such as `in` or `px`; the validator converts it against the actual slide canvas before comparison and map output. The top-level `coordinateSystem` record states the canonical manifest contract explicitly.

PowerPoint may encode a horizontal line or connector with a zero-height raw box, or a vertical one with a zero-width raw box. For line-like objects only, canonical `bbox` is an effective visible box: expand each zero axis by the declared stroke width, with a one-pixel minimum, then convert that box to inches. Preserve the exact source evidence in `rawBboxPx` and describe the adjustment in `bboxAdjustment`; `strokeWidthPx` records the supporting measurement. Do not apply this expansion to ordinary shapes, images, text boxes, tables, or charts. This keeps plan intersection and bounds validation meaningful without falsifying the original geometry.

### Special-object scope and review state

Each present entry in `specialObjects` and `deck.specialFeatures` declares:

- `objectCount` separately from relationship/evidence counts;
- `scope`: `active-slide-dependency`, `unused-template-metadata`, or `package-metadata`;
- `activeSlides`, `activeLayouts`, `activeMasters`, `unusedLayouts`, and `unusedMasters` where applicable;
- `blocking`: an active dependency that prevents automatic reconstruction;
- `manualReview`: content or metadata that requires an explicit preservation decision even when it is not active-slide blocking.

`blockingFeatures` contains only active/package-level blockers. `manualReviewFeatures` also includes non-blocking notes, comments, author history, tags, transitions, unused-layout OLE, and similar preservation risks. `requiresManualCompatibilityReview` is true when either class requires review. Do not infer that a slide depends on every sibling layout merely because those layouts share an active master.

### Verified contact sheet

When `source-montage.webp` is present, `source-montage.json` and `artifacts.montage` must prove that every rendered slide occurs exactly once. The analyzer records source-render hashes, grid positions, tile boxes, output hash, and a lossless per-tile pixel verification result. If any slide render is missing or verification fails, both montage files must be absent and the manifest must contain an explicit warning; never advertise a first-slide-only export as a montage.

## Conversion plan

Use this top-level form:

```json
{
  "schemaVersion": "1.1",
  "mode": "balanced",
  "sourcePptx": "source.pptx",
  "sourceManifest": "analysis/source-manifest.json",
  "outputPptx": "source_editable.pptx",
  "sourceSha256": "...",
  "visualPolicy": {
    "normalizationScope": "full-deck-role-based",
    "iconMode": "extract-or-regenerate-generic",
    "qaStrictness": "tiered",
    "calibrationMode": "automatic"
  },
  "styleProfile": {
    "roles": {},
    "componentFamilies": {},
    "iconFamily": {}
  },
  "calibration": {
    "mode": "automatic",
    "representativeSlides": [1],
    "status": "complete",
    "evidence": [],
    "frozenProfileSha256": "..."
  },
  "feedbackIssues": [],
  "assets": [],
  "slides": [
    {
      "slideNumber": 1,
      "sourceSlide": 1,
      "classification": "mixed",
      "regions": [
        {
          "regionId": "s01-body",
          "slideNumber": 1,
          "bbox": { "x": 0.8, "y": 1.4, "width": 11.7, "height": 5.3, "unit": "in" },
          "sourceObjectIds": ["im/7"],
          "action": "rebuild-text",
          "targetType": "text+shape",
          "expectedText": ["Editable heading", "Editable body text"],
          "confidence": 0.98,
          "reason": "Body is a flattened composite; wording is legible and confirmed.",
          "intent": "style-normalization",
          "styleRole": "body",
          "componentFamily": "intro-card",
          "editability": ["text-editable", "structure-editable"]
        }
      ]
    }
  ]
}
```

Set `sourceManifest` to the absolute path or a path relative to the conversion-plan file. QA uses it to verify source hashes, exact object IDs, and complete large-raster accounting. If omitted, place `source-manifest.json` beside the plan or in the QA workspace; QA must fail when no source manifest can be resolved.

The validator accepts legacy `1.0` manifests/plans for compatibility. New work produced by this skill uses `1.1`. A `1.0` plan may omit the visual fields and receives compatibility warnings rather than invented defaults; it does not claim the new visual-consistency contract.

### Visual policy and frozen profile

Schema `1.1` requires:

- `visualPolicy.normalizationScope`: `full-deck-role-based`.
- `visualPolicy.iconMode`: `extract-or-regenerate-generic`.
- `visualPolicy.qaStrictness`: `tiered`.
- `visualPolicy.calibrationMode`: `automatic` or `user-gated`, matching `calibration.mode`.
- `styleProfile.roles`: canonical style records keyed by lowercase role IDs.
- `styleProfile.componentFamilies`: repeated component definitions keyed by lowercase family IDs.
- `styleProfile.iconFamily`: family ID, shared prompt prefix, construction/style attributes, target optical coverage (default `0.72`), centroid tolerance (default `0.04`), and optical tolerance (default `0.08`).
- `calibration`: mode, representative slide numbers, status, evidence, and `frozenProfileSha256`. A calibration-stage plan may use `pending` without evidence or a frozen hash. A final-stage plan requires `complete` for automatic calibration or `approved` for user-gated calibration, rendered evidence, and a hash equal to the SHA-256 of recursively key-sorted, whitespace-free JSON for `styleProfile`.

Role records declare the tokens that QA can inspect: typeface, font size in pixels, bold/weight, normalized color, alignment, line spacing, insets, and `maxLines`. They may declare `allowedEmphasis` using `bold`, `italic`, `underline`, or `color`; QA applies canonical tokens to every non-empty run and permits only these listed differences. Component-family records declare expected relative part boxes or anchors, flexible axes, and an optional tighter tolerance. Object names use `mppe|role=<role>|family=<family>|instance=<id>|part=<part>` with lowercase ASCII IDs.

Validate a draft and emit a representative-slide-only map with `--stage calibration`. After inspecting the sample, update calibration evidence and the frozen profile hash, then rerun with `--stage final`. Omitting `--stage` remains equivalent to `final`.

The analyzer may emit style clusters and role/component/icon candidates. These are evidence, not permission to infer uncertain semantics.

The validator's explicit `--manifest` argument is authoritative. If `sourceManifest` is present, it must resolve to that exact file; omission is reported as a portability warning. `sourceSha256` is mandatory and must match the manifest. `sourcePptx` must identify a manifest-recorded source file (a portable bare filename is allowed when it matches a recorded source basename and hash). `outputPptx` is mandatory, must be a different `.pptx` path, and must end in `_editable.pptx`.

Every region requires:

- `slideNumber`: one-based integer matching its containing slide.
- `bbox`: finite, non-negative `x`, `y`, `width`, and `height`; declare a supported unit.
- `targetBbox` (optional): the final output region when role-based normalization intentionally moves or resizes content. `bbox` remains the exact source/deletion footprint; both boxes must stay inside the slide.
- `sourceObjectIds`: non-empty exact inspect IDs unless the action is a recorded new construction in a bounded zone.
- `action`: one of the fixed values below.
- `targetType`: expected output object category.
- `expectedText`: string or array; use an empty array only when no text is expected.
- `confidence`: number from `0` through `1`.
- `reason`: concise evidence-based rationale.
- `editability`: non-empty array using only `text-editable`, `structure-editable`, `data-editable`, and `raster-replaceable`.
- `intent` in schema `1.1`: `faithful-rebuild`, `style-normalization`, or `user-approved-redesign`.
- `styleRole` and/or `componentFamily` when the output participates in a frozen role or family; each ID must exist in `styleProfile`.
- `issueRefs` when the region resolves annotated feedback.

`user-approved-redesign` additionally requires `authorization`, `preserveText: true`, and an exact bbox/object binding. It changes visual-difference accounting only inside that region; it does not weaken text, brand, bounds, component, or round-trip checks.

### Annotated feedback issues

Schema `1.1` `feedbackIssues[]` entries contain:

- `issueId`, `slideNumber`, annotation source, annotation bbox, and exact `relatedObjectIds`;
- `category`, `scope` (`local`, `component-family`, or `deck`), and expected fix;
- `componentFamily` when scope is component-family;
- `status`: `open`, `fixed`, or `waived`;
- `inspectedInstanceIds` for family/deck propagation;
- evidence paths and a non-empty waiver reason when status is `waived`.

Plan validation accepts `open` issues while authoring is still planned. Final QA treats every `open` issue as an error and requires family/deck issues to enumerate all planned peer instances inspected.

`targetType` is a `+`, `/`, `,`, or `|` separated list of these components: `text`, `shape`, `group`, `connector`, `image`, `icon`, `table`, `chart`, `object`. Plural connector/shape/image/icon forms and `native-table`, `native-chart`, `native-object`, or `preserved-object` normalize to the singular components. Apply these action contracts:

| Action | Target contract | Required editability |
|---|---|---|
| `keep-native` | Any declared component | Labels implied by the target |
| `rebuild-text` | Must include `text`; may also use shape/group/image/icon | `text-editable`, plus labels implied by other components |
| `rebuild-shape` | Must include shape/group/connector | `structure-editable`, plus text/raster labels when present |
| `rebuild-table` | Must include `table` | `text-editable`, `structure-editable`, `data-editable` |
| `rebuild-chart` | Must include `chart` | `structure-editable`, `data-editable` |
| `extract-raster` / `regenerate-icon` | image/icon only | `raster-replaceable` |
| `retain-raster` | Must include `image` | `raster-replaceable` |
| `manual-review` | Any declared component | Declare the expected post-review level; no automatic deletion is authorized |

All actions require a non-empty reason. Destructive actions (`rebuild-*`, `extract-raster`, `regenerate-icon`) require confidence `>= 0.80`; a lower-confidence region or an uncertain text-ledger entry must use `manual-review`.

Every `retain-raster`, `extract-raster`, or `regenerate-icon` region must match at least one `assets[]` record by slide plus `regionId`, `assetId`, or exact `sourceObjectIds`. The record must contain a unique `assetId`, matching action, source object IDs, source media SHA-256, reason, and `raster-replaceable` editability. Retained/extracted assets require `inputPath` or `mediaRef`; regenerated icons additionally require the planned prompt and provenance. After authoring, QA also requires generated/extracted output paths and hashes.

Schema `1.1` raster asset records also require `assetClass`: `generic-icon`, `logo`, `photo`, `product-ui`, `evidence-screenshot`, `official-diagram`, `complex-illustration`, or `other`. A `generic-icon` must use `regenerate-icon`, include `semanticConcept`, `iconFamilyId`, and `familyStylePrompt`, and match `styleProfile.iconFamily`. Non-generic authenticity classes must not use `regenerate-icon`.

For a destructive region, the normalized box must cover at least 95% of every bound source object. Its combined source footprint must also reach IoU `>= 0.80`, or be fully contained by a replacement region no more than 1.25× the footprint area. Mere intersection never authorizes deletion.

When one composite screenshot contains several semantic areas, consume the source object exactly once with one full-footprint parent region, then describe output-only `subregions` inside it. Each subregion repeats the normal required fields and exact parent `sourceObjectIds`, but it is marked as a semantic child by the validator and never authorizes a second deletion. Use subregion actions only for `rebuild-text`, `rebuild-shape`, `rebuild-table`, `rebuild-chart`, `extract-raster`, `regenerate-icon`, or `manual-review`. The frame map emits one delete and one bounded parent add permission plus the nested semantic zones. Do not model the same source image as several top-level destructive dispositions.

### Composite screenshot visual-asset inventory

When a schema `1.1` parent reconstruction deletes a large raster on a `flattened` or `low-quality-scan` slide, the parent region must contain `visualAssets`:

```json
{
  "complete": true,
  "items": [
    {
      "assetId": "s01-icon-policy",
      "bbox": { "x": 62, "y": 302, "width": 62, "height": 62, "unit": "px" },
      "sourceObjectIds": ["im/7"],
      "assetClass": "generic-icon",
      "disposition": "regenerate-icon",
      "reason": "The cropped source icon has a non-removable white tile."
    },
    {
      "assetId": "s01-header-illustration",
      "bbox": { "x": 0, "y": 0, "width": 250, "height": 120, "unit": "px" },
      "sourceObjectIds": ["im/7"],
      "assetClass": "complex-illustration",
      "disposition": "extract-raster",
      "reason": "Identity-bearing header artwork must remain visually faithful."
    }
  ]
}
```

Every item needs a unique lowercase ASCII `assetId`, a bbox fully inside the parent region, the parent’s exact source IDs, an allowed asset class, a disposition, and a non-empty reason. Allowed dispositions are `retain-raster`, `extract-raster`, `regenerate-icon`, `rebuild-native`, and `manual-review`.

- `generic-icon` from a raster composite must use `extract-raster` with `extractionQuality: "clean"`, or `regenerate-icon` when a clean crop cannot be achieved. Either disposition needs a matching `assets[]` record and one final transparent image; Unicode, emoji, or a letter substituted for the icon is not compliant.
- Logos, photos, product UI, evidence screenshots, official diagrams, and complex illustrations must use `retain-raster`, `extract-raster`, or `manual-review`; they cannot be regenerated or silently replaced by shapes.
- `rebuild-native` is allowed only for a simple non-icon asset classed as `other` and needs a specific reason.
- An empty inventory is allowed only as `{ "complete": true, "items": [], "emptyReason": "..." }` after full-size visual inspection.

Each retained/extracted/regenerated inventory entry must have a matching `assets[]` record by `assetId` and action. QA verifies its source binding, output hash, embedded final media, and region presence. This inventory is independent from semantic `subregions`: it protects visual assets while subregions protect text and structure.

Allowed action enum:

```json
[
  "keep-native",
  "rebuild-text",
  "rebuild-shape",
  "rebuild-table",
  "rebuild-chart",
  "extract-raster",
  "regenerate-icon",
  "retain-raster",
  "manual-review"
]
```

`retain-raster`, `regenerate-icon`, and `manual-review` require a non-empty reason. For `regenerate-icon`, also reference an asset record with the generation prompt and provenance.

Reject:

- wildcards, regex selectors, slide-wide delete rules, or empty source-object selectors;
- object IDs absent from the source manifest;
- regions outside the declared slide canvas;
- overlapping destructive regions unless the overlap is explicitly modeled as one region;
- the same source object consumed by multiple destructive actions;
- an unclassified large slide-local raster;
- high-risk or low-confidence text reconstruction without a review state;
- retained/generated assets without a rationale and editability label.
- a destructive region that only clips or incidentally intersects the object it would delete;
- a plan whose source manifest, source PPTX, or SHA-256 does not match, or whose output path overwrites a source.

## Template frame map

`validate_conversion_plan.mjs` emits one deterministic `template-frame-map.json` with two synchronized views:

- `outputSlides`: the `$presentations` template-following contract consumed by its plan validator and starter-deck helper;
- `slides`: conversion-specific exact deletions, bounded add zones, and region accounting.

The template-following view uses exact inherited IDs and explicit permission for primitives added only inside a deleted composite-image zone:

```json
{
  "outputSlides": [
    {
      "outputSlide": 1,
      "sourceSlide": 1,
      "narrativeRole": "editable content reconstruction",
      "reuseMode": "duplicate-slide",
      "editTargets": [
        { "action": "delete", "sourceElementIds": ["im/7"], "regionId": "s01-body", "reason": "Remove the exact flattened body image." },
        { "action": "add", "newPrimitiveAllowed": true, "mustNotOverlapInherited": true, "zone": { "left": 76.8, "top": 134.4, "width": 1123.2, "height": 508.8, "unit": "px" }, "regionId": "s01-body", "targetType": "text+shape", "reason": "Rebuild only inside the deleted body-image zone." }
      ]
    }
  ]
}
```

The conversion-specific view states:

```json
{
  "slideNumber": 1,
  "sourceSlide": 1,
  "deleteObjectIds": ["im/7"],
  "preserveObjectIds": ["sh/2", "tx/3"],
  "addZones": [
    {
      "regionId": "s01-body",
      "bbox": { "left": 76.8, "top": 134.4, "width": 1123.2, "height": 508.8, "unit": "px" },
      "allowedTargetTypes": ["text", "shape", "image"]
    }
  ]
}
```

For preserved non-empty placeholders, emit an explicit same-content `rewrite` declaration. For empty structural placeholders, emit an exact-ID `delete` declaration so PowerPoint does not expose an unresolved authoring prompt. Never use this compatibility view to weaken the conversion plan's exact-ID and bounded-zone rules.

Default to source slide N → output slide N. A different mapping requires an explicit reason.

## Asset/provenance manifest

Track each retained, extracted, cleaned, or generated asset:

```json
{
  "assetId": "s01-icon-01",
  "slideNumber": 1,
  "action": "regenerate-icon",
  "sourceObjectIds": ["im/8"],
  "inputPath": "source-media/image8.png",
  "outputPath": "assets/s01-icon-01.png",
  "sourceSha256": "...",
  "outputSha256": "...",
  "reason": "Source icon contains unrecoverable atlas seams.",
  "assetClass": "generic-icon",
  "semanticConcept": "database compatibility",
  "iconFamilyId": "main-icons",
  "familyStylePrompt": "consistent flat blue-green outline icon family; transparent background; no backing",
  "prompt": "single isolated generic database icon, transparent background, no text, no frame",
  "provenance": "OpenAI ImageGen",
  "editability": "raster-replaceable"
}
```

For non-generated assets, omit `prompt` but retain source path, hashes, transformation report, and reason.

QA resolves relative `inputPath` and `outputPath` values against the conversion-plan directory. It recomputes `sourceSha256` from every declared input file and/or `mediaRef` inside the exact source PPTX package, binds that hash to the region's exact `sourceObjectIds`, recomputes `outputSha256` from the declared output file, and confirms the same output bytes are embedded in the final PPTX at a substantially matching region. A syntactically valid but unverified hash is not provenance.

## Text ledger normalization

Normalize only for comparison; do not silently alter output wording.

- Apply Unicode normalization consistently.
- Normalize presentation-only whitespace while retaining paragraph and list-item boundaries.
- Preserve numbers, dates, units, acronyms, punctuation, and case.
- Store uncertain readings explicitly, for example `{ "text": "12.5 GB", "confidence": 0.72, "needsReview": true }`.
