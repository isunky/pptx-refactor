---
name: make-pptx-editable
description: Convert image-heavy, flattened, screenshot-based, scanned, or mixed-editability PowerPoint (.pptx) decks into visually faithful, editable, visually consistent PPTX files. Use when a user asks to make slide text or content editable, rebuild picture-based slides, OCR slide content, normalize titles/fonts/bullets, replace screenshots with native PowerPoint text, shapes, tables, or charts, regenerate generic raster icons, or improve editability while preserving the source master and layout. Also use for Chinese requests such as 图片型PPT转可编辑、截图PPT重绘、扫描型PPT、扁平化PPT、内容区域文字可编辑、统一标题字体项目符号、图标重绘、原生元素与位图混合. Do not use for ordinary edits to already-editable decks or for net-new presentations.
---

# Make PPTX Editable

Rebuild picture-heavy `.pptx` decks into faithful, maintainable presentations while stating exactly what is editable. Treat this skill as an orchestration layer over `$presentations`, not as a separate PowerPoint authoring stack.

## Use the required skills and runtime

1. Read and follow `$presentations` before inspecting, editing, or exporting a deck. For an existing PPTX, use its template-following workflow and preserve the master → layout → slide hierarchy.
2. Use the bundled Artifact Tool for presentation import, inspection, editing, rendering, and export. Do not use `python-pptx`.
3. For a visible generic raster icon, first attempt lossless source extraction and cleanup. Read and call `$imagegen` only when it cannot be extracted cleanly; generate one isolated asset per distinct semantic icon. Do not use it for native editable icons or authenticity-sensitive assets.
4. Load bundled workspace dependencies and use their exact absolute Node, Python, module, and binary paths. Do not install global packages.

In command examples, replace `<RUNTIME_NODE_ABSOLUTE_PATH>` with the exact bundled `node.exe` or Node binary returned by the loader; quote it and use the shell's executable invocation syntax when the path contains spaces.

## Define the editability contract

Use **balanced mode** unless the user asks for a different tradeoff. Read [decision-matrix.md](references/decision-matrix.md) before choosing a mode.

- Preserve wording, slide count, order, canvas size, master, layouts, theme, page numbers, headers, and footers by default.
- Rebuild text, lists, cards, tables with recoverable values, simple charts with recoverable data, arrows, timelines, and simple relationship structures as native elements.
- Keep authentic logos, photos, product UI, evidence screenshots, official diagrams, and complex illustrations as independent raster assets.
- Never overwrite the source. Copy it into an isolated working directory, edit the copy, and export a sibling `<stem>_editable.pptx`.
- Do not rewrite content or redesign the deck unless explicitly requested.
- Normalize native and reconstructed objects across the full deck by semantic role. Preserve the distinction between slide titles, section headings, card titles, body text, lists, captions, and callouts.
- Preserve a cleanly extracted generic raster icon by default; regenerate it only when the source crop cannot be cleaned without seams, a tile, or visual loss. Keep an already-native editable icon native; never synthesize an authentic logo, photo, product UI, evidence screenshot, official diagram, or identity-bearing artwork.

Report editability in four separate dimensions:

1. **Text editable** — characters and paragraphs can be edited.
2. **Structure editable** — shapes, connectors, layout, and grouping can be edited.
3. **Data editable** — table cells or chart data can be edited.
4. **Raster replaceable** — an image can be moved, cropped, resized, or replaced, but its pixels are not editable.

Never claim that a deck is “fully editable” without an exception report identifying every retained or regenerated raster asset.

## Follow the conversion workflow

### 1. Preflight the source

- Accept `.pptx` only. Reject legacy `.ppt` and encrypted files.
- Copy the source before opening or modifying it. Record the source SHA-256 and confirm it is unchanged at handoff.
- Detect macros, ActiveX, OLE, SmartArt/diagram parts, audio, video, linked media, animations, transitions, custom shows, equations, Ink, and unsupported vector formats such as EMF/WMF/EPS.
- Stop for explicit confirmation if processing may remove or change an unsupported feature. Otherwise classify it as `manual-review` or preserve it exactly.
- Read [compatibility.md](references/compatibility.md) for platform and feature handling.

### 2. Analyze every slide

Run the active `$presentations` `inspect_template_deck.mjs` first, using the same analysis workspace that will hold the files below. On Windows, invoke it through `<RUNTIME_NODE_ABSOLUTE_PATH> scripts/unzip_compat.mjs --run <inspect_template_deck.mjs> ...`; Node's `spawnSync("unzip")` does not reliably resolve `.cmd` files even when `unzip.cmd` is on `PATH`. Keep the resulting `template-inspect/`, `template-manifest.json`, renders, layouts, media evidence, and contact sheet for the starter-deck workflow.

Run:

```text
<RUNTIME_NODE_ABSOLUTE_PATH> scripts/analyze_hybrid_deck.mjs --pptx <source.pptx> --workspace <analysis-dir>
```

Expect at least:

- `source-manifest.json`
- `theme-profile.json`
- `source-inspect.ndjson`
- `source-renders/`
- `source-layout/`
- `source-media/`
- `source-montage.webp` plus `source-montage.json`, proving that every rendered slide appears exactly once

Use image coverage, native-text density, repeated-image hashes, masters, layouts, placeholders, fonts, crops, and special-object evidence only as a first-pass signal. Inspect each rendered slide at full size; the montage is only for pacing and consistency.

Use the analyzer's text-style clusters, role candidates, component candidates, and icon candidates as evidence. Confirm them visually before assigning final roles or asset classes.

Classify every slide as `native-editable`, `mixed`, `flattened`, or `low-quality-scan`. Do not let a coverage ratio replace visual judgment.

### 3. Inventory and transcribe

- Assign each relevant source object to exactly one disposition: keep native, reconstruct, preserve raster, replace asset, or manual review.
- Keep existing native text when it is correct and usable.
- Transcribe raster text visually and create a normalized text ledger. Preserve intentional wording, punctuation, units, acronyms, and line grouping.
- Mark uncertain digits, units, proper nouns, and unreadable fragments for review. Never guess.
- Treat one logical paragraph or list item as one text block. Do not create one textbox per source image line.
- Record all retained raster assets and all generated assets with slide, source, reason, prompt if applicable, and editability level.
- Classify each image-bearing region as `generic-icon`, `logo`, `photo`, `product-ui`, `evidence-screenshot`, `official-diagram`, `complex-illustration`, or `other`. In the default icon policy, every visible `generic-icon` raster region must use `extract-raster` with clean extraction evidence or `regenerate-icon` when clean extraction is not possible.
- For a flattened or full-page composite screenshot, make a **visual-asset inventory before deleting the parent image**. Inventory every visible logo, photo, background illustration, official diagram, product UI, and generic icon with its source bbox, class, and intended disposition. An explicit empty inventory is allowed only with `complete: true` and an evidence-based `emptyReason`; “the entire screenshot will be rebuilt” is never an asset disposition.
- A generic icon embedded in a raster screenshot must be losslessly cropped or regenerated as a transparent single icon; do not replace it with a Unicode glyph, a generic letter, or an unrelated native symbol. `rebuild-native` is reserved for genuinely simple non-icon visuals and must be explicitly justified.
- When the user supplies red-box or other annotated feedback, create an issue for every marked region, bind it to exact slide/object IDs, and set its propagation scope to local, component-family, or deck-wide.

### 4. Write a conversion plan

Create a schema `1.1` `conversion-plan.json` using [manifest-schema.md](references/manifest-schema.md). Read [visual-consistency.md](references/visual-consistency.md) and declare `visualPolicy`, `styleProfile`, and `calibration`. Every region must include:

- `slideNumber`
- `bbox`
- `sourceObjectIds`
- `action`
- `targetType`
- `expectedText`
- `confidence`
- `reason`
- `intent`
- `styleRole` and/or `componentFamily` when the region participates in normalization

Allowed actions are exactly:

```text
keep-native
rebuild-text
rebuild-shape
rebuild-table
rebuild-chart
extract-raster
regenerate-icon
retain-raster
manual-review
```

Require a reason for `retain-raster`, `regenerate-icon`, and `manual-review`. Bind every destructive edit to exact source inspect IDs. Never use a wildcard or a rule equivalent to “delete all images on this slide.”

Keep the nine actions unchanged. Use `intent: faithful-rebuild`, `style-normalization`, or `user-approved-redesign` to distinguish a faithful reconstruction from normalization or a user-authorized redesign. A redesign intent requires an exact scope, authorization note, and `preserveText: true`.

For every `retain-raster`, `extract-raster`, or `regenerate-icon` region, add a matching `assets[]` provenance record with the exact source IDs, source media SHA-256, source path or media reference, reason, and `raster-replaceable` label. A generated icon record also needs its prompt and provenance; QA later binds the output path and hash.

For generated icons also record `assetClass`, `semanticConcept`, `iconFamilyId`, and the shared family style prompt. Generate the first accepted icon as the family anchor and use the same family specification for subsequent prompts. One initial generation plus one defect-specific retry is the limit; a second failure is a blocker, not permission to retain the old generic icon silently.

Validate the plan:

```text
<RUNTIME_NODE_ABSOLUTE_PATH> scripts/validate_conversion_plan.mjs \
  --manifest <analysis-dir>/source-manifest.json \
  --plan <conversion-plan.json> \
  --out-map <template-frame-map.json>
```

Do not edit the deck until this passes. The resulting map must account for every large slide-local raster and every reconstruction region exactly once.
It must also contain a `$presentations`-compatible `outputSlides` view with exact inherited IDs and bounded `action: "add"` permission, plus the conversion-specific `slides` view. Validate the same map with the active `$presentations` template-plan validator before building the starter.

If one composite screenshot contains multiple semantic areas, bind and delete it once with one full-footprint parent reconstruction region. Put the separate text, bullet, table, chart, or icon decisions in validated nested `subregions`; do not repeat the same source image across several top-level destructive dispositions.
Its parent must also contain a complete `visualAssets` inventory. Each visible asset has one disposition: retain/extract as raster, regenerate as a generic icon, explicitly rebuild as a simple non-icon native visual, or manual review. Preserve authenticity-sensitive content as raster; do not make it disappear merely because the parent screenshot is deleted.

### 5. Prepare the template-preserving starter

- Use the source copy as the starter deck.
- Preserve inherited chrome on the master or layout instead of duplicating it on every slide.
- Build an exact template-frame map from source slide to source slide by default.
- Delete only the exact slide-local object IDs named in the validated plan.
- Permit new objects only inside the mapped reconstruction zone, except for intentional slide-wide structural updates recorded in the plan.
- Full-deck role normalization may update existing native text or repeated component geometry only when exact object IDs, roles, and component families are listed in the plan. Do not use a broad slide selector.
- Run the active `$presentations` template-plan validator against `outputSlides`, then run its `prepare_template_starter_deck.mjs` with the same inspected workspace, source PPTX, and validated map. On Windows set `PYTHON` to the exact bundled Python path and `PYTHONUTF8=1` when requesting its contact sheet. Import and edit the resulting starter PPTX; do not build replacement slides from scratch.

### 6. Calibrate and freeze the visual profile

- If the user explicitly asks to try one page first, use `calibration.mode: user-gated`. Build and render only the representative sample, wait for approval, then record `status: approved` and freeze the approved tokens before authoring the full deck.
- Otherwise use `calibration.mode: automatic`. Select up to three distinct representative slides covering the densest list, a repeated component family, and an icon-rich layout. Render them, resolve hard findings, record the evidence, set `status: complete`, and continue without pausing.
- Once calibration is approved or complete, do not make ad-hoc per-slide style changes. Add an explicit role or family exception with a reason when the content truly requires one.

### 7. Rebuild content

Read [reconstruction-rules.md](references/reconstruction-rules.md) before authoring.

- Infer font families, effective sizes, weights, colors, insets, alignment, line spacing, paragraph spacing, line widths, radii, shadows, and spacing from repeated source evidence. Do not hardcode a font or palette.
- Convert source point sizes explicitly when the target API uses pixels: `px = pt × 96 / 72`.
- Use `scripts/rebuild_helpers.mjs` for theme-injected titles, text blocks, cards, connectors, lists, and transparent icons. Pass style tokens from the source profile.
- Build connectors before nodes. Prefer native shapes for simple diagrams, arrows, timelines, cards, and separators.
- Rebuild tables and charts natively only when values are recoverable with high confidence. Preserve them as raster exceptions when data cannot be established reliably.
- Re-render after any text, size, spacing, crop, or alignment change.
- Apply the frozen role profile to both native and reconstructed text. Derive one canonical style per semantic role from master/layout evidence and the dominant valid source cluster; do not preserve accidental source outliers.
- Give every normalized object a semantic name using `mppe|role=<role>|family=<family>|instance=<id>|part=<part>`. Use lowercase ASCII identifiers so QA can compare peers after round-trip.

### 8. Handle bullets correctly

- Prefer structured native paragraphs with a bullet character, a positive left margin, and a negative hanging indent.
- If compatibility is insufficient or the source requires a visibly heavy dot, create one native solid-circle shape and one textbox per logical list item.
- Never type a literal `•` into the body text to imitate a bullet.
- Start with adaptive metrics relative to body font size: dot diameter `0.38–0.45`, dot-to-text gap `0.55–0.75`, within-item line spacing `1.05–1.15`, and between-item spacing `0.4–0.6`. Confirm all values in rendered output.
- Wrapped lines must align with the first character of the item text, not with the bullet.
- Name independent dots and textboxes with the same list instance and item index. Treat helper-reported overflow as a hard failure.

### 9. Handle raster assets and icons

Use lossless extraction and source cropping for authenticity-sensitive rasters and as visual evidence for generic-icon regeneration. Normalize a retained/extracted source asset when needed:

```text
<RUNTIME_NODE_ABSOLUTE_PATH> scripts/prepare_raster_asset.mjs \
  --input <source-image> \
  --output <clean.png> \
  --mode auto \
  --report <asset-report.json>
```

Use `auto`, `chroma`, or `edge` according to the source. Preserve interior whites, use containment rather than cover-cropping, trim transparent margins, and reject visible seams or opaque rectangular tiles.

For every visible generic raster icon by default:

- First crop and clean the source icon. Keep it as `extract-raster` when the result has clean alpha, no visible tile/seam, and adequate effective resolution.
- If extraction is not clean, invoke `$imagegen` once per distinct icon and request genuine transparent alpha.
- Generate a single isolated icon, not an atlas or sprite sheet.
- Request no text, watermark, frame, tile, shadow, circular backing, or decorative container.
- Do not regenerate a logo, person, product UI, official diagram, screenshot, or evidence image.
- Copy the final asset from `CODEX_HOME` into the workspace and record the source, prompt, output path, and selection reason.
- Draw any supporting circle, card, outline, or shadow as a native PowerPoint shape to avoid double backings.
- Use the same family anchor, palette, stroke-weight class, detail level, optical coverage, and composition constraints for every icon in the deck.
- Preserve a native editable icon instead of replacing it with a raster. If extraction fails and ImageGen is unavailable, stop and report that the fallback icon contract cannot be completed; use a CLI fallback only after the user explicitly chooses it under `$imagegen` rules.

For composite screenshots, preserve or replace every inventory entry before removing the parent screenshot. The final slide must contain a verifiable raster at each retained/extracted/regenerated asset bbox; record the source crop, output file hash, and final embedded-media binding. Never accept a text-and-shape-only result when the inventory includes visual assets.

### 10. Validate, round-trip, and inspect

Run:

```text
<RUNTIME_NODE_ABSOLUTE_PATH> scripts/qa_conversion.mjs \
  --source <source.pptx> \
  --final <output_editable.pptx> \
  --plan <conversion-plan.json> \
  --workspace <qa-dir>
```

Set the plan's `sourceManifest` to the analyzer output, or copy `source-manifest.json` beside the plan or into the QA workspace. Do not accept a QA run that cannot resolve this manifest, because it cannot prove complete large-raster accounting.

Also run the current `$presentations` template-fidelity checks, slide tests, per-slide layout checks, and rendering workflow. On Windows, run Node-based checks that call `spawnSync("unzip")` through `scripts/unzip_compat.mjs --run`; use `scripts/unzip.cmd` for direct shell-style `unzip -Z1` or `unzip -p` calls.

Read [qa-checklist.md](references/qa-checklist.md) and resolve all hard failures. The QA report must bind its findings to the exact final PPTX SHA-256, perform import/export/re-import, render every slide at full size, and separate slide-local objects from inherited layers.

Apply tiered visual QA. Deterministic defects such as clipping, role-style drift, component-family geometry drift, bullet geometry errors, unresolved feedback issues, opaque icon tiles, atlas seams, invalid support layers, and image occlusion are hard failures. Subtle stroke, palette, shadow, or renderer differences are warnings requiring full-size review.

For every feedback issue, fix the marked object, inspect all instances in the same component family, scan the deck for the same defect class, and re-render the affected slide plus its peers. Final QA fails while any issue remains `open`; a `waived` issue requires an evidence-based reason.

### 11. Deliver with an exception ledger

Deliver the new `_editable.pptx` plus:

- conversion plan
- asset/provenance manifest
- retained-raster and manual-review exceptions
- QA report and human-readable ledger
- visual issue ledger, calibration evidence, and role/component consistency summary
- source and final SHA-256 values

Summarize the four editability levels and any remaining caveats. Confirm that the source hash is unchanged and the QA hash matches the delivered file.

## Use the bundled scripts

- `scripts/analyze_hybrid_deck.mjs` — analyze, render, inspect, inventory, and classify a source deck.
- `scripts/validate_conversion_plan.mjs` — validate exact object coverage and emit a safe frame map.
- `scripts/rebuild_helpers.mjs` — provide style-injected native reconstruction helpers.
- `scripts/prepare_raster_asset.mjs` — clean and normalize one raster asset without damaging interior whites.
- `scripts/qa_conversion.mjs` — run editability, structure, visual, round-trip, and hash-bound QA.
- `scripts/unzip_compat.mjs` and `scripts/unzip.cmd` — support the subset of `unzip` used by presentation inspection scripts on Windows.

Keep semantic transcription, object grouping, aesthetic reconstruction, icon selection, and full-size visual review as judgment tasks. Do not turn the workflow into an unattended one-click raster-to-layout algorithm.
