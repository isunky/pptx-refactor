# Targeted Palette Follow-ups

Use this path when the user asks only to harmonize colors on a reconstructed deck, or explicitly invokes this skill for a palette-only edit. Use the latest user-supplied revision as the source. A previously generated deck or an old builder is not a substitute for that revision.

## Choose the palette by role

Inspect every source slide before choosing changes. Keep the existing theme and distinguish text emphasis, headings, module fills, borders, pale background tints, and ordinary icons. Apply an explicit user RGB value exactly to the requested role. Do not apply that RGB to every orange object: a suitable highlight can be unsuitable for a large background or a white-on-color heading bar.

- Keep headings and body copy in the established theme family.
- Use the accent sparingly for emphasis; inspect small text and tinted backgrounds at presentation size.
- Bring generic icons into the same hue family while retaining their recognizable symbols and legibility. Picture color effects preserve the bitmap and can be changed in PowerPoint.
- Preserve authentic logos, official documents, photos, product interfaces and evidence screenshots in their original colors unless the user specifically authorizes changes to those assets.
- Preserve content, typography, geometry, tables, notes and slide order during a palette-only edit. Do not introduce writing changes or reconstruct already-editable slides.

## Bounded package helper

`scripts/apply_palette.py` uses Python's standard library. Run it with the exact bundled Python executable. It replaces only selected RGB attributes and explicit picture-effect elements; it does not resave the deck through a different authoring engine or alter embedded image bytes. This is a narrow package-preserving helper, not a replacement for `$presentations` for reconstruction or content authoring. Follow the active presentation workflow for final rendering and validation.

Create a private JSON plan. The following colors are illustrative, not defaults:

```json
{
  "schemaVersion": "1.0",
  "colorRules": [
    {"from": "EC671B", "to": "FCB824", "scope": "text", "parts": ["ppt/slides/slide2.xml"]},
    {"from": "EC671B", "to": "146CB5", "scope": "shape", "parts": ["ppt/slides/slide2.xml"]}
  ],
  "iconEffects": [
    {"part": "ppt/slides/slide2.xml", "shapeId": "17", "assetClass": "generic-icon", "dark": "0064AA", "light": "FFFFFF"}
  ],
  "protectedPictures": [
    {"part": "ppt/slides/slide2.xml", "shapeId": "18"}
  ]
}
```

Use package part names and native `p:cNvPr/@id` values, not imported-tool object IDs. Determine them from the actual source package. Part names are not necessarily display slide numbers; inspect `ppt/presentation.xml` and its relationships when slide order differs. Every rule lists exact slide parts. Split rules when a color is absent from one selected slide; each declared rule must match on every listed part. Source colors map simultaneously, so `A -> B` and `B -> C` do not turn an original A into C.

```text
<RUNTIME_PYTHON_ABSOLUTE_PATH> scripts/apply_palette.py --source <latest.pptx> --output <new_palette.pptx> --plan <palette-plan.json> --report <dry-run.json> --dry-run
<RUNTIME_PYTHON_ABSOLUTE_PATH> scripts/apply_palette.py --source <latest.pptx> --output <new_palette.pptx> --plan <palette-plan.json> --report <applied.json>
```

Use new report paths for each run. Dry-run writes the report only. The helper refuses source/output collisions, existing outputs/reports, invalid hex colors, missing parts, duplicate/conflicting rules, unmatched colors, unresolved picture selectors, linked images, existing picture effects/extensions and digitally signed packages. `protectedPictures` checks that each protected picture exists and cannot also appear in `iconEffects`. Other pictures remain unchanged automatically.

The helper covers explicit slide RGB colors only. Scheme colors, inherited master/layout styles, charts, SmartArt and embedded vectors are outside its color-rule scope. Do not claim that the entire palette is normalized when those features remain. Do not convert scheme colors to RGB or recolor evidence just to eliminate a warning. An icon's `assetClass` is the operator's classification, not an automated authenticity test; inspect the asset first. Duotone can reduce contrast or change shadows, so compare actual PowerPoint renders before accepting it. Use a supported image-editing workflow if the native effect is unsuitable.

## Validate and hand off

1. Review the dry-run part/role counts against the intended scope.
2. Apply once to a new output. Keep the source and plan unchanged.
3. Use presentation package checks and inspect every rendered final slide, including all instances of changed icon families. Compare with source renders for accidental changes. Prefer PowerPoint rendering when available; claim PowerPoint verification only after actually opening and inspecting it.
4. Match the report's final SHA-256 to the delivered file. Deliver the PPTX; keep plans and validation records private unless requested.

A reconstruction conversion plan and the full reconstruction QA pipeline are not required for this bounded follow-up: those checks assume objects are being rebuilt. The palette report verifies the limited byte changes; it does not establish visual quality, factual correctness, or a successful PowerPoint open.
