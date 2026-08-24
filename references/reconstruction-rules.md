# Reconstruction Rules

Use these rules when turning a validated region plan into native PowerPoint objects. They are defaults, not deck-specific style tokens. Read [visual-consistency.md](visual-consistency.md) when the plan uses role normalization, component families, generated icons, calibration, feedback issues, or authorized redesigns.

## Derive a source style profile

Cluster repeated source evidence before authoring:

- theme colors and dominant fill/stroke colors;
- title, subtitle, body, caption, number, and callout font families;
- effective font sizes and weights;
- left/right/top/bottom insets;
- alignment, line spacing, paragraph spacing, and tab stops;
- corner radii, line widths, shadow blur/distance/opacity;
- recurring gutters, card widths, rows, columns, and baseline rhythm.

Prefer master/layout/theme values over duplicated slide-local evidence. Pass the resulting tokens into helpers; never bake a corporate palette, font, logo, page number, or slide coordinate into the skill.

After clustering, choose one canonical valid style per semantic role and apply it to native and reconstructed objects. Do not preserve an accidental source outlier merely because it is native. Record explicit exceptions instead of introducing per-slide drift.

Artifact Tool numeric `fontSize` values are pixels while many source PPTX values are points. Convert explicitly:

```text
pixels = points × 96 / 72
```

## Rebuild text semantically

- Use one textbox per logical heading, paragraph, caption, label, or list item.
- Preserve a source line break only when it is semantically intentional. Let the target textbox wrap naturally otherwise.
- Match insets and paragraph geometry before forcing manual line breaks.
- Preserve CJK punctuation and prohibit one-character orphan lines when a small width, size, or tracking adjustment can resolve them.
- Preserve the source language-specific font fallback chain. Flag font substitution instead of hiding it with excessive size changes.
- Re-render after changing font, width, height, insets, tracking, or spacing.
- Give normalized objects semantic names following the `mppe|role=...|family=...|instance=...|part=...` contract so final QA can compare peers after round-trip.

## Rebuild bullets and lists

Preferred method: structured native paragraphs with `bulletCharacter`, a positive `marginLeft`, and a negative hanging `indent`. Wrapped lines must begin at the item text position.

Use separate objects only when the renderer cannot reproduce the source or the source requires a heavy dot:

- one native filled-circle shape per item;
- one textbox per logical item;
- no literal `•` typed into body text;
- keep the circle centered on the first text line, not the full multiline block.

Initial metrics relative to body font size:

| Metric | Starting range |
|---|---:|
| Dot diameter | `0.38–0.45 × font size` |
| Dot-to-text gap | `0.55–0.75 × font size` |
| Within-item line spacing | `1.05–1.15 × font size` |
| Between-item spacing | `0.4–0.6 × font size` |

The rendered slide is authoritative. Confirm visual weight, indent, gap, multiline alignment, and vertical rhythm at full size.

For independent dot shapes, use the same list instance and item index in the dot and textbox names. Reject helper-reported overflow. The text-left coordinate must be identical for all items in one list; manual source wrapping must not create a different indent on continuation lines.

## Rebuild cards and simple structures

- Use native shapes for cards, panels, separators, arrow bars, badges, timelines, and basic database/server metaphors.
- Reproduce visual hierarchy rather than tracing every pixel artifact.
- Use groups only when they improve coherent movement and editing; do not create deeply nested groups.
- Preserve z-order deliberately.
- Reconstruct shadows only when they are visually meaningful; use restrained native shadows rather than rasterized effects.
- Assign every repeated structure a component family and instance ID. Use one family profile for card size, icon center, title frame, divider, body inset, radius, line, and shadow; vary only axes explicitly marked flexible.

## Rebuild diagrams connector-first

1. Place or compute anchors.
2. Add native connectors first with stable endpoints and routing.
3. Add nodes above the connectors.
4. Add labels last.
5. Render and verify that connectors are not hidden, detached, or piercing text.

Do not replace a native connector with a thin rectangle merely to mimic appearance.

## Rebuild tables and charts

Rebuild a table natively only when every required cell value, merge relationship, header, and semantic grouping is known. Preserve column alignment, number formats, and units.

Rebuild a chart natively only when categories, series, values, axis semantics, units, and ordering can be recovered. Keep the data source editable and record any inferred styling separately from verified data.

If reliable data cannot be recovered, retain the source visual as `raster-replaceable` with an exception entry. Never invent data to gain editability.

## Extract and prepare raster assets

- Crop from the source at the highest available resolution before generating anything new.
- Preserve true alpha when available.
- Remove only edge-connected background regions when cleaning a light background; preserve interior whites such as paper, highlights, and icon details.
- Trim transparent margins, then place with containment rather than cover-cropping.
- Normalize optical size so icons with different intrinsic whitespace appear consistent without distorting aspect ratio.
- Reject opaque rectangular tiles, clipped strokes, visible atlas seams, color fringes, excessive blur, or stretched aspect ratios.
- Place a clean icon on a native support circle when the design requires one. The icon asset itself must not contain a second support circle.

## Regenerate generic raster icons by default

Use `$imagegen` once per distinct semantic generic raster icon. Request a single isolated bitmap icon with genuine transparent alpha and no text, watermark, frame, tile, shadow, circle, or container. Create one shared icon-family specification and use the first accepted output as the visual anchor for subsequent prompts. Preserve a verified native editable icon rather than reducing it to raster.

Never synthesize:

- company or government logos;
- a named person or official portrait;
- product UI or screenshots;
- evidence images;
- official diagrams or charts;
- licensed or identifying artwork.

Record the exact prompt, selected output, file hash, reason for generation, and the fact that pixels remain non-editable.

Allow one initial generation and one defect-specific retry. If the second output is still opaque, cropped, stylistically incompatible, or contains a backing/container, stop and record a blocker. Do not silently reuse the old generic raster icon.

## Protect inherited presentation structure

- Keep recurring brand art, backgrounds, page numbers, and headers on the existing master/layout.
- Do not duplicate inherited elements as slide-local overlays.
- Delete only exact source IDs listed in the validated plan.
- Preserve speaker notes and add `[Sources]` notes for source slides and generated assets when the presentation workflow supports notes editing.
- Keep object names or alt text meaningful for retained raster assets and reconstructed semantic groups.
