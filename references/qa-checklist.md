# QA Checklist

Run automated checks and inspect every slide individually at full size. A montage is useful for rhythm and consistency only.

## Source integrity and identity

- [ ] Source exists, is `.pptx`, is not encrypted, and was copied before work.
- [ ] Source SHA-256 at handoff equals the initial source SHA-256.
- [ ] Output is a new `_editable.pptx` sibling, never the source path.
- [ ] QA report records the final output SHA-256 and that hash matches the delivered file.
- [ ] Render, inspect, and diff folders were freshly created for this final hash; no stale output was reused.

## Presentation structure

- [ ] Slide count, order, canvas dimensions, orientation, masters, layouts, and section intent are preserved.
- [ ] Headers, footers, page numbers, theme, notes, and inherited artwork are preserved unless explicitly changed.
- [ ] No empty placeholders, empty structural textboxes, broken relationships, or orphaned package parts remain.
- [ ] Unsupported macros, ActiveX, OLE, SmartArt, media, links, transitions, equations, Ink, and vectors are preserved or recorded as explicit exceptions.

## Plan accounting

- [ ] Every planned source object ID exists in the source manifest.
- [ ] Every reconstruction region is mapped exactly once.
- [ ] Nested semantic `subregions` are checked against their own bbox, target type, action, and expected text, while only their full-footprint parent owns the source-object deletion.
- [ ] Every deleted flattened/scan composite has a completed `visualAssets` inventory. Each visible logo, illustration, photo, UI capture, official diagram, and generic ico has a bbox, class, disposition, and non-empty reason.
- [ ] Every retained, extracted, or regenerated composite visual asset has a matching provenance record and is present as final embedded media at its planned bbox. A generic ico has not been replaced by a Unicode glyph, letter, or unrelated native symbol.
- [ ] No source object is deleted by a wildcard or broad slide-level image rule.
- [ ] No unclassified large slide-local raster remains.
- [ ] A retained, extracted, or review-only large raster substantially matches its declared region; a small image merely inside or touching the region cannot authorize it.
- [ ] Every retained or regenerated raster has a reason, source/provenance, hashes, alt text, and editability label.
- [ ] Asset source hashes are recomputed from every recorded input file or source-package `mediaRef`; output hashes are recomputed from output files and match the bytes embedded in the final PPTX.
- [ ] Every `manual-review` item is visible in the human-readable ledger.

## Text and semantic accuracy

- [ ] Normalized source ledger text is accounted for in the final inspect output.
- [ ] Expected native text is found inside its own planned region, not satisfied by duplicate wording elsewhere on the slide; raster-only actions are proven by exact media hashes instead.
- [ ] Numbers, dates, units, acronyms, names, punctuation, and CJK characters match the source.
- [ ] Uncertain text is flagged rather than guessed.
- [ ] Logical paragraphs are not fragmented into one textbox per visual source line.
- [ ] Titles do not wrap unexpectedly.
- [ ] No text is clipped, truncated, overflowing, or covered.
- [ ] No isolated single-character CJK lines remain when avoidable.
- [ ] Font substitutions and fallback changes are identified.
- [ ] Native and reconstructed text use the frozen canonical style for their semantic role, or carry an explicit role exception.
- [ ] Slide titles remain single-line; multi-line card titles stay within their declared role and share a consistent title frame and divider anchor.

## Bullets and paragraph layout

- [ ] Bullets use structured paragraphs or independent native dot shapes, not a literal bullet character embedded in body text.
- [ ] The dot is visually heavy enough and centered on the first line.
- [ ] Dot-to-text gap is consistent and sufficient.
- [ ] Wrapped lines align with item text, not the dot.
- [ ] Within-item and between-item spacing match the slide rhythm.
- [ ] Multiple items use consistent hanging indents and vertical offsets.
- [ ] Independent dot diameter and dot-to-text gap fall within the declared font-relative ranges.
- [ ] Dot centers align to the first line within `0.15 × font size`, all item text-left coordinates match, and helper-reported overflow is zero.

## Geometry and appearance

- [ ] No unexpected overlap, detached connector, or object outside the slide canvas.
- [ ] Cards, grids, baselines, margins, and gutters are visually consistent.
- [ ] Connector routing and z-order remain correct.
- [ ] Lines, radii, shadows, fills, and alignment match source evidence.
- [ ] Objects remain legible at presentation size; no microtext was introduced to force a fit.
- [ ] Every repeated component is assigned to a component family and matches its declared relative geometry within `max(2 px, 0.25% of the corresponding slide axis)` unless that axis is explicitly flexible.

## Raster and icon quality

- [ ] Images preserve aspect ratio and adequate effective resolution.
- [ ] Photos, logos, screenshots, and official artwork were not synthesized.
- [ ] Icons have real transparency where required.
- [ ] No opaque white tile, chroma fringe, atlas seam, clipped edge, double support circle, or unintended shadow is visible.
- [ ] Optical icon sizes are balanced across peers.
- [ ] Every raster image has meaningful alt text or object description.
- [ ] Every generic raster icon is regenerated by default, carries `generic-icon` classification, semantic concept, family ID, shared style prompt, source/output hashes, and ImageGen provenance.
- [ ] Native editable icons remain native; logos, photos, product UI, evidence screenshots, official diagrams, and identity-bearing artwork were not generated.
- [ ] Alpha-visible icon centroids remain within `4%` of the intended frame center and optical coverage stays within the icon-family target `±8%`.
- [ ] Icon image canvases do not occlude numbered badges, card edges, text, or neighboring components; required support circles occur exactly once as named native shapes.

## Calibration and annotated feedback

- [ ] Calibration is `approved` for user-gated work or `complete` for automatic work, with representative slide renders and a frozen profile hash.
- [ ] Every annotated region has a feedback issue bound to exact slide/object IDs and a local, component-family, or deck-wide propagation scope.
- [ ] No feedback issue remains `open`; every `waived` issue has a reason and evidence.
- [ ] Family/deck issues list every peer instance inspected, not only the originally marked object.
- [ ] User-approved redesigns are limited to exact authorized regions, preserve the planned text, and do not suppress QA outside those regions.

## Rendering and round-trip

- [ ] Import the final PPTX, export a new round-trip PPTX, and re-import it successfully.
- [ ] Render every original final slide and every round-trip slide at full size.
- [ ] Compare slide count, canvas, planned text, planned object presence, and visible geometry after round-trip.
- [ ] Generate per-slide visual diffs and review high-difference regions.
- [ ] Visual-difference masks exclude only exact `user-approved-redesign` regions; typography, bounds, text, brand, and round-trip checks still run inside those regions.
- [ ] Run the active `$presentations` template-fidelity and slide tests.
- [ ] Corroborate inherited-template findings from generic slide tests with layout JSON; inherited objects may otherwise look like false positive overlaps or bounds issues.
- [ ] For advanced effects or sensitive fonts, spot-check in desktop PowerPoint because the bundled renderer can differ from Office.

## Final editability report

- [ ] State counts or coverage for `text-editable`, `structure-editable`, `data-editable`, and `raster-replaceable` content.
- [ ] List retained raster, regenerated icons, and manual-review exceptions by slide.
- [ ] Confirm the source hash is unchanged and the final hash is the file actually delivered.
