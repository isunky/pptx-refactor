# Conversion Decision Matrix

Use this matrix after preflight and before writing `conversion-plan.json`. Ratios are triage signals; every slide still requires full-size visual review.

## Choose a mode

| Mode | Native reconstruction | Raster preservation | Use when |
|---|---:|---:|---|
| `balanced` (default) | Text, lists, cards, simple diagrams, recoverable tables/charts | Logos, photos, screenshots, official/complex graphics | The user wants strong editability without sacrificing authenticity or layout stability |
| `maximum-editability` | Rebuild every reliably understood semantic object | Only authenticity, evidence, or unsupported-feature exceptions | The user explicitly prioritizes editability and accepts more reconstruction effort and small visual differences |
| `fidelity-first` | Rebuild only high-confidence text and simple structure | Preserve complex or fragile regions | Exact appearance is more important than deep editability |

Do not infer a redesign request from a conversion request. Preserve wording and visual language unless the user asks to rewrite or redesign.

Across all modes, normalize native and rebuilt content by semantic role and repeated component family. This is consistency repair, not content rewriting. Use `intent: user-approved-redesign` only for an exact region the user explicitly authorized.

## Classify each slide

| Class | Typical evidence | Default treatment |
|---|---|---|
| `native-editable` | Low image coverage, meaningful native text/shapes | Preserve; repair only clear defects |
| `mixed` | Native header/footer plus one or more body screenshots/composites | Rebuild mapped body regions; preserve inherited chrome |
| `flattened` | One dominant slide-local image, little or no native text | Reconstruct the full content region from a region plan |
| `low-quality-scan` | Full-slide raster, blur, skew, compression artifacts, uncertain text | Transcribe high-confidence content; preserve uncertain evidence; require manual review where needed |

Coverage heuristics:

- `< 25%` slide-local raster area: usually preserve native structure and process local images only.
- `25–90%`: usually a mixed slide; map and rebuild distinct regions.
- `>= 90%`: candidate for full content reconstruction.

These thresholds are not decisions. A transparent image may cover a large box without visible content, and inherited artwork may be misreported as slide-local. Confirm against render, inspect output, and PPTX package evidence.

## Decide per object or region

| Source content | Preferred action | Target type | Escalate when |
|---|---|---|---|
| Correct native text | `keep-native` | text | Font substitution or destructive grouping prevents safe reuse |
| Raster text | `rebuild-text` | text | Digits, units, acronyms, or proper nouns are uncertain |
| Cards, dividers, callouts | `rebuild-shape` | shape/group | Appearance depends on unsupported effects |
| Arrow/process/relationship diagram | `rebuild-shape` | connectors + shapes + text | Semantics or routing cannot be determined |
| Table with reliable cells | `rebuild-table` | native table | Merged structure or values are uncertain |
| Chart with recoverable series/categories | `rebuild-chart` | native chart | Data, scale, or series meaning is uncertain |
| Photo or authentic visual | `extract-raster` or `retain-raster` | image | Source quality is insufficient or licensing/provenance is unknown |
| Logo, official UI, evidence screenshot | `retain-raster` | image | Never synthesize a substitute |
| Generic raster icon, any source quality | `regenerate-icon` | transparent image | Stop if it is actually a brand, identity, UI, evidence, or official asset |
| Verified native editable icon | `keep-native` | native shape/group | Normalize its placement without converting it to raster |
| SmartArt/OLE/media/unsupported vector | `manual-review` or `retain-raster` | preserved object/image | Any conversion could remove behavior or data |

## Confidence gates

- `0.95–1.00`: safe for exact transcription or native reconstruction after visual confirmation.
- `0.80–0.94`: reconstruct only when a second signal confirms the content or structure.
- `< 0.80`: do not guess. Preserve as raster or mark `manual-review`.

Use stricter thresholds for financial values, dates, versions, IP addresses, identifiers, units, and chart data.

Any action that deletes/replaces source objects (`rebuild-*`, `extract-raster`, or `regenerate-icon`) is forbidden below `0.80`; choose `manual-review`. Apply the same rule when any text-ledger entry has `needsReview:true` or confidence below `0.80`.

## Destructive-region safety

Normalize plan and source-object boxes to the slide pixel canvas before comparison. A destructive region must cover at least 95% of every bound source object and must approximate the combined deletion footprint (IoU at least 0.80, or a containing replacement zone no more than 1.25× its area). A 1×1 intersection, narrow sliver, or unrelated oversized zone is not permission to delete the source object.

For a single composite screenshot with multiple semantic areas, create one full-footprint parent disposition and nested `subregions`. The parent owns the one source deletion; children describe bullet, table, chart, icon, or other reconstruction zones and repeat the exact source ID without consuming it again.

## Editability labels

Record one or more labels for each output region:

- `text-editable`
- `structure-editable`
- `data-editable`
- `raster-replaceable`

Regenerated icons remain `raster-replaceable`; their position, crop, size, and file are editable, not the pixels themselves.

Asset classes are `generic-icon`, `logo`, `photo`, `product-ui`, `evidence-screenshot`, `official-diagram`, `complex-illustration`, and `other`. A `generic-icon` is eligible for source extraction first and, only when that is not visually clean, regeneration. Other classes are not eligible for regeneration.
