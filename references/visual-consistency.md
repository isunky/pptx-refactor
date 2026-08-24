# Visual Consistency and Feedback Closure

Use this reference after source analysis and before authoring. It converts subjective visual goals into role, component-family, icon-family, calibration, and issue-closure contracts without embedding a deck-specific brand.

## Normalize the full deck by semantic role

Assign every meaningful text object one role. Recommended generic roles are `slide-title`, `section-title`, `card-title`, `callout-title`, `body`, `bullet-body`, `caption`, `number`, and `footer`. Add a role only when the source has a genuinely different hierarchy.

Choose one canonical style per role from master/layout values and the dominant valid source cluster. Apply it to native and reconstructed objects. Exact family, weight, and normalized color must match. Start with a font-size tolerance of `0.5 px` and a line-spacing tolerance of `0.03`; narrower deck-specific tolerances may be declared. Do not treat all headings as one role.

- `slide-title` is single-line unless an explicit exception is recorded.
- A `card-title` may allow one or two lines. Every instance in the same component family uses the same title frame height, vertical alignment, and divider anchor.
- Body roles share typeface, effective size, color, line spacing, paragraph spacing, and insets. Preserve semantic emphasis as text runs, not as unrelated role drift.
- A role exception records the object ID, changed token, reason, and expected visual effect.

## Define repeated component families

Use component families for repeated cards, process steps, statistic tiles, bottom callouts, icon cells, section labels, and similar structures. A family profile declares the component frame and the relative boxes or anchors of its parts.

Name normalized objects:

```text
mppe|role=<role>|family=<family>|instance=<instance>|part=<part>
```

Identifiers use lowercase ASCII letters, digits, `_`, or `-`. An instance may omit a role or family only when that dimension does not apply.

Compare family peers in slide pixels. Unless the plan declares a tighter tolerance, a part may vary by at most `max(2 px, 0.25% of the corresponding slide axis)`. Content-dependent body height may vary only when the family marks that axis flexible. Title baseline, divider position, icon center, support diameter, card edges, and internal gutters are not flexible by default.

## Calibrate before propagating

Use `user-gated` calibration only when the user asks to see a sample first. Produce a sample-only deck or render, record the approved role/family/icon tokens, then freeze them for the full build.

Otherwise use `automatic` calibration and select up to three distinct slides:

1. the slide with the densest or most complex list;
2. a slide containing the most repeated component family;
3. an icon-rich slide.

If one slide covers multiple categories, do not select it twice. Resolve hard findings in the calibration renders before full-deck authoring. Record slide numbers, render paths, status, and the frozen profile hash.

## Regenerate a coherent generic icon family

The default `extract-or-regenerate-generic` policy first preserves a cleanly extracted generic raster icon, then uses regeneration only when extraction would leave a seam, tile, crop loss, or inadequate resolution. Preserve verified native editable icons. Never synthesize a logo, photograph, named person, product UI, evidence screenshot, official diagram/chart, licensed illustration, or identity-bearing artwork.

Create a family specification before generating:

- palette roles rather than deck-specific hardcoded colors;
- stroke-weight class and corner treatment;
- flat, outline, filled, or mixed construction;
- detail level and viewing angle;
- transparent background and target optical coverage, default `0.72`;
- no text, watermark, tile, frame, shadow, circular backing, or decorative container.

Generate one isolated image per distinct semantic icon. The first accepted output is the visual anchor for subsequent assets. Keep a shared prompt prefix and add only the semantic concept per asset. Permit one initial call and one defect-specific retry. If the second output still violates the family contract, mark the asset blocked.

Place the generated bitmap inside a named native icon cell. The alpha-visible centroid should stay within `4%` of the intended frame center and optical coverage within the family target `±8%`. A required support circle is one native shape; more than one named support or a visible backing inside the bitmap is a defect.

## Process annotated visual feedback

Create one `feedbackIssues[]` entry per marked region:

- `issueId`, `slideNumber`, annotation source, and annotation bbox;
- exact related object IDs;
- category such as `title-wrap`, `style-drift`, `bullet-gap`, `bullet-indent`, `line-spacing`, `opaque-tile`, `crop-fragment`, `double-support`, `optical-size`, `component-alignment`, `z-order`, or `occlusion`;
- `scope`: `local`, `component-family`, or `deck`;
- expected fix, status, evidence, and optional waiver reason.

For a family or deck issue, repairing the marked object is insufficient. Enumerate the peer instances inspected and record the result. Final states are `fixed` or `waived`; `open` blocks delivery. A waiver is never a substitute for an unattempted repair.

## Record intentional redesigns

Keep the existing conversion actions. Mark a planned region with `intent: user-approved-redesign`, an authorization note, exact bbox/object IDs, and `preserveText: true`. Source-to-final visual-difference checks may ignore only that bbox. Text coverage, brand preservation, bounds, component consistency, and round-trip checks still apply.
