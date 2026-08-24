# Compatibility and Platform Rules

## Accepted input

Support `.pptx` only.

Block or request explicit conversion outside this workflow for:

- legacy `.ppt` files;
- encrypted or password-protected presentations;
- files that cannot be opened without repairing the package.

Never silently remove or downgrade unsupported content to make a file pass.

## Features requiring preservation or manual review

Detect package parts and relationships for:

- VBA projects or macro-enabled content;
- ActiveX controls;
- embedded or linked OLE objects;
- SmartArt and diagram parts;
- audio, video, linked media, and external data;
- animations, transitions, custom shows, and timing;
- equations and Ink;
- comments, modern comments, notes, and custom XML;
- embedded fonts;
- EMF, WMF, EPS, SVG, and other vector media not handled by the selected raster pipeline.

If the editing/export path cannot guarantee preservation, stop for user direction or record `manual-review` and avoid editing the affected region. Do not convert an authentic object to a guessed visual.

`ensure_raster_image.py` does not convert EMF/WMF/EPS. Preserve the source object, use a verified render crop as a documented raster exception, or regenerate only a non-branded generic icon when permitted.

## Windows execution

- Use exact runtime paths returned by the workspace dependency loader. Do not depend on globally installed Node, Python, LibreOffice, ImageMagick, or unzip.
- Some presentation helpers invoke `unzip -Z1` and `unzip -p`. `unzip.cmd` supports direct shell calls, but Node `spawnSync("unzip")` on Windows may not resolve a `.cmd` through `PATHEXT`.
- For an installed Node helper such as `inspect_template_deck.mjs` or `check_template_fidelity.mjs`, invoke `"$RUNTIME_NODE" scripts/unzip_compat.mjs --run <helper.mjs> <helper arguments...>`. The runner intercepts the helper's exact `unzip` calls and routes them through JSZip without modifying the installed presentation skill.
- Set `RUNTIME_NODE` to the bundled `node.exe` before invoking `unzip.cmd`, and set `RUNTIME_NODE_MODULES` before either compatibility route.
- Resolve bundled ESM packages from the returned runtime module path. If a tool requires a local `node_modules`, create a temporary verified junction or symlink and remove it after testing; do not vendor runtime dependencies into the skill.
- Use Node `path` and argument arrays. Do not build shell commands by concatenating quoted Windows paths.
- Enable UTF-8 for Python subprocesses when CJK paths or text are present, for example `PYTHONUTF8=1`.
- Set `PYTHON` to the exact bundled Python executable before running the upstream starter-deck helper when a contact sheet is requested; that helper otherwise defaults to `python3`, which may not exist on Windows.
- On some Windows hosts an Artifact Tool render process writes valid JSON and all slide PNGs, then exits with a native teardown code. This can make `slides_test.py` report a render failure even though rendering completed. Corroborate the JSON and every referenced file before diagnosing it as a teardown-only false negative. If a temporary launcher is used for testing, it may normalize the exit code only after those checks; never mask an arbitrary nonzero renderer exit.
- Prefer short ASCII intermediate filenames. Preserve the requested Unicode filename only for final handoff.
- Avoid case-only filename distinctions and case-colliding assets.

## File locks and safe copies

- PowerPoint may lock an open source or output file. Copy the source to an isolated temporary workspace and export to a new sibling path.
- Never export directly over an open or source PPTX.
- Use a new QA workspace keyed to the candidate final SHA-256 so stale renders cannot validate an older file.
- Validate absolute paths before recursive cleanup. Keep cleanup targets inside the named temporary workspace.

## Canvas and coordinate handling

- Derive width and height from the source layout or package; never assume `1280×720`, `16:9`, or a particular unit.
- Support 16:9, 4:3, portrait, and custom canvas sizes.
- Convert units explicitly and store the unit with each bounding box.
- Separate slide-local content from inherited master/layout layers before applying bounds or overlap findings.

## Renderer differences

The bundled presentation renderer may not match desktop PowerPoint for embedded fonts, advanced shadows, gradients, transparency, SmartArt, media, or animation. If those features matter:

1. preserve the original feature whenever possible;
2. run the bundled render-and-round-trip checks;
3. flag the affected slides;
4. perform a desktop PowerPoint spot-check before final acceptance.

Do not treat a clean montage as proof of compatibility. Inspect each affected slide and verify editability in the actual PPTX object model.
