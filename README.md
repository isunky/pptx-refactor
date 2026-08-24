# PPTX Refactor skill

This directory is the standalone maintenance source for the Codex
`pptx-refactor`（PPT重构）skill.

## Layout

- `SKILL.md` — workflow and operating contract
- `scripts/` — analyzer, plan validation, reconstruction, asset preparation,
  QA, and Windows unzip compatibility helpers
- `references/` — schemas, compatibility notes, reconstruction rules, and QA
  guidance
- `agents/` — Codex skill metadata

The installed Codex skill path
`C:\Users\ZXHY-NB\.codex\skills\pptx-refactor` is a directory junction to
this folder, so changes made here are immediately used by Codex. Keep relative
paths inside `SKILL.md` and the scripts portable; do not hard-code this checkout
path.
