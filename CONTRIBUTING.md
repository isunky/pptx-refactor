# Contributing / 参与贡献

Thanks for helping improve PPTX Refactor. 欢迎改进 PPTX Refactor。

## Before opening an issue / 提交 Issue 前

- Search existing issues and include the Codex version, operating system, and the selected reconstruction mode.
- Describe the smallest reproducible behavior and the expected result.
- Never upload confidential, customer-owned, identity-bearing, or licensed PowerPoint files to a public issue. Replace them with a synthetic reproduction.
- 请勿在公开 Issue 中上传机密、客户所有、包含身份信息或受许可限制的 PPT 文件；请改用合成样例复现。

## Pull requests

1. Create a focused branch from `main`.
2. Keep Skill paths relative and preserve `$pptx-refactor` as the public invocation ID.
3. Run `node scripts/validate_skill_bundle.mjs` and `node --test tests/*.test.mjs`.
4. Explain observable behavior changes and add or update tests for them.
5. Use synthetic fixtures only; do not commit real customer decks or extracted assets.

## Release policy

Releases use semantic tags such as `v0.2.0`. A published tag and its assets are immutable. Create a new patch version instead of replacing an existing Release.
