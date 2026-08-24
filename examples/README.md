# Synthetic demo / 合成演示案例

This example contains no customer data, third-party artwork, external claims, or identity-bearing assets. It exists only to demonstrate the difference between visual fidelity and structural editability.

本案例不包含客户数据、第三方素材、外部事实声明或身份相关资产，仅用于展示“视觉相似”和“结构可编辑”之间的区别。

![Three-slide synthetic PPTX Refactor demo](demo-preview.webp)

## Files

- [`pptx-refactor-demo-flattened.pptx`](pptx-refactor-demo-flattened.pptx) — three slides, each stored as one full-slide raster image.
- [`pptx-refactor-demo-editable.pptx`](pptx-refactor-demo-editable.pptx) — the same narrative rebuilt with native PowerPoint text and shapes.
- [`demo-structure.json`](demo-structure.json) — inspect-derived object counts for both decks.

## Verified structure

| Deck | Slides | Native text boxes | Native shapes | Full-slide images |
|---|---:|---:|---:|---:|
| Flattened input | 3 | 0 | 0 | 3 |
| Editable output | 3 | 20 | 9 | 0 |

The decks intentionally have the same visible content. The difference appears when a user selects, edits, restyles, or replaces individual PowerPoint objects.

两份文件有意保持相同的可见内容；真正的区别体现在选择、修改、统一样式或替换单个 PowerPoint 对象时。
