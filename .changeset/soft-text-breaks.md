---
'@office-kit/pptx': minor
---

Replace `setShapeParagraphRuns` with the canonical `setShapeParagraphElements` API for authoring ordered text runs and native DrawingML line breaks without converting them into separate paragraphs. Line-break formatting remains read-only and is not accepted by the authoring input.
