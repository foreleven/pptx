---
'@office-kit/pptx': minor
---

Allow `setShapeTextAutoFit` to write paired normal-autofit font-scale and line-spacing-reduction ratios. Safe reads now reject native values outside the public ratio contract, while `getShapeTextAutoFitParamsRaw` exposes them for lossiness diagnostics.
