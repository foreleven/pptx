// internal/drawingml — a: namespace: shapes, text, geometry, color, effects.
// Allowed imports: internal/xml.

export { paragraphText, paragraphsOf, runsOf, textBodyText } from './text-body.ts';
export {
  applyAlignmentToAllParagraphs,
  applyBulletToAllParagraphs,
  applyBulletToParagraph,
  type BulletStyle,
  type ParagraphAlignment,
  replaceTextInTree,
  replaceTokensInTree,
  setTextBody,
} from './text-body-mutation.ts';
export type {
  TextFormat,
  TextGradient,
  TextLineProperties,
  TextOutline,
  TextShadow,
  TextUnderlineFill,
  TextUnderlineLine,
} from './text-format.ts';
export { applyFormatToAllRuns, applyRunFormat, applyRunState } from './text-format.ts';
export type { TextRunState } from './text-format.ts';
export type { ParsedColor } from './color.ts';
export { buildColorElement, parseColor, parseSrgbHex } from './color.ts';
export type {
  GradientFillOptions,
  GradientStop,
  PatternFillOptions,
  PatternPreset,
} from './fill.ts';
export {
  buildGradientFill,
  buildPatternFill,
  clearFill,
  setGradientFill,
  setNoFill,
  setPatternFill,
  setSolidFill,
} from './fill.ts';
export type {
  Effect,
  EffectBlend,
  GlowOptions,
  PresetShadow,
  ShadowOptions,
  ShapeEffectOptions,
} from './effects.ts';
export {
  buildEffectList,
  clearEffects,
  parseEffectList,
  setEffects,
  setGlow,
  setShadow,
} from './effects.ts';
export type {
  ArrowOptions,
  LineAlignment,
  LineCap,
  LineCompound,
  LineDash,
  LineDashStop,
  LineEndSize,
  LineEndType,
  LineJoin,
  LineStyle,
  StrokeOptions,
} from './stroke.ts';
export {
  applyLineStyle,
  clearStroke,
  isLineAlignment,
  isLineCap,
  isLineCompound,
  isLineDash,
  isLineJoin,
  setNoStroke,
  setSolidStroke,
  setStrokeArrow,
  setStrokeCap,
  setStrokeCompound,
  setStrokeCustomDash,
  setStrokeDash,
  setStrokeJoin,
} from './stroke.ts';
export type {
  CustomGeometry,
  GeomCommand,
  GeomPath,
  GeomPoint,
  PathFillMode,
} from './custom-geometry.ts';
export { parseCustomGeometry } from './custom-geometry.ts';
export { applyHyperlinkToAllRuns } from './hyperlink.ts';
export { getPictureEmbedRId } from './picture-mutation.ts';
export type { ImageCrop } from './image-crop.ts';
export { buildImageSourceRectangle, coverImageCrop, croppedImageSize } from './image-crop.ts';
export type { Position, ShapeKindForGeometry, Size } from './geometry.ts';
export { readFlip, readPosition, readRotation, readSize } from './geometry.ts';
export {
  setAdjustValues,
  setFlip,
  setGroupChildSpace,
  setPosition,
  setRotation,
  setSize,
} from './geometry-mutation.ts';
