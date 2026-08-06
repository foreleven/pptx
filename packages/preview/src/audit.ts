// Text-layout audit: measures every shape's laid-out text against its box and
// reports where ink escapes (overflow) or paragraphs soft-wrap (段落ち).
//
// The measurement runs through the SAME pipeline the preview renders with —
// `resolveTextBodyModel` (effective bodyPr / autofit / paragraph model) into
// `layoutCore` (wrap + baseline placement) — so a reported overflow is exactly
// what the rasterized preview would paint outside the box.
//
// Browser-safe like text-layout.ts: the measurer is injected. The default
// heuristic measurer estimates widths per character (results are flagged
// `approximate`); pass `buildFontkitMeasurer()` from
// `@office-kit/pptx-preview/node` for glyph-accurate metrics.

import {
  getGroupChildren,
  getPresentationTheme,
  getShapeBoundsResolved,
  getShapeKind,
  getShapeName,
  getShapePlaceholderType,
  getShapeTextColumns,
  getShapeTextDirection,
  getSlides,
  getSlideShapes,
  type PresentationData,
  type PresentationTheme,
  type SlideShapeData,
} from '@office-kit/pptx';
import {
  buildSvgTextInput,
  EMU_PER_PX,
  resolveTextBodyModel,
  verticalLayoutOf,
} from './render-slide.ts';
import {
  defaultMeasurer,
  layoutCore,
  SANS,
  type ColumnLayout,
  type FontSpec,
  type TextMeasurer,
} from './text-layout.ts';

interface IssueBase {
  /** 0-based deck position of the slide the shape sits on. */
  readonly slideIndex: number;
  readonly shapeName: string | null;
  /** True when any run was measured by estimate rather than real glyph
   *  metrics (heuristic measurer, or a font without the needed glyphs).
   *  Treat borderline overflows as advisory when set. */
  readonly approximate: boolean;
}

export interface TextLayoutMeasurement extends IssueBase {
  readonly box: {
    readonly widthPx: number;
    readonly heightPx: number;
    readonly innerWidthPx: number;
    readonly innerHeightPx: number;
  };
  readonly layout: {
    /** The authored shape dimension that controls the text flow extent. */
    readonly axis: 'height' | 'width';
    readonly requiredInnerExtentPx: number;
    readonly requiredShapeExtentPx: number;
    readonly lineCount: number;
    readonly paragraphLineCounts: ReadonlyArray<number>;
    readonly softWrapCount: number;
    readonly softWraps: ReadonlyArray<{
      readonly paragraphIndex: number;
      readonly extraLines: number;
    }>;
    /** Ink coordinates relative to the inner layout frame's leading edge. */
    readonly inkTopPx: number;
    readonly inkBottomPx: number;
    readonly inkHeightPx: number;
    readonly overflowXPx: number;
    readonly overflowYPx: number;
  };
}

export type TextAuditIssue =
  | (IssueBase & {
      readonly kind: 'overflow-x' | 'overflow-y';
      /** How far the text ink escapes the box, in px at 96 DPI. */
      readonly overflowPx: number;
    })
  | (IssueBase & {
      readonly kind: 'soft-wrap';
      readonly paragraphIndex: number;
      /** Lines beyond the paragraph's authored line count (1 + explicit
       *  breaks) — i.e. how many times the text wrapped on its own. */
      readonly extraLines: number;
    });

export interface AuditTextLayoutOptions {
  /** Measurer for run widths / vertical metrics. Defaults to the heuristic
   *  measurer (browser-safe, approximate); pass the fontkit measurer from
   *  `@office-kit/pptx-preview/node` for real glyph metrics. */
  readonly measureText?: TextMeasurer;
  /** Overflow at or below this many px (96 DPI) is ignored. Default 1 —
   *  measurement and PowerPoint disagree by sub-pixel amounts routinely. */
  readonly tolerancePx?: number;
  /** Also report paragraphs that wrap onto more lines than their explicit
   *  breaks author (段落ち). Off by default: wrapping is normal for body
   *  text, so this is opt-in for single-line-intent content like titles. */
  readonly reportSoftWraps?: boolean;
}

export interface MeasureTextLayoutOptions {
  /** Uses the same injected measurer contract as SVG rendering and the audit API. */
  readonly measureText?: TextMeasurer;
}

const DEFAULT_TOLERANCE_PX = 1;

const round2 = (n: number): number => Math.round(n * 100) / 100;

// Depth-first over group trees. Children are audited in their own (child)
// coordinate space, where box and text agree; a group's non-uniform scale
// stretches both equally in the rendered output, so overflow verdicts hold.
function* walkShapes(shapes: ReadonlyArray<SlideShapeData>): Generator<SlideShapeData> {
  for (const shape of shapes) {
    if (getShapeKind(shape) === 'group') {
      yield* walkShapes(getGroupChildren(shape));
    } else {
      yield shape;
    }
  }
}

// The audit hands the measurer the AUTHORED font name (falling back to the
// generic sans substitute), unlike the render path's `substituteFamily` —
// a fontkit measurer with user-registered fonts resolves it first and only
// then falls back to the same substitution map.
const passthroughFamily = (family: string | null): string => family ?? SANS;

/** Measure one native text body with the exact model and layout core used by SVG rendering. */
const measureShape = (
  pres: PresentationData,
  theme: PresentationTheme | null,
  shape: SlideShapeData,
  slideIndex: number,
  measure: TextMeasurer,
): TextLayoutMeasurement | null => {
  const kind = getShapeKind(shape);
  if (kind !== 'shape' && kind !== 'graphicFrame') return null;
  const bounds = getShapeBoundsResolved(pres, shape);
  if (!bounds) return null;
  const model = resolveTextBodyModel(
    pres,
    shape,
    { x: bounds.x as number, y: bounds.y as number, w: bounds.w as number, h: bounds.h as number },
    theme,
    getShapePlaceholderType(shape),
    measure,
    '#000000', // colors don't affect metrics
  );
  if (model === null) return null;

  const vert = verticalLayoutOf(model.effectiveBody.vert ?? getShapeTextDirection(shape));
  const cols = getShapeTextColumns(shape);
  const columns: ColumnLayout | null =
    vert === 'none' && cols && cols.count >= 2
      ? {
          count: cols.count,
          gapPx: cols.gapEmu !== undefined ? cols.gapEmu / EMU_PER_PX : 12,
        }
      : null;
  const rect = model.svgTextRect(vert);
  if (rect.w <= 0 || rect.h <= 0) return null;

  // SVG semantics (see the render path's `svgScale`): only an authored
  // autofit shrinks; the heuristic factor is foreignObject-only.
  const autoFitScale = model.authoredAutofit ? model.autoFitScale : 1;
  const input = buildSvgTextInput({
    pres,
    shape,
    theme,
    paraData: model.paraData,
    numberLabels: model.numberLabels,
    autoFitScale,
    lineHeightScale: model.lineHeightScale,
    defaultPt: model.defaultPt,
    themeFace: model.themeFace,
    defaultColor: '#000000',
    anchor: model.anchor,
    wrap: model.effectiveBody.wrap !== 'none',
    innerX: rect.x,
    innerY: rect.y,
    innerW: rect.w,
    innerH: rect.h,
    measure,
    vert,
    columns,
    resolveFamily: passthroughFamily,
  });
  const core = layoutCore(input, measure);

  // Mirror layoutCore's frame: the ±90° rotations lay out into a frame with
  // swapped extents re-centred on the box, so ink must be compared against
  // that frame, not the box itself.
  const rotated = vert === 'cw90' || vert === 'cw270';
  const frame = rotated
    ? {
        x: input.boxXpx + input.boxWpx / 2 - input.boxHpx / 2,
        y: input.boxYpx + input.boxHpx / 2 - input.boxWpx / 2,
        w: input.boxHpx,
        h: input.boxWpx,
      }
    : { x: input.boxXpx, y: input.boxYpx, w: input.boxWpx, h: input.boxHpx };

  // Ink extents over every placed line. Empty lines (blank paragraphs, lines
  // of spaces) paint nothing — they influence later lines' positions but are
  // not themselves visible overflow.
  let inkTop = Infinity;
  let inkBottom = -Infinity;
  let overflowX = 0;
  let anyInk = false;
  for (const p of core.placements) {
    const toks = [...p.line.tokens];
    while (toks.length > 0 && (toks[toks.length - 1]!.isSpace || toks[toks.length - 1]!.isBreak)) {
      toks.pop();
    }
    if (!toks.some((t) => !t.isSpace && !t.isBreak)) continue;
    anyInk = true;
    let lineW = 0;
    for (const t of toks) {
      if (!t.isBreak) lineW += t.width;
    }
    const left =
      p.line.textAnchor === 'middle'
        ? p.line.anchorX - lineW / 2
        : p.line.textAnchor === 'end'
          ? p.line.anchorX - lineW
          : p.line.anchorX;
    const over = Math.max(frame.x - (left + p.dx), left + p.dx + lineW - (frame.x + frame.w));
    if (over > overflowX) overflowX = over;
    inkTop = Math.min(inkTop, p.baselineY - p.line.ascent);
    inkBottom = Math.max(inkBottom, p.baselineY + p.line.descent);
  }
  if (!anyInk) return null;

  const approximate = detectApproximate(input.paragraphs, measure);
  const overflowY = Math.max(frame.y - inkTop, inkBottom - (frame.y + frame.h));
  const lineCounts = new Map<number, number>();
  for (const placement of core.placements) {
    lineCounts.set(placement.line.paraIndex, (lineCounts.get(placement.line.paraIndex) ?? 0) + 1);
  }
  const paragraphLineCounts = input.paragraphs.map(
    (_paragraph, index) => lineCounts.get(index) ?? 0,
  );
  // A soft wrap is measured even when the audit caller does not request a warning,
  // because the build artifact needs to expose title/label wrap pressure after rendering.
  const softWraps =
    vert === 'none'
      ? input.paragraphs.flatMap((paragraph, paragraphIndex) => {
          const explicitBreaks = paragraph.pieces.reduce(
            (count, piece) => count + (piece.isBreak ? 1 : 0),
            0,
          );
          const extraLines = (lineCounts.get(paragraphIndex) ?? 0) - 1 - explicitBreaks;
          return extraLines > 0 ? [{ paragraphIndex, extraLines }] : [];
        })
      : [];
  const shapeExtentPx = Number(rotated ? bounds.w : bounds.h) / EMU_PER_PX;
  const requiredShapeExtentPx = shapeExtentPx + core.requiredH - frame.h;
  return {
    slideIndex,
    shapeName: getShapeName(shape),
    approximate,
    box: {
      widthPx: round2(Number(bounds.w) / EMU_PER_PX),
      heightPx: round2(Number(bounds.h) / EMU_PER_PX),
      innerWidthPx: round2(input.boxWpx),
      innerHeightPx: round2(input.boxHpx),
    },
    layout: {
      axis: rotated ? 'width' : 'height',
      requiredInnerExtentPx: round2(core.requiredH),
      requiredShapeExtentPx: round2(Math.max(0, requiredShapeExtentPx)),
      lineCount: core.placements.length,
      paragraphLineCounts,
      softWrapCount: softWraps.reduce((count, wrap) => count + wrap.extraLines, 0),
      softWraps,
      inkTopPx: round2(inkTop - frame.y),
      inkBottomPx: round2(inkBottom - frame.y),
      inkHeightPx: round2(inkBottom - inkTop),
      overflowXPx: round2(Math.max(0, overflowX)),
      overflowYPx: round2(Math.max(0, overflowY)),
    },
  };
};

// A shape's verdict is approximate when any of its runs was measured by
// estimate: the heuristic measurer (no vertical metrics) or a fontkit
// measurer that hit a missing glyph (`approximate` flag).
const detectApproximate = (
  paragraphs: ReturnType<typeof buildSvgTextInput>['paragraphs'],
  measure: TextMeasurer,
): boolean => {
  const seen = new Set<string>();
  for (const para of paragraphs) {
    for (const piece of para.pieces) {
      if (piece.isBreak || piece.text === '') continue;
      const key = `${piece.family}|${piece.sizePx}|${piece.bold}|${piece.italic}|${piece.text}`;
      if (seen.has(key)) continue;
      seen.add(key);
      const spec: FontSpec = {
        family: piece.family,
        sizePx: piece.sizePx,
        bold: piece.bold,
        italic: piece.italic,
        letterSpacingPx: piece.letterSpacingPx,
      };
      const r = measure(piece.text, spec);
      if (r.approximate === true || r.ascentPx === undefined) return true;
    }
  }
  return false;
};

/**
 * Measure every shape text body with the SVG layout engine, including required
 * extent, ink bounds, line counts, and soft wraps. Table cells are not included.
 */
export const measureTextLayout = (
  pres: PresentationData,
  options: MeasureTextLayoutOptions = {},
): ReadonlyArray<TextLayoutMeasurement> => {
  const measure = options.measureText ?? defaultMeasurer;
  const theme = getPresentationTheme(pres);
  const measurements: TextLayoutMeasurement[] = [];
  const slides = getSlides(pres);
  for (let slideIndex = 0; slideIndex < slides.length; slideIndex++) {
    for (const shape of walkShapes(getSlideShapes(slides[slideIndex]!))) {
      const measurement = measureShape(pres, theme, shape, slideIndex, measure);
      if (measurement) measurements.push(measurement);
    }
  }
  return measurements;
};

/** Report measured overflow and, when requested, paragraph soft wraps. */
export const auditTextLayout = (
  pres: PresentationData,
  options: AuditTextLayoutOptions = {},
): ReadonlyArray<TextAuditIssue> => {
  const tolerancePx = options.tolerancePx ?? DEFAULT_TOLERANCE_PX;
  const reportSoftWraps = options.reportSoftWraps ?? false;
  const issues: TextAuditIssue[] = [];
  const measurements = measureTextLayout(
    pres,
    options.measureText ? { measureText: options.measureText } : {},
  );
  for (const measurement of measurements) {
    const issueBase = {
      slideIndex: measurement.slideIndex,
      shapeName: measurement.shapeName,
      approximate: measurement.approximate,
    };
    if (measurement.layout.overflowXPx > tolerancePx) {
      issues.push({ ...issueBase, kind: 'overflow-x', overflowPx: measurement.layout.overflowXPx });
    }
    if (measurement.layout.overflowYPx > tolerancePx) {
      issues.push({ ...issueBase, kind: 'overflow-y', overflowPx: measurement.layout.overflowYPx });
    }
    if (reportSoftWraps) {
      for (const wrap of measurement.layout.softWraps) {
        issues.push({ ...issueBase, kind: 'soft-wrap', ...wrap });
      }
    }
  }
  return issues;
};
