// Run-property authoring for `<a:r><a:rPr>...</a:rPr></a:r>`.
//
// The ECMA-376 surface for `CT_TextCharacterProperties` (§17.18.83) is
// huge. We expose the subset that 90% of authoring scripts reach for:
// font face, size, color, bold, italic, underline.
//
// Sizes are pt and accept fractions; serialized as hundredths of a point
// per the schema (`sz="2400"` = 24pt). Colors accept any of:
//
//   - `#RRGGBB` / `RRGGBB` or CSS-order `#RRGGBBAA` / `RRGGBBAA` for sRGB
//   - `'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1'…'accent6'` for theme
//   - `null` to clear
//
// Mutation strategy: walk the `txBody`, ensure each `a:r` has an `a:rPr`,
// then set the relevant attributes / child elements. Existing properties
// not addressed by the format object are preserved.

import {
  NS,
  type XmlAttr,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  insertChildByRank,
  qname,
} from '../xml/index.ts';
import { fontSizeHundredthPt, textNonNegativePoint, textPointSpacing } from '../bounds.ts';
import { buildColorElement } from './color.ts';
import { buildEffectList, type Effect } from './effects.ts';
import {
  buildGradientFill,
  buildPatternFill,
  type GradientFillOptions,
  type PatternFillOptions,
} from './fill.ts';
import {
  applyLineStyle,
  type ArrowOptions,
  type LineAlignment,
  type LineCap,
  type LineCompound,
  type LineDash,
  type LineDashStop,
  type LineJoin,
} from './stroke.ts';

const NAME_R = qname('a', 'r', NS.dml);
const NAME_RPR = qname('a', 'rPr', NS.dml);
const NAME_LATIN = qname('a', 'latin', NS.dml);
const NAME_EA = qname('a', 'ea', NS.dml);
const NAME_CS = qname('a', 'cs', NS.dml);
const NAME_SOLID_FILL = qname('a', 'solidFill', NS.dml);
const NAME_EFFECT_LST = qname('a', 'effectLst', NS.dml);
const NAME_OUTER_SHDW = qname('a', 'outerShdw', NS.dml);
const ATTR_SZ = qname('', 'sz', '');
const ATTR_B = qname('', 'b', '');
const ATTR_I = qname('', 'i', '');
const ATTR_U = qname('', 'u', '');
const ATTR_STRIKE = qname('', 'strike', '');
const ATTR_SPC = qname('', 'spc', '');
const ATTR_KERN = qname('', 'kern', '');
const ATTR_BASELINE = qname('', 'baseline', '');
const ATTR_CAP = qname('', 'cap', '');
const ATTR_NORMALIZE_HEIGHT = qname('', 'normalizeH', '');
const ATTR_NO_PROOF = qname('', 'noProof', '');
const ATTR_DIRTY = qname('', 'dirty', '');
const ATTR_ERROR = qname('', 'err', '');
const ATTR_SMART_TAG_CLEAN = qname('', 'smtClean', '');
const ATTR_SMART_TAG_ID = qname('', 'smtId', '');
const ATTR_LANGUAGE = qname('', 'lang', '');
const ATTR_ALTERNATIVE_LANGUAGE = qname('', 'altLang', '');
const ATTR_TYPEFACE = qname('', 'typeface', '');
const ATTR_BLUR_RAD = qname('', 'blurRad', '');
const ATTR_DIST = qname('', 'dist', '');
const ATTR_DIR = qname('', 'dir', '');
const ATTR_ALGN = qname('', 'algn', '');
const ATTR_ROT_WITH_SHAPE = qname('', 'rotWithShape', '');
const NAME_HIGHLIGHT = qname('a', 'highlight', NS.dml);
const NAME_U_LN_TX = qname('a', 'uLnTx', NS.dml);
const NAME_U_FILL_TX = qname('a', 'uFillTx', NS.dml);
const NAME_U_FILL = qname('a', 'uFill', NS.dml);
const NAME_NO_FILL = qname('a', 'noFill', NS.dml);

const TEXT_UNDERLINE_TYPES = new Set([
  'none',
  'words',
  'sng',
  'dbl',
  'heavy',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dotDashHeavy',
  'dotDotDash',
  'dotDotDashHeavy',
  'wavy',
  'wavyHeavy',
  'wavyDbl',
]);

// CT_TextCharacterProperties (a:rPr) is an xsd:sequence: children must appear
// in this order or the run fails dml/pml schema validation. Setters strip the
// existing element then re-insert at the mandated slot via insertChildByRank.
const RPR_CHILD_RANK: Record<string, number> = {
  ln: 0,
  noFill: 1,
  solidFill: 1,
  gradFill: 1,
  blipFill: 1,
  pattFill: 1,
  grpFill: 1,
  effectLst: 2,
  effectDag: 2,
  highlight: 3,
  uLnTx: 4,
  uLn: 4,
  uFillTx: 5,
  uFill: 5,
  latin: 6,
  ea: 7,
  cs: 8,
  sym: 9,
  hlinkClick: 10,
  hlinkMouseOver: 11,
  rtl: 12,
  extLst: 13,
};
const rprChildRank = (el: XmlElement): number =>
  el.name.namespaceURI === NS.dml ? (RPR_CHILD_RANK[el.name.localName] ?? 99) : 99;

export interface TextFormat {
  /** Latin font family (`Calibri`, `Arial`, ...). Sets `<a:latin>`. */
  font?: string;
  /**
   * East Asian font family (`游明朝`, `メイリオ`, ...). Sets `<a:ea>`.
   * Renderers pick this typeface for CJK glyphs independently of `font`
   * (which only governs Latin glyphs) — set both when a run mixes Latin
   * and CJK text and needs a consistent look across the whole run.
   */
  fontEastAsian?: string;
  /** Complex-script font family (`Arial`, `Noto Naskh Arabic`, ...). Sets `<a:cs>`. */
  fontComplexScript?: string;
  /** Symbol font family (`Symbol`, `Wingdings`, ...). Sets `<a:sym>`. */
  fontSymbol?: string;
  /** Primary BCP 47 language tag. Mirrors `<a:rPr lang="..."/>`. */
  language?: string;
  /** Alternative BCP 47 language tag. Mirrors `<a:rPr altLang="..."/>`. */
  alternativeLanguage?: string;
  /** Font size in points; fractional values allowed (`12`, `12.5`). */
  size?: number;
  /**
   * Color. Accepts opaque or CSS-order alpha sRGB hex, an ECMA-376 scheme
   * color token (`tx1`, `accent1`, ...), or `null` to clear.
   */
  color?: string | null;
  bold?: boolean;
  italic?: boolean;
  /**
   * Underline style. `true` is shorthand for `'sng'` (single). Pass the
   * exact `ST_TextUnderlineType` token for other styles (`'dbl'`, `'wavy'`,
   * `'dash'`, ...).
   */
  underline?: boolean | string;
  /** Optional underline line choice; `followText` emits `<a:uLnTx/>`. */
  underlineLine?: TextUnderlineLine;
  /** Optional underline paint choice (`uFillTx` or explicit `uFill`). */
  underlineFill?: TextUnderlineFill;
  /**
   * Strikethrough style. `true` is shorthand for `'sngStrike'` (single
   * line). Pass the exact `ST_TextStrikeType` token (`'sngStrike'`,
   * `'dblStrike'`, `'noStrike'`) for other styles. `false` clears.
   */
  strike?: boolean | string;
  /** Direct text outline stored as `<a:rPr><a:ln>…</a:ln></a:rPr>`. */
  outline?: TextOutline;
  /**
   * Character spacing in 1/100 points (`0` = default). Negative values
   * tighten, positive values loosen. Mirrors `<a:rPr spc="…"/>`.
   */
  spc?: number;
  /**
   * Kerning threshold in 1/100 points (`ST_TextNonNegativePoint`, the same
   * unit as `spc`): `0` disables kerning, `1200` = apply kerning for runs
   * ≥12pt. Mirrors `<a:rPr kern="…"/>`.
   */
  kern?: number;
  /**
   * Baseline offset as a fraction of 1 (`0.3` = superscript ~30% up,
   * `-0.25` = subscript). PowerPoint emits ST_Percentage; this getter
   * returns the unit-fraction form for ergonomic comparisons.
   */
  baseline?: number;
  /**
   * Capitalization mode: `'none'`, `'small'` (smallCaps), or `'all'`
   * (allCaps). Mirrors `<a:rPr cap="…"/>`.
   */
  cap?: 'none' | 'small' | 'all';
  /**
   * Highlight color (cell-fill style background per run). Same color
   * format as `color`. Mirrors `<a:rPr><a:highlight>…</a:highlight></a:rPr>`.
   */
  highlight?: string | null;
  /** One editable CSS-compatible outer text shadow, or `null` to clear run effects. */
  textShadow?: TextShadow | null;
  /** Complete editable DrawingML `effectLst`, or `null` to clear run effects. */
  effects?: readonly Effect[] | null;
  /** One editable linear text gradient, or `null` to clear the direct run fill. */
  gradient?: TextGradient | null;
  /** One editable preset pattern text fill, or `null` to clear the direct run fill. */
  patternFill?: TextPatternFill | null;
  /** One editable embedded picture text fill, or `null` to clear the direct run fill. */
  pictureFill?: TextPictureFill | null;
}

/** Structural line semantics shared by text outlines and explicit underline lines. */
export interface TextLineProperties {
  readonly widthPt?: number;
  readonly cap?: LineCap;
  readonly dash?: LineDash;
  readonly customDash?: readonly LineDashStop[];
  readonly join?: LineJoin;
  readonly compound?: LineCompound;
  readonly alignment?: LineAlignment;
  readonly head?: ArrowOptions;
  readonly tail?: ArrowOptions;
}

/** Common editable text-outline subset plus an importer-only diagnostic carrier. */
export type TextOutline =
  | (TextLineProperties & {
      readonly kind: 'solid';
      readonly color: string;
      readonly unsupported?: string;
    })
  | (TextLineProperties & { readonly kind: 'none'; readonly unsupported?: string })
  | (TextLineProperties &
      GradientFillOptions & { readonly kind: 'gradient'; readonly unsupported?: string })
  | (TextLineProperties &
      PatternFillOptions & { readonly kind: 'pattern'; readonly unsupported?: string })
  | { readonly kind: 'unsupported'; readonly reason: string };

export type TextUnderlineLine =
  | { readonly kind: 'followText'; readonly unsupported?: string }
  | (TextLineProperties & {
      readonly kind: 'solid';
      readonly color: string;
      readonly unsupported?: string;
    })
  | (TextLineProperties & { readonly kind: 'none' | 'bare'; readonly unsupported?: string })
  | { readonly kind: 'unsupported'; readonly reason: string };

export type TextUnderlineFill =
  | { readonly kind: 'followText'; readonly unsupported?: string }
  | { readonly kind: 'none'; readonly unsupported?: string }
  | { readonly kind: 'solid'; readonly color: string; readonly unsupported?: string }
  | (GradientFillOptions & { readonly kind: 'gradient'; readonly unsupported?: string })
  | (PatternFillOptions & { readonly kind: 'pattern'; readonly unsupported?: string })
  | { readonly kind: 'group'; readonly unsupported?: string }
  | { readonly kind: 'picture'; readonly relationshipId: string; readonly unsupported?: string }
  | { readonly kind: 'unsupported'; readonly reason: string };

/** Common editable outer-shadow subset plus an importer-only diagnostic carrier. */
export type TextShadow =
  | {
      readonly color: string;
      readonly opacity?: number;
      readonly blurPt: number;
      readonly offsetXPt: number;
      readonly offsetYPt: number;
      readonly unsupported?: string;
    }
  | { readonly unsupported: string };

/** Common editable gradient subset plus an importer-only diagnostic carrier. */
export type TextGradient = GradientFillOptions | { readonly unsupported: string };

/** Editable direct preset-pattern text fill plus an importer diagnostic when normalized. */
export type TextPatternFill =
  | (PatternFillOptions & { readonly unsupported?: string })
  | { readonly unsupported: string };

/** Editable direct picture text fill resolved against the owning slide. */
export type TextPictureFill =
  | { readonly relationshipId: string; readonly unsupported?: string }
  | { readonly unsupported: string };

/** Direct, non-inherited CT_TextCharacterProperties state. */
export interface TextRunState {
  readonly normalizeHeight?: boolean;
  readonly noProof?: boolean;
  readonly dirty?: boolean;
  readonly error?: boolean;
  readonly smartTagClean?: boolean;
  readonly smartTagId?: number;
}

const setOrRemoveAttr = (
  attrs: XmlAttr[],
  name: ReturnType<typeof qname>,
  value: string | null,
): XmlAttr[] => {
  const filtered = attrs.filter((a) => a.name.localName !== name.localName);
  if (value !== null) filtered.push(attr(name, value));
  return filtered;
};

const removeRunFill = (rPr: XmlElement): void => {
  // A run may carry exactly one EG_FillProperties choice.
  rPr.children = rPr.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'].includes(
          c.name.localName,
        )
      ),
  );
};

const setSolidFill = (rPr: XmlElement, value: string | null): void => {
  removeRunFill(rPr);
  if (value === null) return;
  const fill = elem(NAME_SOLID_FILL, { children: [buildColorElement(value)] });
  insertChildByRank(rPr, fill, rprChildRank);
};

const setTextGradient = (rPr: XmlElement, value: TextGradient | null): void => {
  if (value === null) {
    removeRunFill(rPr);
    return;
  }
  if (!('stops' in value)) {
    throw new TypeError('TextFormat.gradient unsupported diagnostic state is read-only.');
  }
  if (value.path !== undefined && value.path !== 'linear') {
    throw new TypeError('TextFormat.gradient supports linear text gradients only.');
  }
  if (value.focus !== undefined) {
    throw new TypeError('TextFormat.gradient focus is not supported for linear text gradients.');
  }
  const gradient = buildGradientFill({
    stops: value.stops,
    ...(value.angleDeg === undefined ? {} : { angleDeg: value.angleDeg }),
  });
  removeRunFill(rPr);
  insertChildByRank(rPr, gradient, rprChildRank);
};

const setTextPatternFill = (rPr: XmlElement, value: TextPatternFill | null): void => {
  if (value === null) {
    removeRunFill(rPr);
    return;
  }
  if (!('preset' in value)) {
    throw new TypeError('TextFormat.patternFill unsupported diagnostic state is read-only.');
  }
  const fill = buildPatternFill(value);
  removeRunFill(rPr);
  insertChildByRank(rPr, fill, rprChildRank);
};

const setTextPictureFill = (rPr: XmlElement, value: TextPictureFill | null): void => {
  if (value === null) {
    removeRunFill(rPr);
    return;
  }
  if (!('relationshipId' in value)) {
    throw new TypeError('TextFormat.pictureFill unsupported diagnostic state is read-only.');
  }
  const fill = elem(qname('a', 'blipFill', NS.dml), {
    children: [
      elem(qname('a', 'blip', NS.dml), {
        attrs: [attr(qname('r', 'embed', NS.officeDocRels), value.relationshipId)],
      }),
      elem(qname('a', 'stretch', NS.dml), {
        children: [elem(qname('a', 'fillRect', NS.dml))],
      }),
    ],
  });
  removeRunFill(rPr);
  insertChildByRank(rPr, fill, rprChildRank);
};

/** Replace one script-specific typeface child without disturbing sibling script faces. */
const setTypeface = (
  rPr: XmlElement,
  name: ReturnType<typeof qname>,
  font: string | null,
): void => {
  rPr.children = rPr.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        c.name.localName === name.localName
      ),
  );
  if (font === null) return;
  insertChildByRank(rPr, elem(name, { attrs: [attr(ATTR_TYPEFACE, font)] }), rprChildRank);
};

const setHighlight = (rPr: XmlElement, value: string | null): void => {
  rPr.children = rPr.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'highlight'),
  );
  if (value === null) return;
  insertChildByRank(
    rPr,
    elem(NAME_HIGHLIGHT, { children: [buildColorElement(value)] }),
    rprChildRank,
  );
};

const setUnderlineLine = (rPr: XmlElement, value: TextUnderlineLine): void => {
  if (value.kind === 'unsupported') {
    throw new TypeError('TextFormat.underlineLine unsupported diagnostic state is read-only.');
  }
  rPr.children = rPr.children.filter(
    (child) =>
      !(
        child.kind === 'element' &&
        child.name.namespaceURI === NS.dml &&
        ['uLnTx', 'uLn'].includes(child.name.localName)
      ),
  );
  if (value.kind === 'followText') {
    insertChildByRank(rPr, elem(NAME_U_LN_TX), rprChildRank);
    return;
  }
  const line = elem(qname('a', 'uLn', NS.dml));
  applyTextLineProperties(line, value);
  insertChildByRank(rPr, line, rprChildRank);
};

const setUnderlineFill = (rPr: XmlElement, value: TextUnderlineFill): void => {
  if (value.kind === 'unsupported') {
    throw new TypeError('TextFormat.underlineFill unsupported diagnostic state is read-only.');
  }
  let fill: XmlElement;
  if (value.kind === 'followText') fill = elem(NAME_U_FILL_TX);
  else {
    let child: XmlElement;
    if (value.kind === 'none') child = elem(NAME_NO_FILL);
    else if (value.kind === 'solid') {
      child = elem(NAME_SOLID_FILL, { children: [buildColorElement(value.color)] });
    } else if (value.kind === 'gradient') {
      child = buildGradientFill(value);
    } else if (value.kind === 'pattern') {
      child = buildPatternFill(value);
    } else if (value.kind === 'group') {
      child = elem(qname('a', 'grpFill', NS.dml));
    } else {
      child = elem(qname('a', 'blipFill', NS.dml), {
        children: [
          elem(qname('a', 'blip', NS.dml), {
            attrs: [attr(qname('r', 'embed', NS.officeDocRels), value.relationshipId)],
          }),
          elem(qname('a', 'stretch', NS.dml), {
            children: [elem(qname('a', 'fillRect', NS.dml))],
          }),
        ],
      });
    }
    fill = elem(NAME_U_FILL, { children: [child] });
  }
  rPr.children = rPr.children.filter(
    (child) =>
      !(
        child.kind === 'element' &&
        child.name.namespaceURI === NS.dml &&
        ['uFillTx', 'uFill'].includes(child.name.localName)
      ),
  );
  insertChildByRank(rPr, fill, rprChildRank);
};

const finiteShadowPoint = (value: number, label: string, allowNegative: boolean): number => {
  if (!Number.isFinite(value) || (!allowNegative && value < 0)) {
    throw new RangeError(
      `TextFormat.textShadow ${label} must be a finite${allowNegative ? '' : ' non-negative'} point value.`,
    );
  }
  const emu = Math.round(value * 12_700);
  if (!Number.isSafeInteger(emu)) {
    throw new RangeError(
      `TextFormat.textShadow ${label} is outside the supported DrawingML coordinate range.`,
    );
  }
  return emu;
};

const setTextShadow = (rPr: XmlElement, value: TextShadow | null): void => {
  const removeTextEffects = (): void => {
    rPr.children = rPr.children.filter(
      (child) =>
        !(
          child.kind === 'element' &&
          child.name.namespaceURI === NS.dml &&
          (child.name.localName === 'effectLst' || child.name.localName === 'effectDag')
        ),
    );
  };
  if (value === null) {
    removeTextEffects();
    return;
  }
  if (!('color' in value)) {
    throw new TypeError('TextFormat.textShadow unsupported diagnostic state is read-only.');
  }
  if (!/^#[\dA-Fa-f]{6}$/u.test(value.color)) {
    throw new TypeError('TextFormat.textShadow color must be #RRGGBB.');
  }
  const blurRad = finiteShadowPoint(value.blurPt, 'blurPt', false);
  const offsetXEmu = finiteShadowPoint(value.offsetXPt, 'offsetXPt', true);
  const offsetYEmu = finiteShadowPoint(value.offsetYPt, 'offsetYPt', true);
  const dist = Math.round(Math.hypot(offsetXEmu, offsetYEmu));
  const direction = Math.round(
    (((((Math.atan2(offsetYEmu, offsetXEmu) * 180) / Math.PI) % 360) + 360) % 360) * 60_000,
  );
  let color = value.color.toUpperCase();
  if (value.opacity !== undefined) {
    if (!Number.isFinite(value.opacity) || value.opacity < 0 || value.opacity > 1) {
      throw new RangeError('TextFormat.textShadow opacity must be a finite number in [0, 1].');
    }
    const alphaByte = Math.round(value.opacity * 255);
    const cssOpacity = alphaByte / 255;
    if (Math.abs(value.opacity - cssOpacity) > 0.000_005) {
      throw new RangeError(
        'TextFormat.textShadow opacity must be exactly representable as a CSS alpha byte.',
      );
    }
    color += alphaByte.toString(16).padStart(2, '0').toUpperCase();
  }
  const shadow = elem(NAME_OUTER_SHDW, {
    attrs: [
      attr(ATTR_BLUR_RAD, String(blurRad)),
      attr(ATTR_DIST, String(dist)),
      attr(ATTR_DIR, String(direction)),
      attr(ATTR_ALGN, 'tl'),
      attr(ATTR_ROT_WITH_SHAPE, '0'),
    ],
    children: [buildColorElement(color)],
  });
  removeTextEffects();
  insertChildByRank(rPr, elem(NAME_EFFECT_LST, { children: [shadow] }), rprChildRank);
};

const setTextEffects = (rPr: XmlElement, value: readonly Effect[] | null): void => {
  rPr.children = rPr.children.filter(
    (child) =>
      !(
        child.kind === 'element' &&
        child.name.namespaceURI === NS.dml &&
        (child.name.localName === 'effectLst' || child.name.localName === 'effectDag')
      ),
  );
  if (value === null) return;
  const effectLst = buildEffectList(value);
  if (effectLst !== null) insertChildByRank(rPr, effectLst, rprChildRank);
};

/** Apply the editable CT_LineProperties subset to one text line host. */
const applyTextLineProperties = (
  line: XmlElement,
  value: TextOutline | Exclude<TextUnderlineLine, { readonly kind: 'followText' }>,
): void => {
  if (value.kind === 'unsupported') {
    throw new TypeError('Unsupported text line diagnostic state is read-only.');
  }
  const widthEmu = value.widthPt === undefined ? undefined : Math.round(value.widthPt * 12_700);
  applyLineStyle(line, {
    ...(value.kind === 'solid'
      ? { fill: { kind: 'solid' as const, color: value.color } }
      : value.kind === 'none'
        ? { fill: { kind: 'none' as const } }
        : {}),
    ...(widthEmu === undefined ? {} : { widthEmu }),
    ...(value.cap === undefined ? {} : { cap: value.cap }),
    ...(value.dash === undefined ? {} : { dash: value.dash }),
    ...(value.customDash === undefined ? {} : { customDash: value.customDash }),
    ...(value.join === undefined ? {} : { join: value.join }),
    ...(value.compound === undefined ? {} : { compound: value.compound }),
    ...(value.alignment === undefined ? {} : { alignment: value.alignment }),
    ...(value.head === undefined ? {} : { head: value.head }),
    ...(value.tail === undefined ? {} : { tail: value.tail }),
  });
  if (value.kind === 'gradient') {
    line.children.unshift(buildGradientFill(value));
  } else if (value.kind === 'pattern') {
    line.children.unshift(buildPatternFill(value));
  }
};

/** Ensure the run outline exists in the first CT_TextCharacterProperties child slot. */
const ensureRunOutline = (rPr: XmlElement): XmlElement => {
  const existing = firstChildElement(rPr, qname('a', 'ln', NS.dml));
  if (existing !== null) return existing;
  const outline = elem(qname('a', 'ln', NS.dml));
  insertChildByRank(rPr, outline, rprChildRank);
  return outline;
};

/** Mutates `rPr` in place per `format`. */
export const applyRunFormat = (rPr: XmlElement, format: TextFormat): void => {
  const directPaints = [
    format.color,
    format.gradient,
    format.patternFill,
    format.pictureFill,
  ].filter((value) => value !== undefined && value !== null);
  if (directPaints.length > 1) {
    throw new TypeError(
      'TextFormat accepts at most one direct text fill: color, gradient, patternFill, or pictureFill.',
    );
  }
  if (format.textShadow !== undefined && format.effects !== undefined) {
    throw new TypeError(
      'TextFormat.textShadow and TextFormat.effects cannot be authored together.',
    );
  }
  let attrs = rPr.attrs;
  if (format.size !== undefined) {
    // Hundredths of a point per the schema (ST_TextFontSize: 1..4000 pt).
    const sz = fontSizeHundredthPt(format.size * 100, 'setShapeRunFormat: size');
    attrs = setOrRemoveAttr(attrs, ATTR_SZ, String(sz));
  }
  if (format.bold !== undefined) {
    attrs = setOrRemoveAttr(attrs, ATTR_B, format.bold ? '1' : '0');
  }
  if (format.italic !== undefined) {
    attrs = setOrRemoveAttr(attrs, ATTR_I, format.italic ? '1' : '0');
  }
  if (format.underline !== undefined) {
    const value =
      format.underline === false || format.underline === 'noUnderline'
        ? 'none'
        : format.underline === true
          ? 'sng'
          : format.underline;
    if (!TEXT_UNDERLINE_TYPES.has(value)) {
      throw new TypeError(
        `TextFormat.underline must be a valid ST_TextUnderlineType token; got ${JSON.stringify(value)}.`,
      );
    }
    attrs = setOrRemoveAttr(attrs, ATTR_U, value);
  }
  if (format.strike !== undefined) {
    const value =
      format.strike === false ? 'noStrike' : format.strike === true ? 'sngStrike' : format.strike;
    attrs = setOrRemoveAttr(attrs, ATTR_STRIKE, value);
  }
  if (format.spc !== undefined) {
    const spc = textPointSpacing(format.spc, 'setShapeRunFormat: spc');
    attrs = setOrRemoveAttr(attrs, ATTR_SPC, String(spc));
  }
  if (format.kern !== undefined) {
    const kern = textNonNegativePoint(format.kern, 'setShapeRunFormat: kern');
    attrs = setOrRemoveAttr(attrs, ATTR_KERN, String(kern));
  }
  if (format.baseline !== undefined) {
    // ST_Percentage; we accept the unit-fraction form on the public API
    // and serialize as the on-the-wire hundredths-of-percent integer.
    const pct = Math.round(format.baseline * 100000);
    attrs = setOrRemoveAttr(attrs, ATTR_BASELINE, String(pct));
  }
  if (format.cap !== undefined) {
    attrs = setOrRemoveAttr(attrs, ATTR_CAP, format.cap);
  }
  if (format.language !== undefined) {
    attrs = setOrRemoveAttr(attrs, ATTR_LANGUAGE, format.language);
  }
  if (format.alternativeLanguage !== undefined) {
    attrs = setOrRemoveAttr(attrs, ATTR_ALTERNATIVE_LANGUAGE, format.alternativeLanguage);
  }
  rPr.attrs = attrs;

  if (format.font !== undefined) setTypeface(rPr, NAME_LATIN, format.font);
  if (format.fontEastAsian !== undefined) setTypeface(rPr, NAME_EA, format.fontEastAsian);
  if (format.fontComplexScript !== undefined) setTypeface(rPr, NAME_CS, format.fontComplexScript);
  if (format.fontSymbol !== undefined) {
    setTypeface(rPr, qname('a', 'sym', NS.dml), format.fontSymbol);
  }
  if (format.color !== undefined) setSolidFill(rPr, format.color);
  if (format.gradient !== undefined) setTextGradient(rPr, format.gradient);
  if (format.patternFill !== undefined) setTextPatternFill(rPr, format.patternFill);
  if (format.pictureFill !== undefined) setTextPictureFill(rPr, format.pictureFill);
  if (format.highlight !== undefined) setHighlight(rPr, format.highlight);
  if (format.underlineLine !== undefined) setUnderlineLine(rPr, format.underlineLine);
  if (format.underlineFill !== undefined) setUnderlineFill(rPr, format.underlineFill);
  if (format.effects !== undefined) setTextEffects(rPr, format.effects);
  if (format.textShadow !== undefined) setTextShadow(rPr, format.textShadow);
  if (format.outline !== undefined) {
    if (format.outline.kind === 'unsupported') {
      throw new TypeError('TextFormat.outline kind unsupported is read-only diagnostic state.');
    }
    applyTextLineProperties(ensureRunOutline(rPr), format.outline);
  }
};

/** Mutates only the explicitly supplied direct run-state attributes. */
export const applyRunState = (rPr: XmlElement, state: TextRunState): void => {
  let attrs = rPr.attrs;
  const booleanAttributes = [
    ['normalizeHeight', ATTR_NORMALIZE_HEIGHT],
    ['noProof', ATTR_NO_PROOF],
    ['dirty', ATTR_DIRTY],
    ['error', ATTR_ERROR],
    ['smartTagClean', ATTR_SMART_TAG_CLEAN],
  ] as const;
  for (const [key, name] of booleanAttributes) {
    const value = state[key];
    if (value === undefined) continue;
    if (typeof value !== 'boolean') throw new TypeError(`TextRunState.${key} must be a boolean.`);
    attrs = setOrRemoveAttr(attrs, name, value ? '1' : '0');
  }
  if (state.smartTagId !== undefined) {
    if (
      !Number.isInteger(state.smartTagId) ||
      state.smartTagId < 0 ||
      state.smartTagId > 0xffffffff
    ) {
      throw new RangeError('TextRunState.smartTagId must be an unsigned 32-bit integer.');
    }
    attrs = setOrRemoveAttr(attrs, ATTR_SMART_TAG_ID, String(state.smartTagId));
  }
  rPr.attrs = attrs;
};

/**
 * Walks `txBody`, ensuring every `<a:r>` has an `<a:rPr>` carrying the
 * supplied format. Existing run-property attributes not addressed by
 * `format` are preserved.
 */
export const applyFormatToAllRuns = (txBody: XmlElement, format: TextFormat): void => {
  // Walk depth-first; runs live two levels deep (txBody > p > r).
  for (const p of txBody.children) {
    if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p') {
      continue;
    }
    for (const r of p.children) {
      if (r.kind !== 'element' || r.name.namespaceURI !== NS.dml || r.name.localName !== 'r') {
        continue;
      }
      let rPr = firstChildElement(r, NAME_RPR);
      if (rPr === null) {
        rPr = elem(NAME_RPR);
        // rPr must be the first child of the run per the schema.
        r.children.unshift(rPr);
      }
      applyRunFormat(rPr, format);
    }
  }
  // Force-touch a NAME_R reference so it isn't elided as unused.
  void NAME_R;
};
