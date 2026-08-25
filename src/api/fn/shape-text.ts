// Shape mutation: text body, autofit, margins, wrap, anchor.

import {
  getShapePlaceholderIdx,
  getShapePlaceholderType,
  matchPlaceholderShape,
} from './shape-read-base.ts';
import { getSlideLayout } from './shape-slide-read.ts';
import {
  type BulletStyle,
  type ParagraphAlignment,
  type TextFormat,
  applyAlignmentToAllParagraphs,
  applyBulletToAllParagraphs,
  applyFormatToAllRuns,
  setTextBody,
} from '../../internal/drawingml/index.ts';
import {
  angle60000,
  emuCoordinate32,
  emuPositiveCoordinate32,
  textColumnCount,
} from '../../internal/bounds.ts';
import { partName, resolveTarget } from '../../internal/opc/index.ts';
import { REL_TYPES, readShapeTreeFromCsldRoot } from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { commitAndRefresh, decode, requireTxBody } from './_helpers.ts';
const NAME_TX_BODY = qname('p', 'txBody', NS.pml);

// ---------------------------------------------------------------------------
// Shape mutation — text.

/**
 * Replaces the shape's visible text with `value`. Newlines start a new
 * paragraph. Existing run/paragraph properties are preserved so font,
 * color, size, alignment, and bullet style stay intact.
 */
export const setShapeText = (
  shape: SlideShapeData,
  value: string,
  options: { bullets?: BulletStyle } = {},
): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape') {
    throw new Error(
      `setShapeText only works on text-bearing shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (txBody === null) {
    throw new Error(`shape "${shape[SHAPE_SNAPSHOT].name}" has no <p:txBody>`);
  }
  setTextBody(txBody, value);
  if (options.bullets !== undefined) {
    applyBulletToAllParagraphs(txBody, options.bullets);
  }
  commitAndRefresh(shape);
};

/**
 * Appends `value` to the shape's existing text on a new line. The
 * shape's existing run / paragraph formatting is preserved by
 * `setTextBody`; the new paragraph inherits the same template.
 *
 * Equivalent to `setShapeText(shape, getShapeText(shape) + '\n' + value)`,
 * minus the leading newline when there was no existing text.
 */
export const appendShapeText = (shape: SlideShapeData, value: string): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'shape') {
    throw new Error(
      `appendShapeText only works on text-bearing shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (txBody === null) {
    throw new Error(`shape "${shape[SHAPE_SNAPSHOT].name}" has no <p:txBody>`);
  }
  const existing = shape[SHAPE_SNAPSHOT].text;
  const combined = existing.length === 0 ? value : `${existing}\n${value}`;
  setTextBody(txBody, combined);
  commitAndRefresh(shape);
};

/** Vertical text-body anchor values stored by `bodyPr@anchor`. */
export type TextAnchor = 'top' | 'center' | 'bottom' | 'justified' | 'distributed';

const TEXT_ANCHOR_TOKEN_BY_VALUE = {
  top: 't',
  center: 'ctr',
  bottom: 'b',
  justified: 'just',
  distributed: 'dist',
} as const satisfies Readonly<Record<TextAnchor, string>>;

const TEXT_ANCHOR_BY_TOKEN = {
  t: 'top',
  ctr: 'center',
  b: 'bottom',
  just: 'justified',
  dist: 'distributed',
} as const satisfies Readonly<Record<(typeof TEXT_ANCHOR_TOKEN_BY_VALUE)[TextAnchor], TextAnchor>>;

/** Parse one native anchor token without widening unknown XML values into the public union. */
const textAnchorFromToken = (token: string | null): TextAnchor | null =>
  token === null
    ? null
    : (TEXT_ANCHOR_BY_TOKEN[token as keyof typeof TEXT_ANCHOR_BY_TOKEN] ?? null);

const NAME_A_BODY_PR = qname('a', 'bodyPr', NS.dml);

/**
 * Word wrap mode on a text body. `'square'` (PowerPoint default for
 * textboxes) wraps lines at the shape's width; `'none'` lets text
 * overflow horizontally.
 */
export type TextWrap = 'none' | 'square';

/** Horizontal text overflow stored by `bodyPr@horzOverflow`. */
export type TextHorizontalOverflow = 'overflow' | 'clip';

/** Vertical text overflow stored by `bodyPr@vertOverflow`. */
export type TextVerticalOverflow = 'overflow' | 'ellipsis' | 'clip';

/** Auto-fit mode on a text body. */
export type TextAutoFit =
  | 'none' // <a:noAutofit/>
  | 'normal' // <a:normAutofit/> — shrink text to fit
  | 'shape'; // <a:spAutoFit/> — resize shape to fit text

/** PowerPoint-computed reduction ratios stored on `normAutofit`; both values use the inclusive `[0, 1]` range. */
export interface TextAutoFitParams {
  fontScale: number;
  lnSpcReduction: number;
}

/** Parsed native `normAutofit` ratios, which Strict percent strings may place outside the safe authoring range. */
export interface TextAutoFitRawParams {
  fontScale: number;
  lnSpcReduction: number;
}

const AUTO_FIT_LOCALS = new Set(['noAutofit', 'normAutofit', 'spAutoFit']);

const requireBodyPr = (shape: SlideShapeData): XmlElement => {
  const txBody = requireTxBody(shape);
  let bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (bodyPr === null) {
    bodyPr = elem(NAME_A_BODY_PR);
    txBody.children.unshift(bodyPr);
  }
  return bodyPr;
};

/**
 * Sets the text-body word-wrap mode.
 *
 *   - `'square'` writes `wrap="square"` — PowerPoint default for textboxes.
 *   - `'none'`   writes `wrap="none"`  — text can overflow horizontally.
 *
 * Throws for non-text-bearing shape kinds.
 */
export const setShapeTextWrap = (shape: SlideShapeData, wrap: TextWrap): void => {
  const bodyPr = requireBodyPr(shape);
  const ATTR_WRAP = qname('', 'wrap', '');
  bodyPr.attrs = bodyPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'wrap'),
  );
  bodyPr.attrs.push(attr(ATTR_WRAP, wrap));
  commitAndRefresh(shape);
};

/** Reads back the bodyPr `wrap` attribute, or `null` when absent. */
export const getShapeTextWrap = (shape: SlideShapeData): TextWrap | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const v = getAttrValue(bodyPr, qname('', 'wrap', ''));
  if (v === 'none' || v === 'square') return v;
  return null;
};

/** Writes or clears one enumerated text-body attribute. */
const setShapeTextBodyToken = (
  shape: SlideShapeData,
  localName: 'horzOverflow' | 'vertOverflow',
  value: string | null,
): void => {
  const bodyPr = requireBodyPr(shape);
  bodyPr.attrs = bodyPr.attrs.filter(
    (candidate) => !(candidate.name.namespaceURI === '' && candidate.name.localName === localName),
  );
  if (value !== null) bodyPr.attrs.push(attr(qname('', localName, ''), value));
  commitAndRefresh(shape);
};

/** Reads horizontal text overflow, or `null` when it is not authored directly. */
export const getShapeTextHorizontalOverflow = (
  shape: SlideShapeData,
): TextHorizontalOverflow | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const value = getAttrValue(bodyPr, qname('', 'horzOverflow', ''));
  return value === 'overflow' || value === 'clip' ? value : null;
};

/** Sets or clears horizontal text overflow (`bodyPr@horzOverflow`). */
export const setShapeTextHorizontalOverflow = (
  shape: SlideShapeData,
  value: TextHorizontalOverflow | null,
): void => setShapeTextBodyToken(shape, 'horzOverflow', value);

/** Reads vertical text overflow, or `null` when it is not authored directly. */
export const getShapeTextVerticalOverflow = (
  shape: SlideShapeData,
): TextVerticalOverflow | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const value = getAttrValue(bodyPr, qname('', 'vertOverflow', ''));
  return value === 'overflow' || value === 'ellipsis' || value === 'clip' ? value : null;
};

/** Sets or clears vertical text overflow (`bodyPr@vertOverflow`). */
export const setShapeTextVerticalOverflow = (
  shape: SlideShapeData,
  value: TextVerticalOverflow | null,
): void => setShapeTextBodyToken(shape, 'vertOverflow', value);

/**
 * Sets the text-body auto-fit mode:
 *
 *   - `'none'`   → `<a:noAutofit/>`
 *   - `'normal'` → `<a:normAutofit/>`   shrink text to fit the shape
 *   - `'shape'`  → `<a:spAutoFit/>`     grow the shape to fit text
 *
 * `params` is valid only for normal autofit and writes the native
 * `fontScale` / `lnSpcReduction` percentages. Replaces any prior auto-fit
 * child on `<a:bodyPr>`. Throws before mutation for invalid ratios, parameter
 * use on another mode, or non-text-bearing shape kinds.
 */
export const setShapeTextAutoFit = (
  shape: SlideShapeData,
  mode: TextAutoFit,
  params?: TextAutoFitParams,
): void => {
  if (params !== undefined && mode !== 'normal') {
    throw new RangeError('Text autofit parameters are supported only for normal mode.');
  }
  if (params !== undefined) {
    assertTextAutoFitRatio(params.fontScale, 'fontScale');
    assertTextAutoFitRatio(params.lnSpcReduction, 'lnSpcReduction');
  }
  const bodyPr = requireBodyPr(shape);
  bodyPr.children = bodyPr.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        AUTO_FIT_LOCALS.has(c.name.localName)
      ),
  );
  const local = mode === 'none' ? 'noAutofit' : mode === 'normal' ? 'normAutofit' : 'spAutoFit';
  const autoFit = elem(qname('a', local, NS.dml));
  if (params !== undefined) {
    autoFit.attrs.push(
      attr(qname('', 'fontScale', ''), serializeTextAutoFitRatio(params.fontScale, 'fontScale')),
      attr(
        qname('', 'lnSpcReduction', ''),
        serializeTextAutoFitRatio(params.lnSpcReduction, 'lnSpcReduction'),
      ),
    );
  }
  bodyPr.children.push(autoFit);
  commitAndRefresh(shape);
};

/** Reject values that cannot be represented by the public normal-autofit ratio contract. */
const assertTextAutoFitRatio = (value: number, name: keyof TextAutoFitParams): void => {
  if (!isTextAutoFitRatio(value)) {
    throw new RangeError(`Text autofit ${name} must be a finite ratio from 0 through 1.`);
  }
};

/** Return whether a value satisfies the safe authored normal-autofit ratio contract. */
const isTextAutoFitRatio = (value: number): boolean =>
  Number.isFinite(value) && value >= 0 && value <= 1;

/** Use Strict percent syntax where Transitional font-scale integers cannot represent the public zero-based range. */
const serializeTextAutoFitRatio = (value: number, name: keyof TextAutoFitParams): string => {
  const units = Math.round(value * 100_000);
  if (name !== 'fontScale' || units >= 1_000) return String(units);
  const percent = (value * 100).toFixed(10).replace(/0+$/u, '').replace(/\.$/u, '');
  return `${percent}%`;
};

/**
 * Reads back the bodyPr auto-fit child, or `null` when none is
 * present (PowerPoint applies a layout-inherited default in that case).
 */
export const getShapeTextAutoFit = (shape: SlideShapeData): TextAutoFit | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  for (const c of bodyPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'noAutofit') return 'none';
    if (c.name.localName === 'normAutofit') return 'normal';
    if (c.name.localName === 'spAutoFit') return 'shape';
  }
  return null;
};

/**
 * Reads the raw scale parameters stored on `<a:normAutofit>`. Returns `null`
 * if the body doesn't carry normal autofit or both attributes are absent.
 * When only one attribute exists, the other field uses its schema default.
 * Strict percent strings may parse outside `[0, 1]`; importers use this raw
 * result to diagnose values that the safe authoring contract cannot preserve.
 * Invalid numeric lexemes also fall back to the corresponding schema default.
 */
export const getShapeTextAutoFitParamsRaw = (
  shape: SlideShapeData,
): TextAutoFitRawParams | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  for (const c of bodyPr.children) {
    if (
      c.kind === 'element' &&
      c.name.namespaceURI === NS.dml &&
      c.name.localName === 'normAutofit'
    ) {
      const fsRaw = getAttrValue(c, qname('', 'fontScale', ''));
      const lsRaw = getAttrValue(c, qname('', 'lnSpcReduction', ''));
      if (fsRaw === null && lsRaw === null) return null;
      return {
        fontScale: parseTextAutoFitRatio(fsRaw, 1),
        lnSpcReduction: parseTextAutoFitRatio(lsRaw, 0),
      };
    }
  }
  return null;
};

/**
 * Reads PowerPoint-computed normal-autofit parameters only when both values
 * satisfy the safe public `[0, 1]` ratio contract. Returns `null` for absent,
 * bare, or out-of-contract native parameters:
 *
 *   - `fontScale`     — multiply every run's font size by this. Default `1`.
 *   - `lnSpcReduction` — subtract from the line-height ratio. Default `0`.
 *
 * Companion to `getShapeTextAutoFit`. Renderers that want to match
 * PowerPoint's actual on-screen text size apply these factors to the
 * authored font sizes; without them, every long title overflows.
 */
export const getShapeTextAutoFitParams = (shape: SlideShapeData): TextAutoFitParams | null => {
  const params = getShapeTextAutoFitParamsRaw(shape);
  return params !== null &&
    isTextAutoFitRatio(params.fontScale) &&
    isTextAutoFitRatio(params.lnSpcReduction)
    ? params
    : null;
};

/** Normalize Transitional 1/1000-percent integers and Strict percent lexemes to a unit ratio. */
const parseTextAutoFitRatio = (value: string | null, fallback: number): number => {
  if (value === null) return fallback;
  const parsed = value.endsWith('%') ? Number(value.slice(0, -1)) / 100 : Number(value) / 100_000;
  return Number.isFinite(parsed) ? parsed : fallback;
};

/**
 * Reads back the vertical text anchor on the shape's `<a:bodyPr>`.
 * Maps the ECMA-376 tokens back to the public union:
 *
 *   `'t'` → `'top'`, `'ctr'` → `'center'`, `'b'` → `'bottom'`,
 *   `'just'` → `'justified'`, `'dist'` → `'distributed'`
 *
 * Returns `null` when the bodyPr is absent or has no anchor attribute.
 */
export const getShapeTextAnchor = (shape: SlideShapeData): TextAnchor | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  return textAnchorFromToken(getAttrValue(bodyPr, qname('', 'anchor', '')));
};

/**
 * Reads back the internal margins of the shape's text frame. Sides
 * that are absent in the XML default to `null` (PowerPoint applies
 * its built-in default for the missing side).
 */
/**
 * Reads the multi-column layout on a text body — `<a:bodyPr
 * numCol="N" spcCol="EMU"/>`. Returns `null` when columns aren't
 * configured (the default single column). `gapEmu` is the
 * inter-column gap in EMU; omitted when `<a:bodyPr>` has no
 * `spcCol` attribute.
 */
export const getShapeTextColumns = (
  shape: SlideShapeData,
): { count: number; gapEmu?: number } | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const numColRaw = getAttrValue(bodyPr, qname('', 'numCol', ''));
  if (numColRaw === null) return null;
  const count = Number.parseInt(numColRaw, 10);
  if (!Number.isFinite(count) || count < 2) return null;
  const gapRaw = getAttrValue(bodyPr, qname('', 'spcCol', ''));
  if (gapRaw !== null) {
    const g = Number.parseInt(gapRaw, 10);
    if (Number.isFinite(g)) return { count, gapEmu: g };
  }
  return { count };
};

/**
 * Sets the multi-column layout on the shape's text body — writes
 * `<a:bodyPr numCol="N" [spcCol="EMU"]/>`. Pass `null` to clear both
 * attributes so the text body falls back to PowerPoint's default
 * single column. `count` must be in `2..16` (ST_TextColumnCount caps at
 * 16, and single column is the `null` default). `gapEmu`, when omitted,
 * removes any prior `spcCol`. Throws for non-text-bearing shape kinds.
 */
export const setShapeTextColumns = (
  shape: SlideShapeData,
  columns: { count: number; gapEmu?: number } | null,
): void => {
  const bodyPr = requireBodyPr(shape);
  bodyPr.attrs = bodyPr.attrs.filter(
    (a) =>
      !(
        a.name.namespaceURI === '' &&
        (a.name.localName === 'numCol' || a.name.localName === 'spcCol')
      ),
  );
  if (columns !== null) {
    if (columns.count < 2) {
      throw new Error(
        `setShapeTextColumns: count must be >= 2 (single column is the default — pass null instead). Got ${columns.count}.`,
      );
    }
    // ST_TextColumnCount caps at 16; spcCol is ST_PositiveCoordinate32.
    const numCol = textColumnCount(columns.count, 'setShapeTextColumns: count');
    bodyPr.attrs.push(attr(qname('', 'numCol', ''), String(numCol)));
    if (columns.gapEmu !== undefined) {
      const spcCol = emuPositiveCoordinate32(columns.gapEmu, 'setShapeTextColumns: gapEmu');
      bodyPr.attrs.push(attr(qname('', 'spcCol', ''), String(spcCol)));
    }
  }
  commitAndRefresh(shape);
};

/** Reads one explicit boolean attribute from `<a:bodyPr>`. */
const getShapeTextBodyBoolean = (
  shape: SlideShapeData,
  localName: 'rtlCol' | 'upright' | 'anchorCtr' | 'compatLnSpc',
): boolean | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const value = getAttrValue(bodyPr, qname('', localName, ''));
  if (value === '1' || value === 'true') return true;
  if (value === '0' || value === 'false') return false;
  return null;
};

/** Writes or clears one explicit boolean attribute on `<a:bodyPr>`. */
const setShapeTextBodyBoolean = (
  shape: SlideShapeData,
  localName: 'rtlCol' | 'upright' | 'anchorCtr' | 'compatLnSpc',
  value: boolean | null,
): void => {
  const bodyPr = requireBodyPr(shape);
  bodyPr.attrs = bodyPr.attrs.filter(
    (candidate) => !(candidate.name.namespaceURI === '' && candidate.name.localName === localName),
  );
  if (value !== null) bodyPr.attrs.push(attr(qname('', localName, ''), value ? '1' : '0'));
  commitAndRefresh(shape);
};

/** Reads whether multi-column text flows from right to left (`bodyPr@rtlCol`). */
export const getShapeTextRtlColumns = (shape: SlideShapeData): boolean | null =>
  getShapeTextBodyBoolean(shape, 'rtlCol');

/** Sets or clears right-to-left text-column flow (`bodyPr@rtlCol`). */
export const setShapeTextRtlColumns = (shape: SlideShapeData, value: boolean | null): void =>
  setShapeTextBodyBoolean(shape, 'rtlCol', value);

/** Reads whether PowerPoint keeps text upright through shape rotation (`bodyPr@upright`). */
export const getShapeTextUpright = (shape: SlideShapeData): boolean | null =>
  getShapeTextBodyBoolean(shape, 'upright');

/** Sets or clears PowerPoint's keep-text-upright flag (`bodyPr@upright`). */
export const setShapeTextUpright = (shape: SlideShapeData, value: boolean | null): void =>
  setShapeTextBodyBoolean(shape, 'upright', value);

/** Reads whether each anchored line is centered within the text body (`bodyPr@anchorCtr`). */
export const getShapeTextAnchorCentering = (shape: SlideShapeData): boolean | null =>
  getShapeTextBodyBoolean(shape, 'anchorCtr');

/** Sets or clears per-line anchor centering (`bodyPr@anchorCtr`). */
export const setShapeTextAnchorCentering = (shape: SlideShapeData, value: boolean | null): void =>
  setShapeTextBodyBoolean(shape, 'anchorCtr', value);

/** Reads the legacy compatible line-spacing flag (`bodyPr@compatLnSpc`). */
export const getShapeTextCompatibilityLineSpacing = (shape: SlideShapeData): boolean | null =>
  getShapeTextBodyBoolean(shape, 'compatLnSpc');

/** Sets or clears the legacy compatible line-spacing flag (`bodyPr@compatLnSpc`). */
export const setShapeTextCompatibilityLineSpacing = (
  shape: SlideShapeData,
  value: boolean | null,
): void => setShapeTextBodyBoolean(shape, 'compatLnSpc', value);

/**
 * Reads the shape's text-body rotation from `<a:bodyPr rot="N"/>`.
 * `rot` is stored in 60000ths of a degree (OOXML angle units); the
 * returned value is in degrees. Positive values rotate clockwise per
 * PowerPoint's convention. Returns `null` when the attribute is
 * absent.
 *
 * Distinct from the shape's own `<p:xfrm rot=…>` (the geometry
 * rotation surfaced via the shape's `rotation`); `bodyPr rot` rotates
 * the text body inside the shape without rotating the shape itself.
 */
export const getShapeTextBodyRotationDeg = (shape: SlideShapeData): number | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const v = getAttrValue(bodyPr, qname('', 'rot', ''));
  if (v === null) return null;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n / 60000 : null;
};

/**
 * Sets the shape's text-body rotation (`<a:bodyPr rot="N"/>`), measured
 * in degrees. Positive rotates clockwise per PowerPoint's convention.
 * Passing `null` clears the attribute so the shape inherits the default
 * (`0`). Throws for non-text-bearing shape kinds.
 *
 * Companion to `setShapeRotation`, which rotates the *whole* shape
 * via `<p:xfrm rot>`. `bodyPr rot` rotates only the text inside.
 */
export const setShapeTextBodyRotationDeg = (
  shape: SlideShapeData,
  rotationDeg: number | null,
): void => {
  const bodyPr = requireBodyPr(shape);
  // Strip any prior rot attribute.
  bodyPr.attrs = bodyPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'rot'),
  );
  if (rotationDeg !== null && rotationDeg !== 0) {
    // `rot` is ST_Angle in 1/60000 degree (xsd:int); guard the *60000 overflow.
    const rot = angle60000(rotationDeg * 60000, 'setShapeTextBodyRotationDeg: rotationDeg');
    bodyPr.attrs.push(attr(qname('', 'rot', ''), String(rot)));
  }
  commitAndRefresh(shape);
};

/**
 * Reads the shape's text-direction token from `<a:bodyPr vert="…"/>`.
 * Per ECMA-376 §17.18.93 `ST_TextVerticalType`:
 *
 *   - `horz` — default left-to-right, top-to-bottom (returns `null`).
 *   - `vert` — 90° rotation, lines run top-to-bottom, columns right-to-left.
 *   - `vert270` — 270° rotation, lines top-to-bottom, columns left-to-right.
 *   - `wordArtVert` — characters not rotated, stacked vertically.
 *   - `eaVert` — East-Asian vertical: characters upright, columns right-to-left.
 *   - `mongolianVert` — Mongolian: rotated 90°, columns left-to-right.
 *   - `wordArtVertRtl` — RTL word-art stacked vertically.
 *
 * Returns `null` when the attribute is absent or set to the default
 * `horz`.
 */
export const getShapeTextDirection = (
  shape: SlideShapeData,
): 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl' | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const v = getAttrValue(bodyPr, qname('', 'vert', ''));
  if (
    v === 'vert' ||
    v === 'vert270' ||
    v === 'wordArtVert' ||
    v === 'eaVert' ||
    v === 'mongolianVert' ||
    v === 'wordArtVertRtl'
  )
    return v;
  return null;
};

/**
 * Sets the shape's text-direction via `<a:bodyPr vert="…"/>`. See
 * `getShapeTextDirection` for the meaning of each value. Passing `null`
 * (or `'horz'`) clears the attribute so the shape uses the default
 * horizontal direction. Throws for non-text-bearing shape kinds.
 */
export const setShapeTextDirection = (
  shape: SlideShapeData,
  direction:
    | 'horz'
    | 'vert'
    | 'vert270'
    | 'wordArtVert'
    | 'eaVert'
    | 'mongolianVert'
    | 'wordArtVertRtl'
    | null,
): void => {
  const bodyPr = requireBodyPr(shape);
  bodyPr.attrs = bodyPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'vert'),
  );
  if (direction !== null && direction !== 'horz') {
    bodyPr.attrs.push(attr(qname('', 'vert', ''), direction));
  }
  commitAndRefresh(shape);
};

export const getShapeTextMargins = (
  shape: SlideShapeData,
): {
  readonly left: number | null;
  readonly top: number | null;
  readonly right: number | null;
  readonly bottom: number | null;
} | null => {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (!txBody) return null;
  const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (!bodyPr) return null;
  const readSide = (local: string): number | null => {
    const v = getAttrValue(bodyPr, qname('', local, ''));
    if (v === null) return null;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n : null;
  };
  return {
    left: readSide('lIns'),
    top: readSide('tIns'),
    right: readSide('rIns'),
    bottom: readSide('bIns'),
  };
};

/**
 * Resolves the effective `<a:bodyPr>` properties — anchor, wrap, overflow,
 * anchor centering, compatibility line spacing, vertical direction, and
 * inset margins — by walking the layout / master cascade the same way
 * `getShapeRunFormatEffective` walks rPr. Returns the innermost value that
 * the cascade supplies, or `null` for properties neither the shape nor any
 * inherited placeholder authors.
 *
 * Companion to `getShapeTextAnchor` / `getShapeTextWrap` /
 * `getShapeTextDirection` / `getShapeTextMargins`, which only report the
 * literal value on the shape itself.
 */
export const getShapeBodyPrEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
): {
  anchor: TextAnchor | null;
  anchorCentering: boolean | null;
  compatibilityLineSpacing: boolean | null;
  horizontalOverflow: TextHorizontalOverflow | null;
  verticalOverflow: TextVerticalOverflow | null;
  wrap: TextWrap | null;
  vert: ReturnType<typeof getShapeTextDirection>;
  margins: { left: number | null; top: number | null; right: number | null; bottom: number | null };
} => {
  const result = {
    anchor: null as TextAnchor | null,
    anchorCentering: null as boolean | null,
    compatibilityLineSpacing: null as boolean | null,
    horizontalOverflow: null as TextHorizontalOverflow | null,
    verticalOverflow: null as TextVerticalOverflow | null,
    wrap: null as TextWrap | null,
    vert: null as ReturnType<typeof getShapeTextDirection>,
    margins: {
      left: null as number | null,
      top: null as number | null,
      right: null as number | null,
      bottom: null as number | null,
    },
  };
  const parseBodyPr = (bodyPr: XmlElement): void => {
    const bodyBoolean = (localName: 'anchorCtr' | 'compatLnSpc'): boolean | null => {
      const value = getAttrValue(bodyPr, qname('', localName, ''));
      if (value === '1' || value === 'true') return true;
      if (value === '0' || value === 'false') return false;
      return null;
    };
    if (result.anchor === null) {
      result.anchor = textAnchorFromToken(getAttrValue(bodyPr, qname('', 'anchor', '')));
    }
    if (result.anchorCentering === null) result.anchorCentering = bodyBoolean('anchorCtr');
    if (result.compatibilityLineSpacing === null) {
      result.compatibilityLineSpacing = bodyBoolean('compatLnSpc');
    }
    if (result.horizontalOverflow === null) {
      const value = getAttrValue(bodyPr, qname('', 'horzOverflow', ''));
      if (value === 'overflow' || value === 'clip') result.horizontalOverflow = value;
    }
    if (result.verticalOverflow === null) {
      const value = getAttrValue(bodyPr, qname('', 'vertOverflow', ''));
      if (value === 'overflow' || value === 'ellipsis' || value === 'clip') {
        result.verticalOverflow = value;
      }
    }
    if (result.wrap === null) {
      const w = getAttrValue(bodyPr, qname('', 'wrap', ''));
      if (w === 'square') result.wrap = 'square';
      else if (w === 'none') result.wrap = 'none';
    }
    if (result.vert === null) {
      const v = getAttrValue(bodyPr, qname('', 'vert', ''));
      if (
        v === 'vert' ||
        v === 'vert270' ||
        v === 'wordArtVert' ||
        v === 'eaVert' ||
        v === 'mongolianVert' ||
        v === 'wordArtVertRtl'
      )
        result.vert = v;
    }
    for (const side of ['l', 't', 'r', 'b'] as const) {
      const target =
        side === 'l' ? 'left' : side === 't' ? 'top' : side === 'r' ? 'right' : 'bottom';
      if (result.margins[target] !== null) continue;
      const v = getAttrValue(bodyPr, qname('', `${side}Ins`, ''));
      if (v === null) continue;
      const n = Number.parseInt(v, 10);
      if (Number.isFinite(n)) result.margins[target] = n;
    }
  };

  // 1. The shape's own bodyPr.
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
  if (txBody) {
    const bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
    if (bodyPr) parseBodyPr(bodyPr);
  }

  // 2-3. Walk layout placeholder and master placeholder bodyPr.
  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  const slide = shape[SHAPE_SLIDE];
  const layout = getSlideLayout(slide);
  if (!layout) return result;

  const findPh = (
    shapes: ReadonlyArray<{
      placeholderIdx: number | null;
      placeholderType: string | null;
      element: XmlElement;
    }>,
  ): XmlElement | null => matchPlaceholderShape(shapes, phIdx, phType)?.element ?? null;

  const layoutPhEl = findPh(layout[LAYOUT_PART].shapes);
  if (layoutPhEl) {
    const layoutTxBody = firstChildElement(layoutPhEl, NAME_TX_BODY);
    if (layoutTxBody) {
      const bodyPr = firstChildElement(layoutTxBody, NAME_A_BODY_PR);
      if (bodyPr) parseBodyPr(bodyPr);
    }
  }

  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return result;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return result;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return result;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  const masterPhEl = findPh(masterShapes);
  if (masterPhEl) {
    const masterTxBody = firstChildElement(masterPhEl, NAME_TX_BODY);
    if (masterTxBody) {
      const bodyPr = firstChildElement(masterTxBody, NAME_A_BODY_PR);
      if (bodyPr) parseBodyPr(bodyPr);
    }
  }
  return result;
};

/** Set the vertical text-body anchor, creating `bodyPr` when the shape can bear text and none exists. */
export const setShapeTextAnchor = (shape: SlideShapeData, anchor: TextAnchor): void => {
  const txBody = requireTxBody(shape);
  let bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (bodyPr === null) {
    bodyPr = elem(NAME_A_BODY_PR);
    txBody.children.unshift(bodyPr);
  }
  const token = TEXT_ANCHOR_TOKEN_BY_VALUE[anchor];
  const ATTR_ANCHOR = qname('', 'anchor', '');
  // Replace any existing anchor attribute.
  bodyPr.attrs = bodyPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'anchor'),
  );
  bodyPr.attrs.push(attr(ATTR_ANCHOR, token));
  commitAndRefresh(shape);
};

/**
 * Sets the internal margins of the shape's text frame in EMU. Each
 * side is independent; omitted sides keep their current value (or the
 * layout-inherited default when the attribute is absent).
 *
 * PowerPoint's defaults for a textbox: left/right 91440 (0.1in),
 * top/bottom 45720 (0.05in).
 *
 *   setShapeTextMargins(shape, { left: 0, right: 0 }); // flush-left text
 */
export const setShapeTextMargins = (
  shape: SlideShapeData,
  margins: { left?: number; top?: number; right?: number; bottom?: number },
): void => {
  const txBody = requireTxBody(shape);
  let bodyPr = firstChildElement(txBody, NAME_A_BODY_PR);
  if (bodyPr === null) {
    bodyPr = elem(NAME_A_BODY_PR);
    txBody.children.unshift(bodyPr);
  }
  const writes: Array<{ name: string; value: number }> = [];
  if (margins.left !== undefined) writes.push({ name: 'lIns', value: margins.left });
  if (margins.top !== undefined) writes.push({ name: 'tIns', value: margins.top });
  if (margins.right !== undefined) writes.push({ name: 'rIns', value: margins.right });
  if (margins.bottom !== undefined) writes.push({ name: 'bIns', value: margins.bottom });

  const localsToClear = new Set(writes.map((w) => w.name));
  bodyPr.attrs = bodyPr.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && localsToClear.has(a.name.localName)),
  );
  for (const w of writes) {
    // Insets are ST_Coordinate32 (xsd:int EMU).
    const emu = emuCoordinate32(w.value, `setShapeTextMargins: ${w.name}`);
    bodyPr.attrs.push(attr(qname('', w.name, ''), String(emu)));
  }
  commitAndRefresh(shape);
};

/** Sets the bullet style on every paragraph in the shape's text body. */
export const setShapeBullets = (shape: SlideShapeData, style: BulletStyle): void => {
  applyBulletToAllParagraphs(requireTxBody(shape), style);
  commitAndRefresh(shape);
};

/** Sets the horizontal alignment of every paragraph in the shape's text. */
export const setShapeAlignment = (shape: SlideShapeData, align: ParagraphAlignment): void => {
  applyAlignmentToAllParagraphs(requireTxBody(shape), align);
  commitAndRefresh(shape);
};

/**
 * Applies `format` to every run in the shape's text. Run-property
 * attributes not addressed by `format` are preserved, so partial
 * updates compose.
 */
export const setShapeTextFormat = (shape: SlideShapeData, format: TextFormat): void => {
  applyFormatToAllRuns(requireTxBody(shape), format);
  commitAndRefresh(shape);
};
