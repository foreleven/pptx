// Shape effects — `<a:effectLst>` builders.
//
// Covers the most-used PowerPoint effects: outer shadow + glow. The
// element ordering on `<p:spPr>` is `xfrm → geometry → fill → ln →
// effectLst → scene3d → sp3d → extLst`. Callers locate the right
// insertion slot using `effectInsertionIndex`.

import { emuExtent } from '../bounds.ts';
import { NS, type XmlElement, attr, elem, qname } from '../xml/index.ts';
import { buildColorElement } from './color.ts';

const NAME_EFFECT_LST = qname('a', 'effectLst', NS.dml);
const NAME_BLUR = qname('a', 'blur', NS.dml);
const NAME_OUTER_SHDW = qname('a', 'outerShdw', NS.dml);
const NAME_INNER_SHDW = qname('a', 'innerShdw', NS.dml);
const NAME_GLOW = qname('a', 'glow', NS.dml);
const NAME_REFLECTION = qname('a', 'reflection', NS.dml);
const NAME_SOFT_EDGE = qname('a', 'softEdge', NS.dml);
const NAME_ALPHA = qname('a', 'alpha', NS.dml);

const ATTR_BLUR_RAD = qname('', 'blurRad', '');
const ATTR_DIST = qname('', 'dist', '');
const ATTR_DIR = qname('', 'dir', '');
const ATTR_ALGN = qname('', 'algn', '');
const ATTR_ROT_WITH_SHAPE = qname('', 'rotWithShape', '');
const ATTR_RAD = qname('', 'rad', '');
const ATTR_GROW = qname('', 'grow', '');
const ATTR_START_ALPHA = qname('', 'stA', '');
const ATTR_END_ALPHA = qname('', 'endA', '');
const ATTR_SCALE_Y = qname('', 'sy', '');
const ATTR_VAL = qname('', 'val', '');

export interface ShadowOptions {
  /** `#RRGGBB`, bare `RRGGBB`, or scheme token. Defaults to black. */
  readonly color?: string;
  /** Edge blur in EMU. Defaults to 50800 (4pt). */
  readonly blurEmu?: number;
  /** Offset distance in EMU. Defaults to 38100 (3pt). */
  readonly offsetEmu?: number;
  /**
   * Direction in degrees, measured clockwise from the right (3 o'clock).
   * Defaults to 45° (down-right).
   */
  readonly angleDeg?: number;
  /** Opacity (0–1). Defaults to fully opaque. */
  readonly opacity?: number;
}

export interface GlowOptions {
  /** `#RRGGBB`, bare `RRGGBB`, or scheme token. */
  readonly color: string;
  /** Glow radius in EMU. Defaults to 63500 (5pt). */
  readonly radiusEmu?: number;
}

export type ShapeEffectOptions =
  | {
      readonly kind: 'blur';
      /** Blur radius in EMU. */
      readonly radiusEmu?: number;
      /** Whether the blurred bounds expand beyond the original geometry. */
      readonly grow?: boolean;
    }
  | ({ readonly kind: 'glow'; readonly opacity?: number } & GlowOptions)
  | {
      readonly kind: 'innerShdw' | 'outerShdw';
      readonly color: string;
      readonly blurEmu?: number;
      readonly distEmu?: number;
      readonly angleDeg?: number;
      readonly opacity?: number;
    }
  | {
      readonly kind: 'reflection';
      readonly blurEmu?: number;
      readonly distEmu?: number;
      readonly angleDeg?: number;
      readonly startOpacity?: number;
      readonly endOpacity?: number;
      readonly scaleY?: number;
    }
  | { readonly kind: 'softEdge'; readonly radiusEmu: number };

/**
 * Computes the index inside `host.children` where an `<a:effectLst>`
 * should be inserted to satisfy the spec's child ordering on
 * `<p:spPr>`.
 */
const effectInsertionIndex = (host: XmlElement): number => {
  for (let i = 0; i < host.children.length; i++) {
    const c = host.children[i];
    if (c?.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'scene3d' || c.name.localName === 'sp3d') return i;
    if (c.name.localName === 'extLst') return i;
  }
  return host.children.length;
};

const removeEffectLst = (host: XmlElement): void => {
  host.children = host.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'effectLst'),
  );
};

const colorWithAlpha = (color: string, opacity: number | undefined): XmlElement => {
  const base = buildColorElement(color);
  if (opacity !== undefined) {
    const amt = fixedPercentage(opacity, 'effect opacity');
    // The explicit effect opacity owns this transform when both forms are
    // present, avoiding duplicate `<a:alpha>` children in the color choice.
    base.children = base.children.filter(
      (child) =>
        !(
          child.kind === 'element' &&
          child.name.namespaceURI === NS.dml &&
          child.name.localName === 'alpha'
        ),
    );
    if (amt === 100000) return base;
    base.children.push(elem(NAME_ALPHA, { attrs: [attr(ATTR_VAL, String(amt))] }));
  }
  return base;
};

const fixedPercentage = (value: number, label: string): number => {
  if (!Number.isFinite(value) || value < 0 || value > 1) {
    throw new RangeError(`${label} must be a finite number in [0, 1], got ${value}`);
  }
  return Math.round(value * 100000);
};

const signedPercentage = (value: number, label: string): number => {
  if (!Number.isFinite(value)) throw new RangeError(`${label} must be finite, got ${value}`);
  return Math.round(value * 100000);
};

const angle = (value: number | undefined, label: string): string => {
  const degrees = value ?? 0;
  if (!Number.isFinite(degrees)) throw new RangeError(`${label} must be finite, got ${degrees}`);
  return String(Math.round((((degrees % 360) + 360) % 360) * 60000));
};

const buildComposedEffect = (effect: ShapeEffectOptions): XmlElement => {
  if (effect.kind === 'blur') {
    return elem(NAME_BLUR, {
      attrs: [
        attr(ATTR_RAD, String(emuExtent(effect.radiusEmu ?? 0, 'setShapeEffects: blur radiusEmu'))),
        attr(ATTR_GROW, effect.grow === false ? '0' : '1'),
      ],
    });
  }
  if (effect.kind === 'glow') {
    return elem(NAME_GLOW, {
      attrs: [
        attr(
          ATTR_RAD,
          String(emuExtent(effect.radiusEmu ?? 63500, 'setShapeEffects: glow radiusEmu')),
        ),
      ],
      children: [colorWithAlpha(effect.color, effect.opacity)],
    });
  }
  if (effect.kind === 'innerShdw' || effect.kind === 'outerShdw') {
    return elem(effect.kind === 'innerShdw' ? NAME_INNER_SHDW : NAME_OUTER_SHDW, {
      attrs: [
        attr(
          ATTR_BLUR_RAD,
          String(emuExtent(effect.blurEmu ?? 0, `setShapeEffects: ${effect.kind} blurEmu`)),
        ),
        attr(
          ATTR_DIST,
          String(emuExtent(effect.distEmu ?? 0, `setShapeEffects: ${effect.kind} distEmu`)),
        ),
        attr(ATTR_DIR, angle(effect.angleDeg, `setShapeEffects: ${effect.kind} angleDeg`)),
        ...(effect.kind === 'outerShdw'
          ? [attr(ATTR_ALGN, 'tl'), attr(ATTR_ROT_WITH_SHAPE, '0')]
          : []),
      ],
      children: [colorWithAlpha(effect.color, effect.opacity)],
    });
  }
  if (effect.kind === 'reflection') {
    return elem(NAME_REFLECTION, {
      attrs: [
        attr(
          ATTR_BLUR_RAD,
          String(emuExtent(effect.blurEmu ?? 0, 'setShapeEffects: reflection blurEmu')),
        ),
        attr(
          ATTR_DIST,
          String(emuExtent(effect.distEmu ?? 0, 'setShapeEffects: reflection distEmu')),
        ),
        attr(ATTR_DIR, angle(effect.angleDeg, 'setShapeEffects: reflection angleDeg')),
        ...(effect.startOpacity === undefined
          ? []
          : [
              attr(
                ATTR_START_ALPHA,
                String(fixedPercentage(effect.startOpacity, 'reflection startOpacity')),
              ),
            ]),
        ...(effect.endOpacity === undefined
          ? []
          : [
              attr(
                ATTR_END_ALPHA,
                String(fixedPercentage(effect.endOpacity, 'reflection endOpacity')),
              ),
            ]),
        ...(effect.scaleY === undefined
          ? []
          : [attr(ATTR_SCALE_Y, String(signedPercentage(effect.scaleY, 'reflection scaleY')))]),
      ],
    });
  }
  if (effect.kind !== 'softEdge') {
    throw new Error(`Unsupported composed effect kind: ${String(effect.kind)}`);
  }
  return elem(NAME_SOFT_EDGE, {
    attrs: [
      attr(ATTR_RAD, String(emuExtent(effect.radiusEmu, 'setShapeEffects: softEdge radiusEmu'))),
    ],
  });
};

/** Replace the current effect list with one schema-ordered composed list. */
export const setEffects = (host: XmlElement, effects: readonly ShapeEffectOptions[]): void => {
  removeEffectLst(host);
  if (effects.length === 0) return;
  const order: Readonly<Record<ShapeEffectOptions['kind'], number>> = {
    blur: 0,
    glow: 1,
    innerShdw: 2,
    outerShdw: 3,
    reflection: 4,
    softEdge: 5,
  };
  const kinds = new Set<string>();
  for (const effect of effects) {
    if (kinds.has(effect.kind))
      throw new Error(`setShapeEffects received duplicate ${effect.kind} effects.`);
    kinds.add(effect.kind);
  }
  const children = [...effects]
    .sort((left, right) => order[left.kind] - order[right.kind])
    .map(buildComposedEffect);
  host.children.splice(effectInsertionIndex(host), 0, elem(NAME_EFFECT_LST, { children }));
};

/**
 * Sets an outer shadow on `host`'s effect list. Replaces any prior
 * `<a:effectLst>` entirely (we treat shadow + glow as mutually
 * exclusive in v1 — multi-effect stacks are a post-1.0 enhancement).
 */
export const setShadow = (host: XmlElement, options: ShadowOptions = {}): void => {
  removeEffectLst(host);
  const color = options.color ?? '#000000';
  // blurRad and dist are ST_PositiveCoordinate (EMU, 0..27273042316900); a
  // fractional/negative/non-finite/over-max value would emit a schema-invalid
  // `<a:outerShdw>`. Validate at this boundary like every other EMU input.
  const blur = emuExtent(options.blurEmu ?? 50800, 'setShapeShadow: blurEmu');
  const dist = emuExtent(options.offsetEmu ?? 38100, 'setShapeShadow: offsetEmu');
  const angleDeg = options.angleDeg ?? 45;
  const dir = String(Math.round((((angleDeg % 360) + 360) % 360) * 60000));

  const outerShdw = elem(NAME_OUTER_SHDW, {
    attrs: [
      attr(ATTR_BLUR_RAD, String(blur)),
      attr(ATTR_DIST, String(dist)),
      attr(ATTR_DIR, dir),
      attr(ATTR_ALGN, 'tl'),
      attr(ATTR_ROT_WITH_SHAPE, '0'),
    ],
    children: [colorWithAlpha(color, options.opacity)],
  });
  const effectLst = elem(NAME_EFFECT_LST, { children: [outerShdw] });
  host.children.splice(effectInsertionIndex(host), 0, effectLst);
};

/**
 * Sets a glow on `host`'s effect list. Replaces any prior
 * `<a:effectLst>`.
 */
export const setGlow = (host: XmlElement, options: GlowOptions): void => {
  removeEffectLst(host);
  // rad is ST_PositiveCoordinate — validate like the shadow EMU inputs above.
  const rad = String(emuExtent(options.radiusEmu ?? 63500, 'setShapeGlow: radiusEmu'));
  const glow = elem(NAME_GLOW, {
    attrs: [attr(ATTR_RAD, rad)],
    children: [buildColorElement(options.color)],
  });
  const effectLst = elem(NAME_EFFECT_LST, { children: [glow] });
  host.children.splice(effectInsertionIndex(host), 0, effectLst);
};

/** Removes any effect list from `host`. */
export const clearEffects = (host: XmlElement): void => {
  removeEffectLst(host);
};
