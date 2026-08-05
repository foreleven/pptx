// Shared color parsing for the authoring APIs.
//
// Accepts opaque or CSS-order alpha sRGB hex (`#RRGGBB`, `RRGGBB`,
// `#RRGGBBAA`, `RRGGBBAA`) and scheme tokens (`tx1`, `accent1`... — bare or
// `scheme:`-prefixed). Anything else throws so callers don't silently emit
// `<a:srgbClr val="undefined"/>`.

import { type XmlElement, NS, attr, elem, qname } from '../xml/index.ts';

const NAME_SRGB_CLR = qname('a', 'srgbClr', NS.dml);
const NAME_SCHEME_CLR = qname('a', 'schemeClr', NS.dml);
const NAME_ALPHA = qname('a', 'alpha', NS.dml);
const ATTR_VAL = qname('', 'val', '');

const SCHEME_TOKENS = new Set([
  'bg1',
  'tx1',
  'bg2',
  'tx2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
  'phClr',
  'lt1',
  'dk1',
  'lt2',
  'dk2',
]);

export type ParsedColor =
  | { kind: 'srgb'; hex: string; alpha?: number }
  | { kind: 'scheme'; token: string };

/**
 * Normalizes an sRGB hex string to the canonical uppercase 6-digit form
 * (no `#`), or returns `null` if it isn't a 3- or 6-digit hex. The CSS-style
 * 3-digit shorthand (`#f00` → `FF0000`) is accepted because LLM authors reach
 * for it constantly. Alpha-bearing colors are parsed separately because
 * OOXML stores alpha as a child transform rather than inside `val`.
 */
const normalizeSrgbHex = (value: string): string | null => {
  const hex = value.startsWith('#') ? value.slice(1) : value;
  if (/^[0-9A-Fa-f]{6}$/.test(hex)) return hex.toUpperCase();
  if (/^[0-9A-Fa-f]{3}$/.test(hex)) {
    return Array.from(hex, (ch) => ch + ch)
      .join('')
      .toUpperCase();
  }
  return null;
};

/** Parse CSS-order `RRGGBBAA`, converting the alpha byte to OOXML's 0..100000 percentage. */
const parseSrgbColor = (value: string): Extract<ParsedColor, { kind: 'srgb' }> | null => {
  const hex = value.startsWith('#') ? value.slice(1) : value;
  if (/^[0-9A-Fa-f]{8}$/.test(hex)) {
    return {
      kind: 'srgb',
      hex: hex.slice(0, 6).toUpperCase(),
      alpha: Math.round((Number.parseInt(hex.slice(6), 16) / 255) * 100000),
    };
  }
  const opaque = normalizeSrgbHex(value);
  return opaque === null ? null : { kind: 'srgb', hex: opaque };
};

/**
 * Parses a user-supplied color string. Returns null on unrecognized input
 * so callers can decide whether to throw with a specific message.
 *
 * Scheme tokens are accepted both bare (`accent1`) and with the explicit
 * `scheme:` prefix (`scheme:accent1`). The prefixed form is what the getters
 * (`getShapeFillColor`, `getSlideBackground`, …) return, so accepting it here
 * is what makes `setX(getX(...))` round-trip instead of throwing.
 */
export const parseColor = (value: string): ParsedColor | null => {
  const token = value.startsWith('scheme:') ? value.slice('scheme:'.length) : value;
  if (SCHEME_TOKENS.has(token)) return { kind: 'scheme', token };
  // A `scheme:`-prefixed value is unambiguously a scheme reference; an unknown
  // token there is an error, not a hex fallthrough.
  if (value !== token) return null;
  return parseSrgbColor(value);
};

/**
 * Parses an opaque sRGB hex color (`#RRGGBB`, `RRGGBB`, or the 3-digit `#RGB`
 * shorthand), returning the normalized uppercase 6-digit hex (no `#`).
 * Returns `null` for anything else — including alpha-bearing colors and scheme
 * tokens, which opaque sRGB-only contexts (themes and charts) must reject.
 */
export const parseSrgbHex = (value: string): string | null => normalizeSrgbHex(value);

/**
 * Returns the `<a:srgbClr>` or `<a:schemeClr>` element for `value`.
 * Throws on unrecognized colors.
 */
export const buildColorElement = (value: string): XmlElement => {
  const parsed = parseColor(value);
  if (parsed === null) throw new Error(`unrecognized color: ${value}`);
  if (parsed.kind === 'scheme') {
    return elem(NAME_SCHEME_CLR, { attrs: [attr(ATTR_VAL, parsed.token)] });
  }
  return elem(NAME_SRGB_CLR, {
    attrs: [attr(ATTR_VAL, parsed.hex)],
    children:
      parsed.alpha === undefined
        ? []
        : [elem(NAME_ALPHA, { attrs: [attr(ATTR_VAL, String(parsed.alpha))] })],
  });
};
