// Color transforms and rPr-like element parsing.

import { NAME_A_RPR, requireRun } from './shape-runs.ts';
import {
  isLineAlignment,
  isLineCap,
  isLineCompound,
  isLineDash,
  isLineJoin,
  parseEffectList,
  type Effect,
  type TextFormat,
  type TextLineProperties,
  type TextRunState,
} from '../../internal/drawingml/index.ts';
import { PATTERN_PRESETS } from '../../internal/drawingml/fill.ts';
import {
  NS,
  type XmlElement,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import { type SlideShapeData } from '../_internal-symbols.ts';
import { lineWidthEmu } from '../../internal/bounds.ts';
import { type PresentationTheme } from './theme.ts';

/**
 * Direct, non-inherited DrawingML run state that does not describe visual CSS formatting.
 * These values are exposed so importers can diagnose proofing and smart-tag metadata
 * instead of silently flattening it into ordinary text.
 */
const optionalOnOffAttr = (element: XmlElement, localName: string): boolean | undefined => {
  const value = getAttrValue(element, qname('', localName, ''));
  if (value === '1' || value === 'true') return true;
  if (value === '0' || value === 'false') return false;
  return undefined;
};

/** Name every attribute outside one element's exact unqualified OOXML vocabulary. */
const unexpectedAttributes = (
  element: XmlElement,
  allowed: ReadonlySet<string>,
  label: string,
): string[] =>
  element.attrs
    .filter(
      (attribute) => attribute.name.namespaceURI !== '' || !allowed.has(attribute.name.localName),
    )
    .map((attribute) => `${label} attribute ${attribute.name.localName}`);

const RUN_FILL_CHOICES = new Set([
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
]);

/** Parse the structural subset shared by text `<a:ln>` and `<a:uLn>`. */
const parseTextLineProperties = (
  line: XmlElement,
  label: string,
  options: { readonly rejectZeroWidth: boolean },
): { readonly properties: TextLineProperties; readonly unsupported: string[] } => {
  const properties: {
    -readonly [Key in keyof TextLineProperties]?: TextLineProperties[Key];
  } = {};
  const unsupported = unexpectedAttributes(line, new Set(['w', 'cap', 'cmpd', 'algn']), label);

  const widthRaw = getAttrValue(line, qname('', 'w', ''));
  if (widthRaw !== null) {
    if (!/^\d+$/u.test(widthRaw)) {
      unsupported.push(`width ${widthRaw}`);
    } else {
      try {
        const widthEmu = lineWidthEmu(Number(widthRaw), `${label}: width`);
        if (options.rejectZeroWidth && widthEmu === 0) unsupported.push('width 0');
        else properties.widthPt = widthEmu / 12_700;
      } catch {
        unsupported.push(`width ${widthRaw}`);
      }
    }
  }

  const cap = getAttrValue(line, qname('', 'cap', ''));
  if (cap !== null) {
    if (isLineCap(cap)) properties.cap = cap;
    else unsupported.push(`cap ${cap}`);
  }
  const compound = getAttrValue(line, qname('', 'cmpd', ''));
  if (compound !== null) {
    if (isLineCompound(compound)) properties.compound = compound;
    else unsupported.push(`compound ${compound}`);
  }
  const alignment = getAttrValue(line, qname('', 'algn', ''));
  if (alignment !== null) {
    if (isLineAlignment(alignment)) properties.alignment = alignment;
    else unsupported.push(`alignment ${alignment}`);
  }

  const presetDashes = line.children.filter(
    (child): child is XmlElement =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.dml &&
      child.name.localName === 'prstDash',
  );
  const customDashes = line.children.filter(
    (child): child is XmlElement =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.dml &&
      child.name.localName === 'custDash',
  );
  if (presetDashes.length + customDashes.length > 1) unsupported.push('multiple line dash choices');
  const presetDash = presetDashes[0];
  if (presetDash) {
    unsupported.push(...unexpectedAttributes(presetDash, new Set(['val']), `${label} dash`));
    if (presetDash.children.some((child) => child.kind === 'element')) {
      unsupported.push(`${label} dash children`);
    }
    const dash = getAttrValue(presetDash, qname('', 'val', ''));
    if (dash !== null && isLineDash(dash)) {
      properties.dash = dash;
    } else unsupported.push(`${label} dash ${dash ?? 'missing'}`);
  }
  const customDash = customDashes[0];
  if (customDash) {
    unsupported.push(...unexpectedAttributes(customDash, new Set(), `${label} custom dash`));
    const stops: Array<{ dash: number; space: number }> = [];
    const percentage = (raw: string): number | null => {
      if (/^\d+$/u.test(raw)) return Number(raw);
      const match = /^(\d+(?:\.\d+)?)%$/u.exec(raw);
      return match ? Number(match[1]) * 1_000 : null;
    };
    for (const child of customDash.children) {
      if (
        child.kind !== 'element' ||
        child.name.namespaceURI !== NS.dml ||
        child.name.localName !== 'ds'
      ) {
        unsupported.push(
          `${label} custom dash child ${child.kind === 'element' ? child.name.localName : child.kind}`,
        );
        continue;
      }
      unsupported.push(
        ...unexpectedAttributes(child, new Set(['d', 'sp']), `${label} custom dash stop`),
      );
      const dash = percentage(getAttrValue(child, qname('', 'd', '')) ?? '');
      const space = percentage(getAttrValue(child, qname('', 'sp', '')) ?? '');
      if (dash === null || space === null || dash < 0 || space < 0) {
        unsupported.push(`${label} invalid custom dash stop`);
      } else stops.push({ dash, space });
    }
    if (stops.length === 0) unsupported.push(`${label} empty custom dash`);
    else properties.customDash = stops;
  }

  const joins = line.children.filter(
    (child): child is XmlElement =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.dml &&
      (child.name.localName === 'round' ||
        child.name.localName === 'bevel' ||
        child.name.localName === 'miter'),
  );
  if (joins.length > 1) unsupported.push('multiple line join choices');
  const join = joins[0];
  if (join) {
    if (isLineJoin(join.name.localName)) properties.join = join.name.localName;
    unsupported.push(...unexpectedAttributes(join, new Set(), `${label} join`));
    if (join.children.some((child) => child.kind === 'element')) {
      unsupported.push(`${label} join children`);
    }
  }

  const knownChildren = new Set([
    'noFill',
    'solidFill',
    'gradFill',
    'pattFill',
    'prstDash',
    'custDash',
    'round',
    'bevel',
    'miter',
    'headEnd',
    'tailEnd',
    'extLst',
  ]);
  for (const child of line.children) {
    if (child.kind !== 'element') continue;
    if (child.name.namespaceURI !== NS.dml) {
      unsupported.push(`foreign child ${child.name.localName}`);
    } else if (child.name.localName === 'headEnd' || child.name.localName === 'tailEnd') {
      const type = getAttrValue(child, qname('', 'type', '')) ?? 'none';
      const width = getAttrValue(child, qname('', 'w', ''));
      const length = getAttrValue(child, qname('', 'len', ''));
      const allowedTypes = new Set(['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow']);
      const allowedSizes = new Set(['sm', 'med', 'lg']);
      unsupported.push(
        ...unexpectedAttributes(child, new Set(['type', 'w', 'len']), `${label} arrow`),
      );
      if (
        !allowedTypes.has(type) ||
        (width !== null && !allowedSizes.has(width)) ||
        (length !== null && !allowedSizes.has(length))
      ) {
        unsupported.push(`${label} invalid arrow ${child.name.localName}`);
      } else {
        const arrow = {
          type: type as 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow',
          ...(width === null ? {} : { width: width as 'sm' | 'med' | 'lg' }),
          ...(length === null ? {} : { length: length as 'sm' | 'med' | 'lg' }),
        };
        if (child.name.localName === 'headEnd') properties.head = arrow;
        else properties.tail = arrow;
      }
    } else if (child.name.localName === 'extLst') {
      // Opaque line extensions are preserved by the scoped text-extension
      // carrier; they do not invalidate the known editable base line.
    } else if (!knownChildren.has(child.name.localName)) {
      unsupported.push(`child ${child.name.localName}`);
    }
  }
  return { properties, unsupported };
};

type ParsedTextLinePaint =
  | { readonly kind: 'bare' | 'none'; readonly unsupported: string[] }
  | { readonly kind: 'solid'; readonly color: string; readonly unsupported: string[] }
  | ({ readonly kind: 'gradient'; readonly unsupported: string[] } & {
      readonly stops: readonly { readonly offset: number; readonly color: string }[];
      readonly angleDeg: number;
    })
  | ({ readonly kind: 'pattern'; readonly unsupported: string[] } & {
      readonly preset: (typeof PATTERN_PRESETS)[number];
      readonly foreground: string;
      readonly background: string;
    })
  | { readonly kind: 'unsupported'; readonly reason: string };

/** Parse the strict editable linear-gradient subset shared by text fills and line paints. */
const parseLinearTextGradientElement = (
  gradient: XmlElement,
  label: string,
  ctx?: { readonly theme: PresentationTheme | null },
): {
  readonly stops: readonly { readonly offset: number; readonly color: string }[];
  readonly angleDeg: number;
  readonly unsupported: readonly string[];
} | null => {
  const gsList = firstChildElement(gradient, qname('a', 'gsLst', NS.dml));
  const linear = firstChildElement(gradient, qname('a', 'lin', NS.dml));
  if (!gsList || !linear) return null;
  const stops: Array<{ offset: number; color: string }> = [];
  const unsupported: string[] = [];
  let previous = -1;
  for (const stop of gsList.children) {
    if (
      stop.kind !== 'element' ||
      stop.name.namespaceURI !== NS.dml ||
      stop.name.localName !== 'gs'
    ) {
      return null;
    }
    const rawPosition = getAttrValue(stop, qname('', 'pos', '')) ?? '';
    const position = Number(rawPosition);
    const color = stop.children.find((child): child is XmlElement => child.kind === 'element');
    if (
      !/^\d+$/u.test(rawPosition) ||
      !Number.isSafeInteger(position) ||
      position < previous ||
      position > 100_000 ||
      color?.name.namespaceURI !== NS.dml
    ) {
      return null;
    }
    const transforms = color.children.filter(
      (child): child is XmlElement => child.kind === 'element',
    );
    let colorValue: string | null = null;
    if (ctx) {
      colorValue = resolveDrawingColor(color, ctx.theme);
      if (colorValue === null) return null;
      if (
        color.name.localName !== 'srgbClr' ||
        transforms.some(
          (transform) =>
            transform.name.namespaceURI !== NS.dml || transform.name.localName !== 'alpha',
        )
      ) {
        unsupported.push(
          `${label} scheme/transformed color resolved to fixed sRGB; future theme responsiveness lost`,
        );
      }
    } else {
      const rawColor = getAttrValue(color, qname('', 'val', ''));
      if (
        color.name.localName !== 'srgbClr' ||
        rawColor === null ||
        !/^[\dA-Fa-f]{6}$/u.test(rawColor) ||
        transforms.length > 1 ||
        (transforms[0] &&
          (transforms[0].name.namespaceURI !== NS.dml || transforms[0].name.localName !== 'alpha'))
      ) {
        return null;
      }
      colorValue = `#${rawColor.toUpperCase()}`;
    }
    const alphaTransform = transforms.find(
      (transform) => transform.name.namespaceURI === NS.dml && transform.name.localName === 'alpha',
    );
    if (alphaTransform) {
      const rawAlpha = getAttrValue(alphaTransform, qname('', 'val', '')) ?? '';
      const alpha = Number(rawAlpha);
      if (!/^\d+$/u.test(rawAlpha) || alpha < 0 || alpha > 100_000) return null;
      const alphaByte = Math.round((alpha / 100_000) * 255);
      if (Math.round((alphaByte / 255) * 100_000) !== alpha) {
        unsupported.push(`${label} alpha ${rawAlpha} rounded to the nearest CSS alpha byte`);
      }
      if (!ctx) colorValue += alphaByte.toString(16).padStart(2, '0').toUpperCase();
    }
    stops.push({ offset: position / 100_000, color: colorValue });
    previous = position;
  }
  const rawAngle = getAttrValue(linear, qname('', 'ang', '')) ?? '';
  const angle = Number(rawAngle);
  if (
    stops.length < 2 ||
    !/^\d+$/u.test(rawAngle) ||
    !Number.isSafeInteger(angle) ||
    angle >= 21_600_000
  ) {
    return null;
  }
  void label;
  return { stops, angleDeg: angle / 60_000, unsupported };
};

/** Parse the no/solid paint subset of one CT_LineProperties element. */
const parseTextLinePaint = (
  line: XmlElement,
  label: string,
  ctx?: { readonly theme: PresentationTheme | null },
): ParsedTextLinePaint => {
  const fills = line.children.filter(
    (child): child is XmlElement =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.dml &&
      ['noFill', 'solidFill', 'gradFill', 'pattFill'].includes(child.name.localName),
  );
  if (fills.length === 0) return { kind: 'bare', unsupported: [] };
  if (fills.length > 1) return { kind: 'unsupported', reason: `multiple ${label} fill choices` };
  const fill = fills[0]!;
  const allowedFillAttributes =
    fill.name.localName === 'gradFill'
      ? new Set(['flip', 'rotWithShape'])
      : fill.name.localName === 'pattFill'
        ? new Set(['prst'])
        : new Set<string>();
  const unsupported = unexpectedAttributes(
    fill,
    allowedFillAttributes,
    `${label} ${fill.name.localName}`,
  );
  if (fill.name.localName === 'noFill') {
    if (fill.children.some((child) => child.kind === 'element')) {
      unsupported.push(`${label} noFill children`);
    }
    return { kind: 'none', unsupported };
  }
  if (fill.name.localName === 'gradFill') {
    const parsed = parseLinearTextGradientElement(fill, label, ctx);
    return parsed === null
      ? { kind: 'unsupported', reason: `${label} invalid gradFill` }
      : {
          kind: 'gradient',
          stops: parsed.stops,
          angleDeg: parsed.angleDeg,
          unsupported: [...unsupported, ...parsed.unsupported],
        };
  }
  if (fill.name.localName === 'pattFill') {
    const preset = getAttrValue(fill, qname('', 'prst', ''));
    if (preset === null || !(PATTERN_PRESETS as readonly string[]).includes(preset)) {
      return { kind: 'unsupported', reason: `${label} invalid pattFill preset` };
    }
    const colorFrom = (localName: 'fgClr' | 'bgClr'): string | null => {
      const container = firstChildElement(fill, qname('a', localName, NS.dml));
      const color = container?.children.find(
        (child): child is XmlElement => child.kind === 'element',
      );
      if (!color) return null;
      if (ctx) {
        const resolved = resolveDrawingColor(color, ctx.theme);
        if (
          resolved !== null &&
          (color.name.localName !== 'srgbClr' ||
            color.children.some((child) => child.kind === 'element'))
        ) {
          unsupported.push(
            `${label} pattern scheme/transformed color resolved to fixed sRGB; future theme responsiveness lost`,
          );
        }
        return resolved;
      }
      if (color.name.namespaceURI !== NS.dml || color.name.localName !== 'srgbClr') return null;
      const raw = getAttrValue(color, qname('', 'val', ''));
      return raw !== null && /^[\dA-Fa-f]{6}$/u.test(raw) ? `#${raw.toUpperCase()}` : null;
    };
    const foreground = colorFrom('fgClr');
    const background = colorFrom('bgClr');
    if (foreground === null || background === null) {
      return { kind: 'unsupported', reason: `${label} pattern colors` };
    }
    return {
      kind: 'pattern',
      preset: preset as (typeof PATTERN_PRESETS)[number],
      foreground,
      background,
      unsupported,
    };
  }
  if (fill.name.localName !== 'solidFill') {
    return { kind: 'unsupported', reason: `${label} ${fill.name.localName}` };
  }
  const colors = fill.children.filter((child): child is XmlElement => child.kind === 'element');
  const color = colors[0];
  if (colors.length !== 1 || color?.name.namespaceURI !== NS.dml) {
    return { kind: 'unsupported', reason: `${label} color ${color?.name.localName ?? 'missing'}` };
  }
  unsupported.push(...unexpectedAttributes(color, new Set(['val']), `${label} color`));
  const hasColorTransforms = color.children.some((child) => child.kind === 'element');
  let colorValue: string | null = null;
  if (ctx) {
    colorValue = resolveDrawingColor(color, ctx.theme);
    if (colorValue !== null && (color.name.localName !== 'srgbClr' || hasColorTransforms)) {
      unsupported.push(
        `${label} scheme/transformed color resolved to fixed sRGB; future theme responsiveness lost`,
      );
    }
  } else if (color.name.localName === 'srgbClr') {
    const value = getAttrValue(color, qname('', 'val', ''));
    if (value !== null && /^[\dA-Fa-f]{6}$/u.test(value)) colorValue = `#${value.toUpperCase()}`;
  } else if (color.name.localName === 'schemeClr') {
    colorValue = getAttrValue(color, qname('', 'val', ''));
  }
  if (!ctx && color.name.localName !== 'srgbClr')
    unsupported.push(`${label} color ${color.name.localName}`);
  if (colorValue === null) {
    return { kind: 'unsupported', reason: unsupported.join(', ') || `${label} invalid solid fill` };
  }
  return { kind: 'solid', color: colorValue, unsupported };
};

const parseTextGradient = (rPr: XmlElement): NonNullable<TextFormat['gradient']> | undefined => {
  const fills = rPr.children.filter(
    (child): child is XmlElement =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.dml &&
      RUN_FILL_CHOICES.has(child.name.localName),
  );
  const gradient = fills.find((fill) => fill.name.localName === 'gradFill');
  if (!gradient) return undefined;

  const unsupported: string[] = [];
  if (fills.length !== 1) unsupported.push('multiple text fill choices');
  unsupported.push(
    ...unexpectedAttributes(gradient, new Set(['flip', 'rotWithShape']), 'text gradient'),
  );
  const flip = getAttrValue(gradient, qname('', 'flip', ''));
  if (flip !== null && flip !== 'none') unsupported.push(`text gradient flip ${flip}`);
  const rotate = getAttrValue(gradient, qname('', 'rotWithShape', ''));
  if (rotate !== null && rotate !== '1' && rotate !== 'true') {
    unsupported.push(`text gradient rotWithShape ${rotate}`);
  }

  const children = gradient.children.filter(
    (child): child is XmlElement => child.kind === 'element',
  );
  const gsLists = children.filter(
    (child) => child.name.namespaceURI === NS.dml && child.name.localName === 'gsLst',
  );
  const linear = children.filter(
    (child) => child.name.namespaceURI === NS.dml && child.name.localName === 'lin',
  );
  if (gsLists.length !== 1) unsupported.push(`text gradient gsLst count ${gsLists.length}`);
  if (linear.length !== 1) unsupported.push(`text gradient lin count ${linear.length}`);
  for (const child of children) {
    if (
      child.name.namespaceURI !== NS.dml ||
      (child.name.localName !== 'gsLst' && child.name.localName !== 'lin')
    ) {
      unsupported.push(`text gradient child ${child.name.localName}`);
    }
  }

  const stops: Array<{ offset: number; color: string }> = [];
  const gsList = gsLists[0];
  if (gsList) {
    unsupported.push(...unexpectedAttributes(gsList, new Set(), 'text gradient stop list'));
    const stopElements = gsList.children.filter(
      (child): child is XmlElement => child.kind === 'element',
    );
    if (stopElements.length < 2)
      unsupported.push(`text gradient stop count ${stopElements.length}`);
    let previousPosition = -1;
    for (const stop of stopElements) {
      if (stop.name.namespaceURI !== NS.dml || stop.name.localName !== 'gs') {
        unsupported.push(`text gradient stop child ${stop.name.localName}`);
        continue;
      }
      unsupported.push(...unexpectedAttributes(stop, new Set(['pos']), 'text gradient stop'));
      const positionRaw = getAttrValue(stop, qname('', 'pos', '')) ?? '';
      const position = Number(positionRaw);
      if (
        !/^\d+$/u.test(positionRaw) ||
        !Number.isSafeInteger(position) ||
        position < 0 ||
        position > 100_000
      ) {
        unsupported.push(`text gradient stop position ${positionRaw || 'missing'}`);
      } else if (position < previousPosition) {
        unsupported.push('text gradient stop positions decrease');
      }
      previousPosition = position;

      const colors = stop.children.filter((child): child is XmlElement => child.kind === 'element');
      const color = colors[0];
      if (
        colors.length !== 1 ||
        color?.name.namespaceURI !== NS.dml ||
        color.name.localName !== 'srgbClr'
      ) {
        unsupported.push(`text gradient stop color ${color?.name.localName ?? 'missing'}`);
        continue;
      }
      unsupported.push(
        ...unexpectedAttributes(color, new Set(['val']), 'text gradient stop color'),
      );
      const colorRaw = getAttrValue(color, qname('', 'val', '')) ?? '';
      if (!/^[\dA-Fa-f]{6}$/u.test(colorRaw)) {
        unsupported.push(`text gradient stop color ${colorRaw || 'missing'}`);
        continue;
      }
      const transforms = color.children.filter(
        (child): child is XmlElement => child.kind === 'element',
      );
      let alphaHex = '';
      if (
        transforms.length > 1 ||
        (transforms[0] &&
          (transforms[0].name.namespaceURI !== NS.dml || transforms[0].name.localName !== 'alpha'))
      ) {
        unsupported.push('text gradient stop color transforms');
      } else if (transforms[0]) {
        const alpha = transforms[0];
        unsupported.push(
          ...unexpectedAttributes(alpha, new Set(['val']), 'text gradient stop alpha'),
        );
        if (alpha.children.some((child) => child.kind === 'element')) {
          unsupported.push('text gradient stop alpha children');
        }
        const alphaRaw = getAttrValue(alpha, qname('', 'val', '')) ?? '';
        const alphaValue = Number(alphaRaw);
        if (
          !/^\d+$/u.test(alphaRaw) ||
          !Number.isInteger(alphaValue) ||
          alphaValue < 0 ||
          alphaValue > 100_000
        ) {
          unsupported.push(`text gradient stop alpha ${alphaRaw || 'missing'}`);
        } else {
          const alphaByte = Math.round((alphaValue / 100_000) * 255);
          const cssAlphaValue = Math.round((alphaByte / 255) * 100_000);
          if (alphaValue !== cssAlphaValue) {
            unsupported.push(`text gradient stop alpha ${alphaRaw} exceeds CSS hex-byte precision`);
          } else {
            alphaHex = alphaByte.toString(16).padStart(2, '0').toUpperCase();
          }
        }
      }
      if (/^\d+$/u.test(positionRaw) && position >= 0 && position <= 100_000) {
        stops.push({ offset: position / 100_000, color: `#${colorRaw.toUpperCase()}${alphaHex}` });
      }
    }
  }

  let angleDeg = 0;
  const lin = linear[0];
  if (lin) {
    unsupported.push(...unexpectedAttributes(lin, new Set(['ang', 'scaled']), 'text gradient lin'));
    if (lin.children.some((child) => child.kind === 'element')) {
      unsupported.push('text gradient lin children');
    }
    const scaled = getAttrValue(lin, qname('', 'scaled', ''));
    if (scaled !== null && scaled !== '0' && scaled !== 'false') {
      unsupported.push(`text gradient scaled ${scaled}`);
    }
    const angleRaw = getAttrValue(lin, qname('', 'ang', '')) ?? '';
    const angle = Number(angleRaw);
    if (
      !/^\d+$/u.test(angleRaw) ||
      !Number.isSafeInteger(angle) ||
      angle < 0 ||
      angle >= 21_600_000
    ) {
      unsupported.push(`text gradient angle ${angleRaw || 'missing'}`);
    } else {
      angleDeg = angle / 60_000;
    }
  }

  if (unsupported.length > 0) return { unsupported: unsupported.join(', ') };
  return { stops, angleDeg };
};

const parseTextPatternFill = (
  rPr: XmlElement,
  ctx?: { readonly theme: PresentationTheme | null },
): NonNullable<TextFormat['patternFill']> | undefined => {
  const fills = rPr.children.filter(
    (child): child is XmlElement =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.dml &&
      RUN_FILL_CHOICES.has(child.name.localName),
  );
  if (!fills.some((fill) => fill.name.localName === 'pattFill')) return undefined;
  const parsed = parseTextLinePaint(rPr, 'text fill', ctx);
  if (parsed.kind !== 'pattern') {
    return {
      unsupported:
        parsed.kind === 'unsupported'
          ? parsed.reason
          : `text fill ${parsed.kind} instead of pattern`,
    };
  }
  const unsupported = [
    ...(fills.length === 1 ? [] : ['multiple text fill choices']),
    ...parsed.unsupported,
  ];
  return {
    preset: parsed.preset,
    foreground: parsed.foreground,
    background: parsed.background,
    ...(unsupported.length === 0 ? {} : { unsupported: unsupported.join(', ') }),
  };
};

const parseTextPictureFill = (
  rPr: XmlElement,
): NonNullable<TextFormat['pictureFill']> | undefined => {
  const fills = rPr.children.filter(
    (child): child is XmlElement =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.dml &&
      RUN_FILL_CHOICES.has(child.name.localName),
  );
  const picture = fills.find((fill) => fill.name.localName === 'blipFill');
  if (!picture) return undefined;
  const unsupported = [
    ...(fills.length === 1 ? [] : ['multiple text fill choices']),
    ...unexpectedAttributes(picture, new Set(), 'text picture fill'),
  ];
  const children = picture.children.filter(
    (child): child is XmlElement => child.kind === 'element',
  );
  const blips = children.filter(
    (child) => child.name.namespaceURI === NS.dml && child.name.localName === 'blip',
  );
  const stretches = children.filter(
    (child) => child.name.namespaceURI === NS.dml && child.name.localName === 'stretch',
  );
  if (blips.length !== 1) unsupported.push(`text picture fill blip count ${blips.length}`);
  if (stretches.length !== 1)
    unsupported.push(`text picture fill stretch count ${stretches.length}`);
  for (const child of children) {
    if (
      child.name.namespaceURI !== NS.dml ||
      (child.name.localName !== 'blip' && child.name.localName !== 'stretch')
    ) {
      unsupported.push(`text picture fill child ${child.name.localName}`);
    }
  }
  const blip = blips[0];
  const relationshipId = blip ? getAttrValue(blip, qname('r', 'embed', NS.officeDocRels)) : null;
  if (blip) {
    const link = getAttrValue(blip, qname('r', 'link', NS.officeDocRels));
    if (link !== null) unsupported.push('text picture fill external link');
    if (blip.children.some((child) => child.kind === 'element')) {
      unsupported.push('text picture fill blip effects');
    }
  }
  const stretch = stretches[0];
  if (stretch) {
    unsupported.push(...unexpectedAttributes(stretch, new Set(), 'text picture fill stretch'));
    const stretchChildren = stretch.children.filter(
      (child): child is XmlElement => child.kind === 'element',
    );
    if (
      stretchChildren.length !== 1 ||
      stretchChildren[0]?.name.namespaceURI !== NS.dml ||
      stretchChildren[0].name.localName !== 'fillRect'
    ) {
      unsupported.push('text picture fill non-canonical stretch');
    } else {
      unsupported.push(
        ...unexpectedAttributes(stretchChildren[0], new Set(), 'text picture fill fillRect'),
      );
    }
  }
  return relationshipId === null
    ? { unsupported: unsupported.join(', ') || 'text picture fill missing embedded relationship' }
    : {
        relationshipId,
        ...(unsupported.length === 0 ? {} : { unsupported: unsupported.join(', ') }),
      };
};

const parseTextShadow = (rPr: XmlElement): NonNullable<TextFormat['textShadow']> | undefined => {
  if (firstChildElement(rPr, qname('a', 'effectDag', NS.dml))) {
    return { unsupported: 'text effect DAG' };
  }
  const effectLst = firstChildElement(rPr, qname('a', 'effectLst', NS.dml));
  if (!effectLst) return undefined;

  const unsupported = unexpectedAttributes(effectLst, new Set(), 'text effect list');
  const effects = effectLst.children.filter(
    (child): child is XmlElement => child.kind === 'element',
  );
  if (
    effects.length !== 1 ||
    effects[0]?.name.namespaceURI !== NS.dml ||
    effects[0].name.localName !== 'outerShdw'
  ) {
    const names = effects.map((effect) => effect.name.localName).join(', ') || 'empty';
    unsupported.push(`text effects ${names}`);
    return { unsupported: unsupported.join(', ') };
  }

  const shadow = effects[0];
  unsupported.push(
    ...unexpectedAttributes(
      shadow,
      new Set(['blurRad', 'dist', 'dir', 'algn', 'rotWithShape']),
      'outer text shadow',
    ),
  );
  const alignment = getAttrValue(shadow, qname('', 'algn', ''));
  if (alignment !== null && alignment !== 'tl')
    unsupported.push(`outer text shadow alignment ${alignment}`);
  const rotate = getAttrValue(shadow, qname('', 'rotWithShape', ''));
  if (rotate !== null && rotate !== '0' && rotate !== 'false') {
    unsupported.push(`outer text shadow rotWithShape ${rotate}`);
  }

  const nonNegativeInteger = (name: string): number => {
    const raw = getAttrValue(shadow, qname('', name, '')) ?? '0';
    if (!/^\d+$/u.test(raw)) {
      unsupported.push(`outer text shadow ${name} ${raw}`);
      return 0;
    }
    const value = Number(raw);
    if (!Number.isSafeInteger(value)) {
      unsupported.push(`outer text shadow ${name} ${raw}`);
      return 0;
    }
    return value;
  };
  const blurEmu = nonNegativeInteger('blurRad');
  const distEmu = nonNegativeInteger('dist');
  const directionRaw = getAttrValue(shadow, qname('', 'dir', '')) ?? '0';
  if (!/^-?\d+$/u.test(directionRaw) || !Number.isSafeInteger(Number(directionRaw))) {
    unsupported.push(`outer text shadow dir ${directionRaw}`);
  }
  const directionDeg = Number(directionRaw) / 60_000;

  const colorChildren = shadow.children.filter(
    (child): child is XmlElement => child.kind === 'element',
  );
  const color = colorChildren[0];
  if (
    colorChildren.length !== 1 ||
    color?.name.namespaceURI !== NS.dml ||
    color.name.localName !== 'srgbClr'
  ) {
    unsupported.push(`outer text shadow color ${color?.name.localName ?? 'missing'}`);
  }
  let colorValue = '';
  let opacity: number | undefined;
  if (color?.name.namespaceURI === NS.dml && color.name.localName === 'srgbClr') {
    unsupported.push(...unexpectedAttributes(color, new Set(['val']), 'outer text shadow color'));
    const rawColor = getAttrValue(color, qname('', 'val', '')) ?? '';
    if (!/^[\dA-Fa-f]{6}$/u.test(rawColor))
      unsupported.push(`outer text shadow color ${rawColor || 'missing'}`);
    else colorValue = `#${rawColor.toUpperCase()}`;
    const transforms = color.children.filter(
      (child): child is XmlElement => child.kind === 'element',
    );
    if (transforms.length > 1 || (transforms[0] && transforms[0].name.localName !== 'alpha')) {
      unsupported.push('outer text shadow color transforms');
    } else if (transforms[0]) {
      const alpha = transforms[0];
      unsupported.push(...unexpectedAttributes(alpha, new Set(['val']), 'outer text shadow alpha'));
      if (alpha.children.some((child) => child.kind === 'element')) {
        unsupported.push('outer text shadow alpha children');
      }
      const rawAlpha = getAttrValue(alpha, qname('', 'val', '')) ?? '';
      const alphaValue = Number(rawAlpha);
      if (
        !/^\d+$/u.test(rawAlpha) ||
        !Number.isInteger(alphaValue) ||
        alphaValue < 0 ||
        alphaValue > 100_000
      ) {
        unsupported.push(`outer text shadow alpha ${rawAlpha || 'missing'}`);
      } else {
        const alphaByte = Math.round((alphaValue / 100_000) * 255);
        const cssAlphaValue = Math.round((alphaByte / 255) * 100_000);
        if (alphaValue !== cssAlphaValue) {
          unsupported.push(`outer text shadow alpha ${rawAlpha} exceeds CSS hex-byte precision`);
        } else {
          opacity = alphaValue / 100_000;
        }
      }
    }
  }

  if (unsupported.length > 0) return { unsupported: unsupported.join(', ') };
  const radians = (directionDeg * Math.PI) / 180;
  const roundPoint = (value: number): number => Number(value.toFixed(4));
  return {
    color: colorValue,
    ...(opacity === undefined ? {} : { opacity }),
    blurPt: roundPoint(blurEmu / 12_700),
    offsetXPt: roundPoint((distEmu * Math.cos(radians)) / 12_700),
    offsetYPt: roundPoint((distEmu * Math.sin(radians)) / 12_700),
  };
};

const parseTextEffects = (
  rPr: XmlElement,
  theme: PresentationTheme | null,
): readonly Effect[] | undefined => {
  const effectLst = firstChildElement(rPr, qname('a', 'effectLst', NS.dml));
  if (!effectLst) return undefined;
  return parseEffectList(effectLst, (color) => resolveDrawingColor(color, theme));
};
// -- Color transforms (ECMA-376 §20.1.2.3.x) --------------------------------
//
// DrawingML color elements (`<a:srgbClr>`, `<a:schemeClr>`, `<a:sysClr>`,
// `<a:prstClr>`) may carry one or more transform children — `lumMod`,
// `lumOff`, `shade`, `tint`, `satMod`, `hueMod`, `alpha`, `gray`, `inv`,
// `comp`, etc. — that adjust the base color before it's painted. Real
// templates use them heavily for "tinted accent" backgrounds and "shaded
// hover" states, so any visual-fidelity story has to apply them.
//
// Percentages in the spec use the `ST_Percentage` style — `100000`
// represents 100% — though some third-party tools emit bare floats; we
// accept both forms.

type ColorTransformOp =
  | {
      readonly kind:
        | 'lumMod'
        | 'lumOff'
        | 'shade'
        | 'tint'
        | 'satMod'
        | 'satOff'
        | 'hueMod'
        | 'hueOff'
        | 'alpha'
        | 'alphaMod'
        | 'alphaOff';
      readonly val: number;
    }
  | { readonly kind: 'gray' | 'inv' | 'comp' };

const COLOR_TRANSFORM_LOCALS: ReadonlySet<string> = new Set([
  'lumMod',
  'lumOff',
  'shade',
  'tint',
  'satMod',
  'satOff',
  'hueMod',
  'hueOff',
  'alpha',
  'alphaMod',
  'alphaOff',
  'gray',
  'inv',
  'comp',
]);

const parseColorTransforms = (colorEl: XmlElement): readonly ColorTransformOp[] => {
  const out: ColorTransformOp[] = [];
  for (const child of colorEl.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    const local = child.name.localName;
    if (!COLOR_TRANSFORM_LOCALS.has(local)) continue;
    if (local === 'gray' || local === 'inv' || local === 'comp') {
      out.push({ kind: local });
      continue;
    }
    const raw = getAttrValue(child, qname('', 'val', ''));
    if (raw === null) continue;
    let n = Number.parseFloat(raw);
    if (!Number.isFinite(n)) continue;
    // PowerPoint emits ST_Percentage (`100000` = 100%); tolerate the
    // bare-float form some third-party tools emit.
    if (Math.abs(n) > 1) n = n / 100000;
    out.push({ kind: local as Exclude<ColorTransformOp['kind'], 'gray' | 'inv' | 'comp'>, val: n });
  }
  return out;
};

const hexToRgb01 = (hex: string): [number, number, number] => {
  const h = hex.startsWith('#') ? hex.slice(1) : hex;
  return [
    Number.parseInt(h.slice(0, 2), 16) / 255,
    Number.parseInt(h.slice(2, 4), 16) / 255,
    Number.parseInt(h.slice(4, 6), 16) / 255,
  ];
};

const rgb01ToHex = (r: number, g: number, b: number): string => {
  const clamp = (v: number): number => Math.max(0, Math.min(255, Math.round(v * 255)));
  const part = (n: number): string => n.toString(16).padStart(2, '0').toUpperCase();
  return `#${part(clamp(r))}${part(clamp(g))}${part(clamp(b))}`;
};

const rgbToHsl = (r: number, g: number, b: number): [number, number, number] => {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
};

const hueToRgb = (p: number, q: number, t: number): number => {
  if (t < 0) t += 1;
  if (t > 1) t -= 1;
  if (t < 1 / 6) return p + (q - p) * 6 * t;
  if (t < 1 / 2) return q;
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
  return p;
};

const hslToRgb = (h: number, s: number, l: number): [number, number, number] => {
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hueToRgb(p, q, h + 1 / 3), hueToRgb(p, q, h), hueToRgb(p, q, h - 1 / 3)];
};

// PowerPoint applies <a:tint> / <a:shade> in LINEAR-LIGHT RGB, not in sRGB —
// this contradicts the literal ECMA-376 "N% of input + (100-N)% white/black"
// wording, but it is what PowerPoint computes and what LibreOffice renders
// (a 75% tint of black is mid-gray ~#8B8B8B, not the sRGB-lerp #404040). The
// next reader will expect the sRGB formula, hence this note.
const srgbToLinear = (c: number): number =>
  c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
const linearToSrgb = (c: number): number =>
  c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055;

const applyColorTransforms = (hex: string, transforms: readonly ColorTransformOp[]): string => {
  if (transforms.length === 0) return hex;
  let [r, g, b] = hexToRgb01(hex);
  for (const t of transforms) {
    switch (t.kind) {
      case 'inv':
        r = 1 - r;
        g = 1 - g;
        b = 1 - b;
        break;
      case 'gray': {
        const y = 0.3 * r + 0.59 * g + 0.11 * b;
        r = g = b = y;
        break;
      }
      case 'comp': {
        const [h, s, l] = rgbToHsl(r, g, b);
        [r, g, b] = hslToRgb((h + 0.5) % 1, s, l);
        break;
      }
      case 'shade':
        // Mix toward black in linear light: out = srgb(linear(base) * val)
        r = linearToSrgb(srgbToLinear(r) * t.val);
        g = linearToSrgb(srgbToLinear(g) * t.val);
        b = linearToSrgb(srgbToLinear(b) * t.val);
        break;
      case 'tint':
        // Mix toward white in linear light: out = srgb(linear(base)*val + (1-val))
        r = linearToSrgb(srgbToLinear(r) * t.val + (1 - t.val));
        g = linearToSrgb(srgbToLinear(g) * t.val + (1 - t.val));
        b = linearToSrgb(srgbToLinear(b) * t.val + (1 - t.val));
        break;
      case 'lumMod':
      case 'lumOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const newL = Math.max(0, Math.min(1, t.kind === 'lumMod' ? l * t.val : l + t.val));
        [r, g, b] = hslToRgb(h, s, newL);
        break;
      }
      case 'satMod':
      case 'satOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const newS = Math.max(0, Math.min(1, t.kind === 'satMod' ? s * t.val : s + t.val));
        [r, g, b] = hslToRgb(h, newS, l);
        break;
      }
      case 'hueMod':
      case 'hueOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const newH = (((t.kind === 'hueMod' ? h * t.val : h + t.val / 360) % 1) + 1) % 1;
        [r, g, b] = hslToRgb(newH, s, l);
        break;
      }
      // alpha / alphaMod / alphaOff intentionally don't touch RGB — they
      // surface as `fill-opacity`, not as a tinted color.
    }
  }
  return rgb01ToHex(r, g, b);
};

/** Preserve DrawingML alpha transforms as the CSS-order AA suffix used by SVG renderers. */
const applyAlphaTransforms = (hex: string, transforms: readonly ColorTransformOp[]): string => {
  let alpha = 1;
  let authored = false;
  for (const transform of transforms) {
    if (transform.kind === 'alpha') {
      alpha = transform.val;
      authored = true;
    } else if (transform.kind === 'alphaMod') {
      alpha *= transform.val;
      authored = true;
    } else if (transform.kind === 'alphaOff') {
      alpha += transform.val;
      authored = true;
    }
  }
  if (!authored) return hex;
  const byte = Math.round(Math.max(0, Math.min(1, alpha)) * 255);
  return `${hex}${byte.toString(16).padStart(2, '0').toUpperCase()}`;
};

const SCHEME_TOKEN_TO_THEME_KEY: Record<string, keyof Omit<PresentationTheme, 'name'>> = {
  tx1: 'dark1',
  dk1: 'dark1',
  bg1: 'light1',
  lt1: 'light1',
  tx2: 'dark2',
  dk2: 'dark2',
  bg2: 'light2',
  lt2: 'light2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hyperlink',
  folHlink: 'followedHyperlink',
};

/**
 * Resolves a scheme token (`tx1`, `bg1`, `accent1`, …) to its `#RRGGBB`.
 *
 * When `clrMap` is supplied, the slide token is first remapped through it
 * (`<p:clrMap>` / `<a:overrideClrMapping>`) — `tx1` may point at `dk1` or
 * `lt1` depending on the deck — and only then indexed into the theme. Without
 * a map the token is indexed directly, preserving the historical behavior
 * (correct for the standard map, the overwhelming common case).
 *
 * @internal
 */
export const resolveSchemeToken = (
  token: string,
  theme: PresentationTheme | null,
  clrMap?: Readonly<Record<string, string>> | null,
): string | null => {
  if (!theme) return null;
  const mapped = clrMap?.[token] ?? token;
  const key = SCHEME_TOKEN_TO_THEME_KEY[mapped] ?? SCHEME_TOKEN_TO_THEME_KEY[token];
  if (!key) return null;
  const hex = theme[key];
  if (typeof hex !== 'string') return null;
  const normalized = hex.startsWith('#') ? hex : `#${hex}`;
  return /^#[0-9A-Fa-f]{6}$/.test(normalized) ? normalized.toUpperCase() : null;
};

/**
 * Resolves a DrawingML color element (`<a:srgbClr>` / `<a:schemeClr>` /
 * `<a:sysClr>` / `<a:prstClr>`) with all its `<a:lumMod>` / `<a:tint>` /
 * `<a:shade>` / `<a:satMod>` etc. transform children applied. Alpha is
 * returned as a CSS-order `AA` suffix. Returns `null` when the color is a
 * scheme token and no theme is supplied to resolve it.
 *
 * Exposed because both run-format and fill-format code paths need to
 * apply the same transform pipeline; keeping a single implementation
 * means future spec-coverage additions only have to land in one place.
 */
export const resolveDrawingColor = (
  colorEl: XmlElement,
  theme: PresentationTheme | null,
  clrMap?: Readonly<Record<string, string>> | null,
): string | null => {
  if (colorEl.name.namespaceURI !== NS.dml) return null;
  const local = colorEl.name.localName;
  let baseHex: string | null = null;
  if (local === 'srgbClr') {
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v) baseHex = `#${v.toUpperCase()}`;
  } else if (local === 'schemeClr') {
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v) baseHex = resolveSchemeToken(v, theme, clrMap);
  } else if (local === 'sysClr') {
    const last = getAttrValue(colorEl, qname('', 'lastClr', ''));
    if (last) baseHex = `#${last.toUpperCase()}`;
  } else if (local === 'prstClr') {
    // Preset colors aren't worth a full lookup table in this pass —
    // black / white cover most cases anyone reaches for in PresentationML.
    const v = getAttrValue(colorEl, qname('', 'val', ''));
    if (v === 'black') baseHex = '#000000';
    else if (v === 'white') baseHex = '#FFFFFF';
  }
  if (!baseHex) return null;
  const transforms = parseColorTransforms(colorEl);
  return applyAlphaTransforms(applyColorTransforms(baseHex, transforms), transforms);
};

// Reads any element shaped like `CT_TextCharacterProperties` (the schema
// shared by `<a:rPr>`, `<a:defRPr>`, and `<a:endParaRPr>`) into a partial
// TextFormat. Used by both the literal-only `getShapeRunFormat` and the
// inheritance-aware `getShapeRunFormatEffective`.
//
// When `ctx.theme` is provided, scheme tokens are resolved to concrete
// `#RRGGBB[AA]` and color transforms (`<a:lumMod>` etc.) are applied. Without
// a theme, transforms are not applied and theme tokens are passed through
// verbatim — this preserves the legacy `getShapeRunFormat` behavior.
export const parseRPrLikeElement = (
  rPr: XmlElement,
  ctx?: { readonly theme: PresentationTheme | null },
): Partial<TextFormat> => {
  const out: Partial<TextFormat> = {};
  const sz = getAttrValue(rPr, qname('', 'sz', ''));
  if (sz !== null) {
    const n = Number.parseInt(sz, 10);
    if (Number.isFinite(n)) out.size = n / 100;
  }
  const b = getAttrValue(rPr, qname('', 'b', ''));
  if (b !== null) out.bold = b !== '0';
  const i = getAttrValue(rPr, qname('', 'i', ''));
  if (i !== null) out.italic = i !== '0';
  const u = getAttrValue(rPr, qname('', 'u', ''));
  if (u !== null) {
    if (u === 'none') out.underline = false;
    else if (u === 'sng') out.underline = true;
    else out.underline = u;
  }
  const underlineLines = rPr.children.filter(
    (child): child is XmlElement =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.dml &&
      (child.name.localName === 'uLnTx' || child.name.localName === 'uLn'),
  );
  if (underlineLines.length > 0) {
    const line = underlineLines[0]!;
    const choiceUnsupported =
      underlineLines.length === 1 ? [] : ['multiple underline line choices'];
    if (line.name.localName === 'uLnTx') {
      const unsupported = [
        ...choiceUnsupported,
        ...unexpectedAttributes(line, new Set(), 'underline line-follow-text'),
        ...(line.children.some((child) => child.kind === 'element')
          ? ['underline line-follow-text children']
          : []),
      ];
      out.underlineLine = {
        kind: 'followText',
        ...(unsupported.length === 0 ? {} : { unsupported: unsupported.join(', ') }),
      };
    } else {
      const parsed = parseTextLineProperties(line, 'underline line', {
        rejectZeroWidth: false,
      });
      const paint = parseTextLinePaint(line, 'underline line', ctx);
      if (paint.kind === 'unsupported') {
        out.underlineLine = {
          kind: 'unsupported',
          reason: [...choiceUnsupported, ...parsed.unsupported, paint.reason].join(', '),
        };
      } else {
        const unsupported = [...choiceUnsupported, ...parsed.unsupported, ...paint.unsupported];
        const diagnostic = unsupported.length === 0 ? {} : { unsupported: unsupported.join(', ') };
        out.underlineLine =
          paint.kind === 'solid'
            ? { kind: 'solid', color: paint.color, ...parsed.properties, ...diagnostic }
            : paint.kind === 'gradient' || paint.kind === 'pattern'
              ? {
                  kind: 'unsupported',
                  reason: `underline line ${paint.kind} paint is not authorable`,
                }
              : { kind: paint.kind, ...parsed.properties, ...diagnostic };
      }
    }
  }
  const underlineFills = rPr.children.filter(
    (child): child is XmlElement =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.dml &&
      (child.name.localName === 'uFillTx' || child.name.localName === 'uFill'),
  );
  if (underlineFills.length > 0) {
    const fillChoice = underlineFills[0]!;
    const unsupported = [
      ...(underlineFills.length === 1 ? [] : ['multiple underline fill choices']),
      ...unexpectedAttributes(fillChoice, new Set(), 'underline fill'),
    ];
    if (fillChoice.name.localName === 'uFillTx') {
      if (fillChoice.children.some((child) => child.kind === 'element')) {
        unsupported.push('underline fill-follow-text children');
      }
      out.underlineFill = {
        kind: 'followText',
        ...(unsupported.length === 0 ? {} : { unsupported: unsupported.join(', ') }),
      };
    } else {
      const fills = fillChoice.children.filter(
        (child): child is XmlElement => child.kind === 'element',
      );
      const fill = fills[0];
      if (fills.length !== 1 || fill?.name.namespaceURI !== NS.dml) {
        out.underlineFill = {
          kind: 'unsupported',
          reason: 'underline fill without one DrawingML fill choice',
        };
      } else if (fill.name.localName === 'noFill') {
        unsupported.push(...unexpectedAttributes(fill, new Set(), 'underline noFill'));
        if (fill.children.some((child) => child.kind === 'element')) {
          unsupported.push('underline noFill children');
        }
        out.underlineFill = {
          kind: 'none',
          ...(unsupported.length === 0 ? {} : { unsupported: unsupported.join(', ') }),
        };
      } else if (fill.name.localName === 'solidFill') {
        unsupported.push(...unexpectedAttributes(fill, new Set(), 'underline solidFill'));
        const colors = fill.children.filter(
          (child): child is XmlElement => child.kind === 'element',
        );
        const color = colors[0];
        let colorValue: string | null = null;
        const colorTransforms =
          color?.children.filter((child): child is XmlElement => child.kind === 'element') ?? [];
        const needsFixedResolution =
          color?.name.localName !== 'srgbClr' ||
          colorTransforms.some(
            (transform) =>
              transform.name.namespaceURI !== NS.dml || transform.name.localName !== 'alpha',
          );
        if (colors.length !== 1 || color?.name.namespaceURI !== NS.dml) {
          unsupported.push(`underline fill color ${color?.name.localName ?? 'missing'}`);
        } else if (ctx && needsFixedResolution) {
          colorValue = resolveDrawingColor(color, ctx.theme);
          if (colorValue === null) {
            unsupported.push(`underline fill color ${color.name.localName}`);
          } else {
            unsupported.push(
              'underline fill scheme/transformed color resolved to fixed sRGB; future theme responsiveness lost',
            );
          }
        } else if (color.name.localName !== 'srgbClr') {
          unsupported.push(`underline fill color ${color.name.localName}`);
        } else {
          unsupported.push(
            ...unexpectedAttributes(color, new Set(['val']), 'underline fill color'),
          );
          const rawColor = getAttrValue(color, qname('', 'val', '')) ?? '';
          if (!/^[\dA-Fa-f]{6}$/u.test(rawColor)) {
            unsupported.push(`underline fill color ${rawColor || 'missing'}`);
          } else {
            const transforms = colorTransforms;
            let alphaHex = '';
            if (
              transforms.length > 1 ||
              (transforms[0] &&
                (transforms[0].name.namespaceURI !== NS.dml ||
                  transforms[0].name.localName !== 'alpha'))
            ) {
              unsupported.push('underline fill color transforms');
            } else if (transforms[0]) {
              const alpha = transforms[0];
              unsupported.push(
                ...unexpectedAttributes(alpha, new Set(['val']), 'underline fill alpha'),
              );
              const rawAlpha = getAttrValue(alpha, qname('', 'val', '')) ?? '';
              const alphaValue = Number(rawAlpha);
              const alphaByte = Math.round((alphaValue / 100_000) * 255);
              if (!/^\d+$/u.test(rawAlpha) || alphaValue < 0 || alphaValue > 100_000) {
                unsupported.push(
                  `underline fill alpha ${rawAlpha || 'missing'} exceeds CSS hex-byte precision`,
                );
              } else {
                if (Math.round((alphaByte / 255) * 100_000) !== alphaValue) {
                  unsupported.push(
                    `underline fill alpha ${rawAlpha} rounded to the nearest CSS alpha byte`,
                  );
                }
                alphaHex = alphaByte.toString(16).padStart(2, '0').toUpperCase();
              }
            }
            colorValue = `#${rawColor.toUpperCase()}${alphaHex}`;
          }
        }
        out.underlineFill =
          colorValue === null
            ? {
                kind: 'unsupported',
                reason: unsupported.join(', ') || 'invalid underline solid fill',
              }
            : {
                kind: 'solid',
                color: colorValue,
                ...(unsupported.length === 0 ? {} : { unsupported: unsupported.join(', ') }),
              };
      } else if (fill.name.localName === 'gradFill') {
        const gradient = parseLinearTextGradientElement(fill, 'underline fill', ctx);
        out.underlineFill =
          gradient === null
            ? { kind: 'unsupported', reason: 'underline fill invalid gradFill' }
            : {
                kind: 'gradient',
                stops: gradient.stops,
                angleDeg: gradient.angleDeg,
                ...([...unsupported, ...gradient.unsupported].length === 0
                  ? {}
                  : { unsupported: [...unsupported, ...gradient.unsupported].join(', ') }),
              };
      } else if (fill.name.localName === 'pattFill') {
        const preset = getAttrValue(fill, qname('', 'prst', ''));
        const colorFrom = (localName: 'fgClr' | 'bgClr'): string | null => {
          const container = firstChildElement(fill, qname('a', localName, NS.dml));
          const color = container?.children.find(
            (child): child is XmlElement => child.kind === 'element',
          );
          if (!color || color.name.namespaceURI !== NS.dml) return null;
          if (ctx) {
            const resolved = resolveDrawingColor(color, ctx.theme);
            if (
              resolved !== null &&
              (color.name.localName !== 'srgbClr' ||
                color.children.some((child) => child.kind === 'element'))
            ) {
              unsupported.push(
                'underline fill pattern scheme/transformed color resolved to fixed sRGB; future theme responsiveness lost',
              );
            }
            return resolved;
          }
          if (color.name.localName !== 'srgbClr') return null;
          const raw = getAttrValue(color, qname('', 'val', ''));
          return raw !== null && /^[\dA-Fa-f]{6}$/u.test(raw) ? `#${raw.toUpperCase()}` : null;
        };
        const foreground = colorFrom('fgClr');
        const background = colorFrom('bgClr');
        out.underlineFill =
          preset !== null &&
          (PATTERN_PRESETS as readonly string[]).includes(preset) &&
          foreground !== null &&
          background !== null
            ? {
                kind: 'pattern',
                preset: preset as (typeof PATTERN_PRESETS)[number],
                foreground,
                background,
                ...(unsupported.length === 0 ? {} : { unsupported: unsupported.join(', ') }),
              }
            : { kind: 'unsupported', reason: 'underline fill invalid pattFill' };
      } else if (fill.name.localName === 'grpFill') {
        out.underlineFill = {
          kind: 'group',
          ...(unsupported.length === 0 ? {} : { unsupported: unsupported.join(', ') }),
        };
      } else if (fill.name.localName === 'blipFill') {
        const blip = firstChildElement(fill, qname('a', 'blip', NS.dml));
        const relationshipId = blip
          ? getAttrValue(blip, qname('r', 'embed', NS.officeDocRels))
          : null;
        out.underlineFill =
          relationshipId === null
            ? { kind: 'unsupported', reason: 'underline fill picture relationship missing' }
            : {
                kind: 'picture',
                relationshipId,
                ...(unsupported.length === 0 ? {} : { unsupported: unsupported.join(', ') }),
              };
      } else {
        out.underlineFill = {
          kind: 'unsupported',
          reason: `underline fill ${fill.name.localName}`,
        };
      }
    }
  }
  const strike = getAttrValue(rPr, qname('', 'strike', ''));
  if (strike !== null) {
    if (strike === 'noStrike') out.strike = false;
    else if (strike === 'sngStrike') out.strike = true;
    else out.strike = strike;
  }
  const outline = firstChildElement(rPr, qname('a', 'ln', NS.dml));
  if (outline !== null) {
    const parsed = parseTextLineProperties(outline, 'text outline', {
      rejectZeroWidth: false,
    });
    const paint = parseTextLinePaint(outline, 'text outline', ctx);
    const hasInvalidExplicitWidth =
      getAttrValue(outline, qname('', 'w', '')) !== null && parsed.properties.widthPt === undefined;
    if (paint.kind === 'unsupported' || paint.kind === 'bare' || hasInvalidExplicitWidth) {
      out.outline = {
        kind: 'unsupported',
        reason: hasInvalidExplicitWidth
          ? parsed.unsupported.join(', ')
          : paint.kind === 'unsupported'
            ? [...parsed.unsupported, paint.reason].join(', ')
            : [...parsed.unsupported, 'text outline without a supported solid/no fill'].join(', '),
      };
    } else {
      const unsupported = [...parsed.unsupported, ...paint.unsupported];
      const diagnostic = unsupported.length === 0 ? {} : { unsupported: unsupported.join(', ') };
      out.outline =
        paint.kind === 'solid'
          ? { kind: 'solid', color: paint.color, ...parsed.properties, ...diagnostic }
          : paint.kind === 'gradient'
            ? {
                kind: 'gradient',
                stops: paint.stops,
                angleDeg: paint.angleDeg,
                ...parsed.properties,
                ...diagnostic,
              }
            : paint.kind === 'pattern'
              ? {
                  kind: 'pattern',
                  preset: paint.preset,
                  foreground: paint.foreground,
                  background: paint.background,
                  ...parsed.properties,
                  ...diagnostic,
                }
              : { kind: 'none', ...parsed.properties, ...diagnostic };
    }
  }
  const spc = getAttrValue(rPr, qname('', 'spc', ''));
  if (spc !== null) {
    const n = Number.parseInt(spc, 10);
    if (Number.isFinite(n)) out.spc = n;
  }
  const kern = getAttrValue(rPr, qname('', 'kern', ''));
  if (kern !== null) {
    const n = Number.parseInt(kern, 10);
    if (Number.isFinite(n)) out.kern = n;
  }
  const baselineAttr = getAttrValue(rPr, qname('', 'baseline', ''));
  if (baselineAttr !== null) {
    // ST_Percentage: 100000 = 100%; tolerate bare floats.
    let n = Number.parseFloat(baselineAttr);
    if (Number.isFinite(n)) {
      if (Math.abs(n) > 1) n = n / 100000;
      out.baseline = n;
    }
  }
  const cap = getAttrValue(rPr, qname('', 'cap', ''));
  if (cap === 'none' || cap === 'small' || cap === 'all') {
    out.cap = cap;
  }
  // <a:highlight><a:srgbClr val="…"/></a:highlight>
  const highlight = firstChildElement(rPr, qname('a', 'highlight', NS.dml));
  if (highlight !== null) {
    let hlChild: XmlElement | null = null;
    for (const c of highlight.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      hlChild = c;
      break;
    }
    if (hlChild) {
      if (ctx) {
        const hex = resolveDrawingColor(hlChild, ctx.theme);
        if (hex !== null) out.highlight = hex;
      } else if (hlChild.name.localName === 'srgbClr') {
        const v = getAttrValue(hlChild, qname('', 'val', ''));
        if (v !== null) out.highlight = `#${v.toUpperCase()}`;
      } else if (hlChild.name.localName === 'schemeClr') {
        const v = getAttrValue(hlChild, qname('', 'val', ''));
        if (v !== null) out.highlight = v;
      }
    }
  }
  const solidFill = firstChildElement(rPr, qname('a', 'solidFill', NS.dml));
  if (solidFill !== null) {
    // Find the inner color element (srgbClr / schemeClr / sysClr / prstClr).
    // CT_SolidColorFillProperties holds exactly one EG_ColorChoice child.
    let colorChild: XmlElement | null = null;
    for (const c of solidFill.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      colorChild = c;
      break;
    }
    if (colorChild) {
      if (ctx) {
        // Apply transforms + resolve scheme tokens to hex.
        const hex = resolveDrawingColor(colorChild, ctx.theme);
        if (hex !== null) out.color = hex;
        else if (colorChild.name.localName === 'schemeClr') {
          // Theme not provided / token not in scheme — surface the raw token.
          const v = getAttrValue(colorChild, qname('', 'val', ''));
          if (v !== null) out.color = v;
        }
      } else {
        // Legacy `getShapeRunFormat` path: no transforms, scheme tokens
        // emitted as bare strings to match prior public behavior.
        if (colorChild.name.localName === 'srgbClr') {
          const v = getAttrValue(colorChild, qname('', 'val', ''));
          if (v !== null) out.color = `#${v.toUpperCase()}`;
        } else if (colorChild.name.localName === 'schemeClr') {
          const v = getAttrValue(colorChild, qname('', 'val', ''));
          if (v !== null) out.color = v;
        }
      }
    }
  }
  const effects = parseTextEffects(rPr, ctx?.theme ?? null);
  if (effects !== undefined) out.effects = effects;
  if (effects === undefined) {
    const textShadow = parseTextShadow(rPr);
    if (textShadow !== undefined) out.textShadow = textShadow;
  }
  const gradient = parseTextGradient(rPr);
  if (gradient !== undefined) out.gradient = gradient;
  const patternFill = parseTextPatternFill(rPr, ctx);
  if (patternFill !== undefined) out.patternFill = patternFill;
  const pictureFill = parseTextPictureFill(rPr);
  if (pictureFill !== undefined) out.pictureFill = pictureFill;
  const latin = firstChildElement(rPr, qname('a', 'latin', NS.dml));
  if (latin !== null) {
    const t = getAttrValue(latin, qname('', 'typeface', ''));
    if (t !== null) out.font = t;
  }
  const ea = firstChildElement(rPr, qname('a', 'ea', NS.dml));
  if (ea !== null) {
    const t = getAttrValue(ea, qname('', 'typeface', ''));
    if (t !== null) out.fontEastAsian = t;
  }
  const cs = firstChildElement(rPr, qname('a', 'cs', NS.dml));
  if (cs !== null) {
    const t = getAttrValue(cs, qname('', 'typeface', ''));
    if (t !== null) out.fontComplexScript = t;
  }
  const sym = firstChildElement(rPr, qname('a', 'sym', NS.dml));
  if (sym !== null) {
    const t = getAttrValue(sym, qname('', 'typeface', ''));
    if (t !== null) out.fontSymbol = t;
  }
  const language = getAttrValue(rPr, qname('', 'lang', ''));
  if (language !== null) out.language = language;
  const alternativeLanguage = getAttrValue(rPr, qname('', 'altLang', ''));
  if (alternativeLanguage !== null) out.alternativeLanguage = alternativeLanguage;
  return out;
};

/**
 * Reads back the format of a single run. Returns `null` when the run
 * has no `<a:rPr>` (it inherits its format from the paragraph /
 * layout / master). Boolean attributes that are explicitly `"0"`
 * decode to `false`.
 *
 * Use `getShapeRunFormatEffective` if you want the resolved format
 * after walking the placeholder / lstStyle / master inheritance chain.
 */
export const getShapeRunFormat = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): TextFormat | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, NAME_A_RPR);
  if (rPr === null) return null;
  return parseRPrLikeElement(rPr) as TextFormat;
};

/**
 * Reads direct CT_TextCharacterProperties state that is intentionally excluded
 * from the inherited visual `TextFormat` cascade. Returns `null` when the run
 * has no `<a:rPr>` or none of these state attributes is authored.
 */
/** @internal Reads direct non-visual run state from any CT_TextCharacterProperties element. */
export const parseTextRunState = (rPr: XmlElement): TextRunState | null => {
  const normalizeHeight = optionalOnOffAttr(rPr, 'normalizeH');
  const noProof = optionalOnOffAttr(rPr, 'noProof');
  const dirty = optionalOnOffAttr(rPr, 'dirty');
  const error = optionalOnOffAttr(rPr, 'err');
  const smartTagClean = optionalOnOffAttr(rPr, 'smtClean');
  const smartTagIdValue = getAttrValue(rPr, qname('', 'smtId', ''));
  const smartTagId = smartTagIdValue === null ? Number.NaN : Number(smartTagIdValue);
  const state: TextRunState = {
    ...(normalizeHeight !== undefined ? { normalizeHeight } : {}),
    ...(noProof !== undefined ? { noProof } : {}),
    ...(dirty !== undefined ? { dirty } : {}),
    ...(error !== undefined ? { error } : {}),
    ...(smartTagClean !== undefined ? { smartTagClean } : {}),
    ...(Number.isInteger(smartTagId) && smartTagId >= 0 && smartTagId <= 0xffffffff
      ? { smartTagId }
      : {}),
  };
  return Object.keys(state).length > 0 ? state : null;
};

export const getShapeRunState = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): TextRunState | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, NAME_A_RPR);
  return rPr === null ? null : parseTextRunState(rPr);
};
