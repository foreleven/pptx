export interface SvgEffectBounds {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export type SvgEffect =
  | { readonly kind: 'blur'; readonly radiusPx: number; readonly grow?: boolean }
  | {
      readonly kind: 'fillOverlay';
      readonly color: string;
      readonly opacity?: number;
      readonly blend: 'over' | 'mult' | 'screen' | 'darken' | 'lighten';
    }
  | {
      readonly kind: 'glow';
      readonly color: string;
      readonly radiusPx: number;
      readonly opacity?: number;
    }
  | {
      readonly kind: 'innerShadow' | 'outerShadow';
      readonly color: string;
      readonly blurPx: number;
      readonly distancePx: number;
      readonly angleDeg: number;
      readonly opacity?: number;
    }
  | {
      readonly kind: 'presetShadow';
      readonly color: string;
      readonly blurPx: number;
      readonly distancePx: number;
      readonly angleDeg: number;
      readonly opacity?: number;
    }
  | {
      readonly kind: 'reflection';
      readonly blurPx: number;
      readonly distancePx: number;
      readonly angleDeg: number;
      readonly startOpacity?: number;
      readonly endOpacity?: number;
      readonly scaleY?: number;
    }
  | { readonly kind: 'softEdge'; readonly radiusPx: number };

export interface SvgEffectPlan {
  readonly defs: string;
  readonly filterId?: string;
  readonly reflect: (source: string) => string;
}

export interface SvgEffectPlanOptions {
  readonly id: string;
  readonly bounds: SvgEffectBounds;
  readonly reflectionMaskMode?: 'userSpaceOnUse' | 'objectBoundingBox';
  readonly reflectionDataAttribute?: {
    readonly name: `data-${string}`;
    readonly value: string;
  };
}

/**
 * Build one deterministic SVG filter/reflection plan from normalized px-native effects.
 * Source-changing effects form a chain while glow/shadow layers remain composited, so
 * a later fill overlay, blur, or soft edge never silently discards an earlier effect.
 */
export const buildSvgEffectPlan = (
  effects: readonly SvgEffect[],
  options: SvgEffectPlanOptions,
): SvgEffectPlan => {
  const filterId = options.id;
  if (!/^[A-Za-z_][\w:.-]*$/u.test(filterId)) {
    throw new TypeError(`Invalid SVG effect id: ${filterId}`);
  }
  const primitives: string[] = [];
  const backgroundLayers: string[] = [];
  const foregroundLayers: string[] = [];
  let sourceResult = 'SourceGraphic';
  let reflection: Extract<SvgEffect, { kind: 'reflection' }> | undefined;

  for (const effect of effects) {
    const suffix = String(primitives.length);
    if (effect.kind === 'fillOverlay') {
      const mode =
        effect.blend === 'mult' ? 'multiply' : effect.blend === 'over' ? 'normal' : effect.blend;
      primitives.push(
        `<feFlood flood-color="${escapeXml(effect.color)}" flood-opacity="${number(effect.opacity ?? 1)}" result="overlayFlood${suffix}"/>`,
        `<feComposite in="overlayFlood${suffix}" in2="SourceAlpha" operator="in" result="overlayClip${suffix}"/>`,
        `<feBlend in="${sourceResult}" in2="overlayClip${suffix}" mode="${mode}" result="overlayOut${suffix}"/>`,
      );
      sourceResult = `overlayOut${suffix}`;
    } else if (effect.kind === 'outerShadow' || effect.kind === 'presetShadow') {
      const offset = effectOffset(effect.distancePx, effect.angleDeg);
      primitives.push(
        `<feDropShadow dx="${number(offset.x)}" dy="${number(offset.y)}" stdDeviation="${number(effect.blurPx / 2)}" flood-color="${escapeXml(effect.color)}" flood-opacity="${number(effect.opacity ?? 1)}" result="shadowOut${suffix}"/>`,
      );
      backgroundLayers.push(`shadowOut${suffix}`);
    } else if (effect.kind === 'innerShadow') {
      const offset = effectOffset(effect.distancePx, effect.angleDeg);
      primitives.push(
        `<feGaussianBlur in="SourceAlpha" stdDeviation="${number(effect.blurPx / 2)}" result="innerBlur${suffix}"/>`,
        `<feOffset in="innerBlur${suffix}" dx="${number(offset.x)}" dy="${number(offset.y)}" result="innerOffset${suffix}"/>`,
        `<feComposite in="innerOffset${suffix}" in2="SourceAlpha" operator="arithmetic" k2="-1" k3="1" result="innerMask${suffix}"/>`,
        `<feFlood flood-color="${escapeXml(effect.color)}" flood-opacity="${number(effect.opacity ?? 1)}" result="innerColor${suffix}"/>`,
        `<feComposite in="innerColor${suffix}" in2="innerMask${suffix}" operator="in" result="innerOut${suffix}"/>`,
      );
      foregroundLayers.push(`innerOut${suffix}`);
    } else if (effect.kind === 'glow') {
      primitives.push(
        `<feMorphology in="SourceAlpha" operator="dilate" radius="${number(effect.radiusPx * 0.5)}" result="glowExpand${suffix}"/>`,
        `<feGaussianBlur in="glowExpand${suffix}" stdDeviation="${number(effect.radiusPx * 0.3)}" result="glowBlur${suffix}"/>`,
        `<feFlood flood-color="${escapeXml(effect.color)}" flood-opacity="${number(effect.opacity ?? 1)}" result="glowColor${suffix}"/>`,
        `<feComposite in="glowColor${suffix}" in2="glowBlur${suffix}" operator="in" result="glowOut${suffix}"/>`,
      );
      backgroundLayers.push(`glowOut${suffix}`);
    } else if (effect.kind === 'softEdge' || effect.kind === 'blur') {
      const result = `${effect.kind === 'softEdge' ? 'soft' : 'blur'}Out${suffix}`;
      primitives.push(
        `<feGaussianBlur in="${sourceResult}" stdDeviation="${number(effect.radiusPx / 2)}" result="${result}"/>`,
      );
      sourceResult = result;
    } else if (effect.kind === 'reflection') {
      reflection = effect;
    }
  }

  if (primitives.length > 0) {
    primitives.push(
      `<feMerge>${backgroundLayers.map((layer) => `<feMergeNode in="${layer}"/>`).join('')}<feMergeNode in="${sourceResult}"/>${foregroundLayers.map((layer) => `<feMergeNode in="${layer}"/>`).join('')}</feMerge>`,
    );
  }

  const reflectionId = `${filterId}-reflection`;
  const defs = [
    primitives.length > 0
      ? `<filter id="${escapeXml(filterId)}" x="-50%" y="-50%" width="200%" height="250%" color-interpolation-filters="sRGB">${primitives.join('')}</filter>`
      : '',
    reflection
      ? reflectionDefs(
          reflection,
          reflectionId,
          options.bounds,
          options.reflectionMaskMode ?? 'userSpaceOnUse',
        )
      : '',
  ].join('');
  return {
    defs,
    ...(primitives.length > 0 ? { filterId } : {}),
    reflect: reflection
      ? (source) =>
          renderReflection(
            source,
            reflection,
            reflectionId,
            options.bounds,
            options.reflectionDataAttribute,
          )
      : () => '',
  };
};

const reflectionDefs = (
  effect: Extract<SvgEffect, { kind: 'reflection' }>,
  id: string,
  bounds: SvgEffectBounds,
  mode: NonNullable<SvgEffectPlanOptions['reflectionMaskMode']>,
): string => {
  const start = number(effect.startOpacity ?? 1);
  const end = number(effect.endOpacity ?? 0);
  const top = bounds.y + bounds.height + effect.distancePx;
  const bottom = top + bounds.height * Math.abs(effect.scaleY ?? -1);
  const blur =
    effect.blurPx > 0
      ? `<filter id="${escapeXml(`${id}-blur`)}" x="-30%" y="-30%" width="160%" height="160%"><feGaussianBlur stdDeviation="${number(effect.blurPx / 2)}"/></filter>`
      : '';
  const maskX = bounds.x - bounds.width;
  const maskWidth = bounds.width * 3;
  const maskHeight = Math.max(bounds.height, bottom - top);
  if (mode === 'objectBoundingBox') {
    return (
      `<linearGradient id="${escapeXml(`${id}-gradient`)}" x1="0" y1="0" x2="0" y2="1">` +
      `<stop offset="0" stop-color="#fff" stop-opacity="${end}"/><stop offset="1" stop-color="#fff" stop-opacity="${start}"/></linearGradient>` +
      `<mask id="${escapeXml(`${id}-mask`)}" maskContentUnits="objectBoundingBox">` +
      `<rect width="1" height="1" fill="url(#${escapeXml(`${id}-gradient`)})"/></mask>${blur}`
    );
  }
  return (
    `<linearGradient id="${escapeXml(`${id}-gradient`)}" gradientUnits="userSpaceOnUse" x1="0" y1="${number(top)}" x2="0" y2="${number(bottom)}">` +
    `<stop offset="0" stop-color="#fff" stop-opacity="${start}"/><stop offset="1" stop-color="#fff" stop-opacity="${end}"/></linearGradient>` +
    `<mask id="${escapeXml(`${id}-mask`)}" maskUnits="userSpaceOnUse" x="${number(maskX)}" y="${number(top)}" width="${number(maskWidth)}" height="${number(maskHeight)}">` +
    `<rect x="${number(maskX)}" y="${number(top)}" width="${number(maskWidth)}" height="${number(maskHeight)}" fill="url(#${escapeXml(`${id}-gradient`)})"/></mask>${blur}`
  );
};

const renderReflection = (
  source: string,
  effect: Extract<SvgEffect, { kind: 'reflection' }>,
  id: string,
  bounds: SvgEffectBounds,
  dataAttribute: SvgEffectPlanOptions['reflectionDataAttribute'],
): string => {
  const scale = Math.abs(effect.scaleY ?? -1);
  const translate =
    bounds.y + bounds.height + effect.distancePx + (bounds.y + bounds.height) * scale;
  const filter = effect.blurPx > 0 ? ` filter="url(#${escapeXml(`${id}-blur`)})"` : '';
  if (dataAttribute && !/^data-[A-Za-z_:][\w:.-]*$/u.test(dataAttribute.name)) {
    throw new TypeError(`Invalid SVG data attribute name: ${dataAttribute.name}`);
  }
  const marker = dataAttribute ? ` ${dataAttribute.name}="${escapeXml(dataAttribute.value)}"` : '';
  return `<g transform="translate(0 ${number(translate)}) scale(1 -${number(scale)})" mask="url(#${escapeXml(`${id}-mask`)})"${filter}${marker}>${source}</g>`;
};

const effectOffset = (distancePx: number, angleDeg: number): { x: number; y: number } => {
  const radians = (angleDeg * Math.PI) / 180;
  return { x: distancePx * Math.cos(radians), y: distancePx * Math.sin(radians) };
};

const number = (value: number): string => Number(value.toFixed(3)).toString();

const escapeXml = (value: string): string =>
  value
    .replace(/&/gu, '&amp;')
    .replace(/</gu, '&lt;')
    .replace(/>/gu, '&gt;')
    .replace(/"/gu, '&quot;');
