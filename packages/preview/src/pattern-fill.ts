// Deterministic SVG approximations for every ECMA-376 ST_PresetPatternVal.
// Shape, glyph, and decoration renderers share this module so the same token
// never changes visual meaning between DrawingML surfaces.

import type { PatternPreset } from '@office-kit/pptx';

const ORTHO_TILE = 4;
const ORTHO_WIDE_TILE = 16;
const DIAG_TILE = 16;
const MOTIF_TILE = 8;

const PATTERN_TILE_SIZE: Readonly<Partial<Record<PatternPreset, number>>> = {
  horz: 6,
  ltHorz: 8,
  narHorz: 3,
  dashHorz: ORTHO_TILE,
  horzBrick: ORTHO_TILE,
  dkHorz: ORTHO_TILE,
  vert: 6,
  ltVert: 8,
  narVert: 3,
  dashVert: ORTHO_TILE,
  dkVert: ORTHO_TILE,
  cross: ORTHO_TILE,
  dotGrid: ORTHO_TILE,
  smGrid: ORTHO_TILE,
  lgGrid: ORTHO_WIDE_TILE,
  plaid: ORTHO_WIDE_TILE,
  lgCheck: ORTHO_WIDE_TILE,
  lgConfetti: ORTHO_WIDE_TILE,
  wave: MOTIF_TILE,
  zigZag: MOTIF_TILE,
  weave: MOTIF_TILE,
  divot: MOTIF_TILE,
  sphere: MOTIF_TILE,
  solidDmnd: MOTIF_TILE,
  openDmnd: MOTIF_TILE,
  dotDmnd: MOTIF_TILE,
  wdDnDiag: 24,
  wdUpDiag: 24,
  trellis: MOTIF_TILE,
};

export interface SvgPresetPatternInput {
  readonly id: string;
  readonly preset: PatternPreset;
  readonly foreground: string;
  readonly background: string;
}

const number = (value: number): string => String(Number(value.toFixed(3)));

/** Build one standalone SVG pattern definition for a legal DrawingML preset token. */
export const svgPresetPatternDefinition = (pattern: SvgPresetPatternInput): string => {
  const { id, preset, foreground: fg, background: bg } = pattern;
  const densityMatch = /^pct(\d+)$/u.exec(preset);
  const density = densityMatch
    ? Math.min(100, Math.max(0, Number.parseInt(densityMatch[1]!, 10))) / 100
    : null;
  const densityFloor = 0.05;
  const width =
    density === null
      ? (PATTERN_TILE_SIZE[preset] ?? DIAG_TILE)
      : Math.max(6, DIAG_TILE * 2 * Math.sqrt(densityFloor / Math.max(density, densityFloor)));
  const height = width;
  const stripe = (orientation: 'h' | 'v' | 'd' | 'a', strokeWidth = 0.8, dash = ''): string => {
    const path =
      orientation === 'h'
        ? `M0 ${number(height / 2)}H${number(width)}`
        : orientation === 'v'
          ? `M${number(width / 2)} 0V${number(height)}`
          : orientation === 'd'
            ? `M0 0L${number(width)} ${number(height)}`
            : `M${number(width)} 0L0 ${number(height)}`;
    return `<path d="${path}" stroke="${fg}" stroke-width="${number(strokeWidth)}"${dash ? ` stroke-dasharray="${dash}"` : ''}/>`;
  };

  let body: string;
  if (density !== null) {
    body = density >= 0.3 ? stripe('d') + stripe('a') : stripe('d');
  } else if (preset === 'horz' || preset === 'ltHorz' || preset === 'narHorz') {
    body = stripe('h', preset === 'ltHorz' ? 0.5 : 0.8);
  } else if (preset === 'dkHorz') {
    body = stripe('h', 2);
  } else if (preset === 'dashHorz') {
    body = stripe('h', 0.8, '2 2');
  } else if (preset === 'vert' || preset === 'ltVert' || preset === 'narVert') {
    body = stripe('v', preset === 'ltVert' ? 0.5 : 0.8);
  } else if (preset === 'dkVert') {
    body = stripe('v', 2);
  } else if (preset === 'dashVert') {
    body = stripe('v', 0.8, '2 2');
  } else if (preset === 'upDiag' || preset === 'ltUpDiag' || preset === 'wdUpDiag') {
    body = stripe('d', preset === 'ltUpDiag' ? 0.5 : 0.8);
  } else if (preset === 'dkUpDiag') {
    body = stripe('d', 2);
  } else if (preset === 'dashUpDiag') {
    body = stripe('d', 0.8, '3 2');
  } else if (preset === 'dnDiag' || preset === 'ltDnDiag' || preset === 'wdDnDiag') {
    body = stripe('a', preset === 'ltDnDiag' ? 0.5 : 0.8);
  } else if (preset === 'dkDnDiag') {
    body = stripe('a', 2);
  } else if (preset === 'dashDnDiag') {
    body = stripe('a', 0.8, '3 2');
  } else if (preset === 'cross' || preset === 'smGrid') {
    body = stripe('h') + stripe('v');
  } else if (preset === 'lgGrid') {
    body = stripe('h', 1.2) + stripe('v', 1.2);
  } else if (preset === 'dotGrid') {
    body = `<circle cx="${number(width / 2)}" cy="${number(height / 2)}" r="0.8" fill="${fg}"/>`;
  } else if (preset === 'diagCross' || preset === 'trellis') {
    body = stripe('d') + stripe('a');
  } else if (preset === 'smCheck' || preset === 'lgCheck') {
    const half = number(width / 2);
    body = `<path d="M0 0H${half}V${half}H0ZM${half} ${half}H${number(width)}V${number(height)}H${half}Z" fill="${fg}"/>`;
  } else if (preset === 'smConfetti' || preset === 'lgConfetti') {
    const size = preset === 'smConfetti' ? 1 : 2;
    body = `<rect x="1" y="1" width="${size}" height="${size}" fill="${fg}"/><rect x="${number(width / 2 + 1)}" y="${number(height / 2 + 1)}" width="${size}" height="${size}" fill="${fg}" transform="rotate(35 ${number(width / 2 + 1)} ${number(height / 2 + 1)})"/>`;
  } else if (preset === 'horzBrick') {
    body = `<path d="M0 2H${number(width)}M0 6H${number(width)}M2 0V2M${number(width - 2)} 2V6" stroke="${fg}" stroke-width="0.8" fill="none"/>`;
  } else if (preset === 'diagBrick') {
    body = `${stripe('d')}<path d="M0 ${number(height / 2)}L${number(width / 2)} ${number(height)}M${number(width / 2)} 0L${number(width)} ${number(height / 2)}" stroke="${fg}" stroke-width="0.8"/>`;
  } else if (preset === 'solidDmnd' || preset === 'openDmnd' || preset === 'dotDmnd') {
    const fill = preset === 'solidDmnd' ? fg : 'none';
    const dash = preset === 'dotDmnd' ? ' stroke-dasharray="1 1"' : '';
    body = `<path d="M4 1L7 4 4 7 1 4Z" fill="${fill}" stroke="${fg}" stroke-width="0.6"${dash}/>`;
  } else if (preset === 'plaid') {
    body =
      stripe('h', 2) +
      stripe('v', 1) +
      `<path d="M0 4H${number(width)}M4 0V${number(height)}" stroke="${fg}" stroke-width="0.5" opacity="0.65"/>`;
  } else if (preset === 'sphere') {
    body = `<circle cx="4" cy="4" r="3" fill="${fg}" fill-opacity="0.7"/>`;
  } else if (preset === 'weave') {
    body = `<path d="M0 0L4 4 0 8M4 0L8 4 4 8" stroke="${fg}" stroke-width="0.8" fill="none"/>`;
  } else if (preset === 'divot') {
    body = `<path d="M1 2L2 1 3 2M5 6L6 5 7 6" stroke="${fg}" stroke-width="0.8" fill="none"/>`;
  } else if (preset === 'shingle') {
    body = `<path d="M0 4Q2 0 4 4T8 4M0 8Q2 4 4 8T8 8" stroke="${fg}" stroke-width="0.8" fill="none"/>`;
  } else if (preset === 'wave') {
    body = `<path d="M0 4Q2 2 4 4T8 4" stroke="${fg}" stroke-width="0.8" fill="none"/>`;
  } else if (preset === 'zigZag') {
    body = `<path d="M0 5L2 3 4 5 6 3 8 5" stroke="${fg}" stroke-width="0.8" fill="none"/>`;
  } else {
    body = stripe('d') + stripe('a');
  }

  return `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${number(width)}" height="${number(height)}"><rect width="${number(width)}" height="${number(height)}" fill="${bg}"/>${body}</pattern>`;
};
