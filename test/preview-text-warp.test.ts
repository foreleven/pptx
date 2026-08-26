// WordArt preview proof: native preset metadata remains editable in OOXML while
// the server's pure-SVG path provides a deterministic visual approximation.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  inches,
  setParagraphAlignment,
  setShapeRunFormat,
  setShapeTextFromWordArt,
  setShapeTextWrap,
  setShapeTextWarp,
} from '../src/api/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

interface PixelImage {
  readonly width: number;
  readonly height: number;
  readonly data: Uint8Array;
}

/** Measure non-white preview content within one non-overlapping WordArt band. */
const foregroundBounds = (
  image: PixelImage,
  y0: number,
  y1: number,
): { x: number; y: number; width: number; height: number } => {
  let minX = image.width;
  let minY = image.height;
  let maxX = -1;
  let maxY = -1;
  for (let y = y0; y < Math.min(y1, image.height); y++) {
    for (let x = 0; x < image.width; x++) {
      const offset = (y * image.width + x) * 4;
      if (
        image.data[offset]! < 245 ||
        image.data[offset + 1]! < 245 ||
        image.data[offset + 2]! < 245
      ) {
        minX = Math.min(minX, x);
        minY = Math.min(minY, y);
        maxX = Math.max(maxX, x);
        maxY = Math.max(maxY, y);
      }
    }
  }
  return { x: minX, y: minY, width: maxX - minX + 1, height: maxY - minY + 1 };
};

const addWarpedText = (
  slide: Parameters<typeof addSlideTextBox>[0],
  text: string,
  y: number,
  preset: 'textArchUp' | 'textWave1' | 'textInflate',
  options: {
    readonly width?: number;
    readonly fontSize?: number;
    readonly adjustments?: Readonly<Record<string, number>>;
  } = {},
): void => {
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(y),
    w: inches(options.width ?? 6),
    h: inches(1.25),
    text,
  });
  setParagraphAlignment(shape, 0, 'center');
  setShapeRunFormat(shape, 0, 0, {
    font: 'Arial',
    size: options.fontSize ?? 32,
    bold: true,
    color: '#3659E3',
  });
  setShapeTextFromWordArt(shape, true);
  setShapeTextWrap(shape, 'none');
  setShapeTextWarp(shape, {
    preset,
    ...(options.adjustments !== undefined
      ? { adjustments: options.adjustments }
      : preset === 'textArchUp'
        ? { adjustments: { adj: 10800000 } }
        : preset === 'textWave1'
          ? { adjustments: { adj1: 12500, adj2: 0 } }
          : {}),
  });
};

describe('renderSlideToSvg — WordArt preset text-warp preview', () => {
  it('emits representative native presets through the pure-SVG warp path', () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    addWarpedText(slide, 'ARCH UP', 0.75, 'textArchUp');
    addWarpedText(slide, 'WAVE FLOW', 2.5, 'textWave1');
    addWarpedText(slide, 'INFLATE', 4.25, 'textInflate');

    const svg = renderSlideToSvg(presentation, slide, { textLayout: 'svg' });
    expect(svg).not.toContain('<foreignObject');
    expect(svg).toContain('data-text-warp-preview="textArchUp"');
    expect(svg).toContain('data-text-warp-preview="textWave1"');
    expect(svg).toContain('data-text-warp-preview="textInflate"');
    expect(svg).toContain('ARCH UP');
    expect(svg).toContain('WAVE FLOW');
    expect(svg).toContain('INFLATE');
  });

  it('rasterizes the SVG approximation into a deterministic non-blank PNG', () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    addWarpedText(slide, 'ARCH UP', 0.75, 'textArchUp');
    addWarpedText(slide, 'WAVE FLOW', 2.5, 'textWave1');
    addWarpedText(slide, 'INFLATE', 4.25, 'textInflate');

    const first = renderSlideToRgba(presentation, slide, { width: 640 });
    const second = renderSlideToRgba(presentation, slide, { width: 640 });
    expect(first.png.slice(0, 4)).toEqual(new Uint8Array([0x89, 0x50, 0x4e, 0x47]));
    expect(Buffer.from(first.png).equals(Buffer.from(second.png))).toBe(true);
    const arch = foregroundBounds(first.image, 0, 110);
    const wave = foregroundBounds(first.image, 110, 195);
    const inflate = foregroundBounds(first.image, 195, 300);
    expect(arch.width).toBeGreaterThan(240);
    expect(wave.width).toBeGreaterThan(240);
    expect(wave.height).toBeGreaterThan(50);
    expect(inflate.width).toBeGreaterThanOrEqual(240);
    expect(inflate.height).toBeGreaterThan(50);
    let nonWhite = 0;
    for (let offset = 0; offset < first.image.data.length; offset += 4) {
      if (
        first.image.data[offset] !== 255 ||
        first.image.data[offset + 1] !== 255 ||
        first.image.data[offset + 2] !== 255
      ) {
        nonWhite += 1;
      }
    }
    expect(nonWhite).toBeGreaterThan(500);
  });

  it('uses the native textInflate guide to change preview geometry', () => {
    const render = (adjustment: number) => {
      const presentation = createPresentation();
      const slide = addBlankSlide(presentation);
      addWarpedText(slide, 'INFLATE', 1, 'textInflate', { adjustments: { adj: adjustment } });
      return renderSlideToRgba(presentation, slide, { width: 640 });
    };
    const flat = render(0);
    const inflated = render(20000);

    expect(Buffer.from(flat.png).equals(Buffer.from(inflated.png))).toBe(false);
    expect(foregroundBounds(inflated.image, 35, 130).height).toBeGreaterThan(
      foregroundBounds(flat.image, 35, 130).height,
    );
  });

  it('maps long or oversized WordArt into the authored frame instead of overflowing', () => {
    const render = (text: string, fontSize: number) => {
      const presentation = createPresentation();
      const slide = addBlankSlide(presentation);
      addWarpedText(slide, text, 1, 'textInflate', { width: 2, fontSize });
      return foregroundBounds(
        renderSlideToRgba(presentation, slide, { width: 640 }).image,
        35,
        135,
      );
    };
    const short = render('FIT', 32);
    const long = render('A MUCH LONGER WORDART LABEL', 32);
    const oversized = render('FIT', 160);

    expect(Math.abs(long.width - short.width)).toBeLessThanOrEqual(12);
    expect(long.width).toBeLessThanOrEqual(115);
    expect(Math.abs(oversized.height - short.height)).toBeLessThanOrEqual(25);
    expect(oversized.height).toBeLessThanOrEqual(90);
  });
});
