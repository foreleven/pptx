import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getSlides,
  inches,
  loadPresentation,
  setShapeRunFormat,
} from '../src/api/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';

const fixture = fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url));

const render = async (
  format: Parameters<typeof setShapeRunFormat>[3],
): Promise<ReturnType<typeof renderSlideToRgba>> => {
  const presentation = await loadPresentation(await readFile(fixture));
  const slide = getSlides(presentation)[0]!;
  const text = addSlideTextBox(slide, {
    x: inches(0.5),
    y: inches(0.5),
    w: inches(7),
    h: inches(1.5),
    text: 'Decoration',
  });
  setShapeRunFormat(text, 0, 0, { size: 54, font: 'Arial', ...format });
  return renderSlideToRgba(presentation, slide, { width: 640 });
};

const coloredPixels = (
  pixels: Uint8Array,
  predicate: (red: number, green: number, blue: number, alpha: number) => boolean,
): number => {
  let count = 0;
  for (let index = 0; index < pixels.length; index += 4) {
    if (predicate(pixels[index]!, pixels[index + 1]!, pixels[index + 2]!, pixels[index + 3]!))
      count += 1;
  }
  return count;
};

describe('text decoration PNG rendering', () => {
  it('rasterizes gradient underline and pattern outline paints instead of dropping them', async () => {
    const { image, png } = await render({
      color: '#111111',
      underline: true,
      underlineFill: {
        kind: 'gradient',
        stops: [
          { offset: 0, color: '#3659E3' },
          { offset: 1, color: '#F26B5B' },
        ],
        angleDeg: 0,
      },
      outline: {
        kind: 'pattern',
        preset: 'pct20',
        foreground: '#3659E3',
        background: '#FFFFFF',
        widthPt: 1.5,
      },
    });
    expect(Array.from(png.slice(0, 8))).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(
      coloredPixels(image.data, (r, g, b, a) => a > 200 && b > r + 35 && b > g + 35),
    ).toBeGreaterThan(10);
    expect(
      coloredPixels(image.data, (r, g, b, a) => a > 200 && r > b + 45 && r > g + 35),
    ).toBeGreaterThan(10);
  });

  it('keeps deterministic solid fallbacks visible for group fill and omitted outline arrows', async () => {
    const baseline = await render({ color: '#111111' });
    const degraded = await render({
      color: '#111111',
      underline: true,
      underlineFill: { kind: 'group' },
      outline: {
        kind: 'solid',
        color: '#F26B5B',
        widthPt: 1.5,
        head: { type: 'triangle' },
        tail: { type: 'oval' },
      },
    });
    let changed = 0;
    for (let index = 0; index < baseline.image.data.length; index += 1) {
      if (baseline.image.data[index] !== degraded.image.data[index]) changed += 1;
    }
    expect(changed).toBeGreaterThan(250);
    expect(
      coloredPixels(degraded.image.data, (r, g, b, a) => a > 200 && r > 150 && g > 35 && b < 150),
    ).toBeGreaterThan(10);
  });
});
