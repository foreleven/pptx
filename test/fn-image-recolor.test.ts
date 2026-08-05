import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getShapeImageBiLevelThreshold,
  getShapeImageDuotone,
  getShapeKind,
  getSlideShapes,
  getSlides,
  isShapeImageGrayscale,
  loadPresentation,
  setShapeImageBiLevel,
  setShapeImageDuotone,
  setShapeImageGrayscale,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: native image recolor', () => {
  it('round-trips grayscale, bi-level, and duotone transforms', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres)[0]!).find(
      (shape) => getShapeKind(shape) === 'picture',
    )!;

    setShapeImageGrayscale(picture, true);
    setShapeImageBiLevel(picture, 0.42);
    setShapeImageDuotone(picture, { dark: '#112233', light: '#DDEEFF' });

    expect(isShapeImageGrayscale(picture)).toBe(true);
    expect(getShapeImageBiLevelThreshold(picture)).toBeCloseTo(42);
    expect(getShapeImageDuotone(pres, picture)).toEqual({
      firstColor: '#112233',
      secondColor: '#DDEEFF',
    });

    setShapeImageGrayscale(picture, false);
    setShapeImageBiLevel(picture, null);
    setShapeImageDuotone(picture, null);
    expect(isShapeImageGrayscale(picture)).toBe(false);
    expect(getShapeImageBiLevelThreshold(picture)).toBeNull();
    expect(getShapeImageDuotone(pres, picture)).toBeNull();
  });

  it('rejects an invalid bi-level threshold', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres)[0]!).find(
      (shape) => getShapeKind(shape) === 'picture',
    )!;
    expect(() => setShapeImageBiLevel(picture, -0.1)).toThrow(RangeError);
    expect(() => setShapeImageBiLevel(picture, 1.1)).toThrow(RangeError);
  });
});
