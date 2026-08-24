// Two small additions: slide visibility (`show="0"`) and stroke dash
// patterns (`<a:prstDash>`).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getShapeStrokeCustomDash,
  getShapeStrokeDash,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  inches,
  isSlideHidden,
  loadPresentation,
  savePresentation,
  setShapeStroke,
  setShapeStrokeCustomDash,
  setShapeStrokeDash,
  setSlideHidden,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

describe('fn API: slide visibility', () => {
  it('toggles <p:sld show="0"/> correctly', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(isSlideHidden(slide)).toBe(false);

    setSlideHidden(slide, true);
    expect(isSlideHidden(slide)).toBe(true);
    expect(await slideXml(await savePresentation(pres), 0)).toContain('show="0"');

    setSlideHidden(slide, false);
    expect(isSlideHidden(slide)).toBe(false);
    expect(await slideXml(await savePresentation(pres), 0)).not.toContain('show="0"');
  });
});

describe('fn API: setShapeStrokeDash', () => {
  it('writes <a:prstDash val="..."/>', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapeStroke(shape, { color: '#000000', widthEmu: 12700 });
    setShapeStrokeDash(shape, 'dash');
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:prstDash val="dash"');
  });

  it('reapplying replaces the prior dash choice', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapeStroke(shape, { color: '#FF0000' });
    setShapeStrokeDash(shape, 'dot');
    setShapeStrokeDash(shape, 'dashDot');
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('val="dashDot"');
    expect(xml).not.toContain('val="dot"');
  });

  it('round-trips a custom dash and replaces the preset dash choice', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'line',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(0),
    });
    setShapeStrokeDash(shape, 'dash');
    setShapeStrokeCustomDash(shape, [
      { dash: 250_000, space: 150_000 },
      { dash: 50_000, space: 150_000 },
    ]);

    expect(getShapeStrokeCustomDash(shape)).toEqual([
      { dash: 250_000, space: 150_000 },
      { dash: 50_000, space: 150_000 },
    ]);
    expect(await slideXml(await savePresentation(pres), 0)).toContain(
      '<a:custDash><a:ds d="250000" sp="150000"/><a:ds d="50000" sp="150000"/></a:custDash>',
    );
    expect(getShapeStrokeDash(shape)).toBeNull();

    const entries = unzipSync(await savePresentation(pres));
    const slidePart = entries['ppt/slides/slide1.xml']!;
    entries['ppt/slides/slide1.xml'] = strToU8(
      strFromU8(slidePart)
        .replace('d="250000" sp="150000"', 'd="250%" sp="150%"')
        .replace('d="50000" sp="150000"', 'd="50%" sp="150%"'),
    );
    const strict = await loadPresentation(zipSync(entries));
    expect(getShapeStrokeCustomDash(getSlideShapes(getSlides(strict)[0]!).at(-1)!)).toEqual([
      { dash: 250_000, space: 150_000 },
      { dash: 50_000, space: 150_000 },
    ]);
  });

  it('rejects invalid custom dash sequences while preserving fractional Strict units', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const shape = addSlideShape(getSlides(pres)[0]!, {
      preset: 'line',
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(0),
    });

    expect(() => setShapeStrokeCustomDash(shape, [])).toThrow(/must not be empty/u);
    expect(() => setShapeStrokeCustomDash(shape, [{ dash: -1, space: 100_000 }])).toThrow(
      /non-negative/u,
    );
    expect(() => setShapeStrokeCustomDash(shape, [{ dash: Number.NaN, space: 100_000 }])).toThrow(
      /finite/u,
    );
    expect(() =>
      setShapeStrokeCustomDash(shape, [{ dash: Number.POSITIVE_INFINITY, space: 100_000 }]),
    ).toThrow(/finite/u);
    expect(() =>
      setShapeStrokeCustomDash(shape, [{ dash: Number.MIN_VALUE, space: 100_000 }]),
    ).toThrow(/too small/u);

    setShapeStrokeCustomDash(shape, [{ dash: 1e-10, space: 100_000 }]);
    expect(getShapeStrokeCustomDash(shape)).toEqual([{ dash: 1e-10, space: 100_000 }]);
    expect(await slideXml(await savePresentation(pres), 0)).toContain('d="0.0000000000001%"');
  });
});
