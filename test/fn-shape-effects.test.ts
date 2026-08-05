// Shape effects: shadow + glow.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  clearShapeEffects,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeEffects,
  setShapeGlow,
  setShapeShadow,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

describe('fn API: shape effects', () => {
  it('setShapeShadow writes an outerShdw with computed direction', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapeShadow(shape, { color: '#000000', angleDeg: 45, opacity: 0.5 });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:effectLst>');
    expect(xml).toContain('<a:outerShdw');
    expect(xml).toContain('dir="2700000"');
    expect(xml).toContain('val="50000"');
  });

  it('setShapeGlow writes a glow with the configured radius', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapeGlow(shape, { color: '#FF0000', radiusEmu: 90000 });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:glow');
    expect(xml).toContain('rad="90000"');
    expect(xml).toContain('FF0000');
  });

  it('shadow then glow replaces the prior effect list', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapeShadow(shape);
    setShapeGlow(shape, { color: '#00FF00' });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:glow');
    expect(xml).not.toContain('<a:outerShdw');
  });

  it('clearShapeEffects removes any effectLst', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapeShadow(shape);
    expect(await slideXml(await savePresentation(pres), 0)).toContain('<a:effectLst>');
    clearShapeEffects(shape);
    expect(await slideXml(await savePresentation(pres), 0)).not.toContain('<a:effectLst>');
  });

  it('setShapeEffects writes one schema-ordered composed effect list', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapeEffects(shape, [
      { kind: 'softEdge', radiusEmu: 76200 },
      { kind: 'outerShdw', color: '#112233', blurEmu: 50800, distEmu: 25400 },
      { kind: 'glow', color: '#445566', radiusEmu: 63500, opacity: 0.5 },
      { kind: 'blur', radiusEmu: 38100, grow: false },
      { kind: 'reflection', blurEmu: 25400, distEmu: 101600, startOpacity: 0.7, endOpacity: 0.1 },
      { kind: 'innerShdw', color: '#778899', blurEmu: 12700, distEmu: 6350 },
    ]);

    const xml = await slideXml(await savePresentation(pres), 0);
    const tags = [
      '<a:blur ',
      '<a:glow ',
      '<a:innerShdw ',
      '<a:outerShdw ',
      '<a:reflection ',
      '<a:softEdge ',
    ];
    expect(tags.every((tag) => xml.includes(tag))).toBe(true);
    for (let index = 1; index < tags.length; index++) {
      expect(xml.indexOf(tags[index]!)).toBeGreaterThan(xml.indexOf(tags[index - 1]!));
    }
    expect(xml).toContain('grow="0"');
    expect(xml).toContain('val="50000"');
  });
});
