// getParagraphBullet — read back the bullet style on a single paragraph.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  findShapeByText,
  getMediaParts,
  getParagraphBullet,
  getParagraphBulletImageBytes,
  getParagraphBulletPropertiesEffective,
  getParagraphBulletStyle,
  getParagraphIndent,
  getSlidePartName,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphBullet,
  setParagraphBulletImage,
} from '../src/api/index.ts';

const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const ALT_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getParagraphBullet', () => {
  it('returns null when no bullet is set on the paragraph', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'plain',
    });
    expect(getParagraphBullet(tb, 0)).toBeNull();
  });

  it('round-trips bullet / number / none / custom char / autoNum / explicit start', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A\nB\nC\nD\nE\nF',
    });
    setParagraphBullet(tb, 0, 'bullet');
    setParagraphBullet(tb, 1, 'number');
    setParagraphBullet(tb, 2, 'none');
    setParagraphBullet(tb, 3, { char: '★', font: 'Arial' });
    setParagraphBullet(tb, 4, { autoNum: 'romanLcPeriod', font: 'Aptos' });
    setParagraphBullet(tb, 5, { autoNum: 'arabicPeriod', startAt: 5 });

    expect(getParagraphBullet(tb, 0)).toBe('bullet');
    expect(getParagraphBullet(tb, 1)).toBe('number');
    expect(getParagraphBullet(tb, 2)).toBe('none');
    expect(getParagraphBullet(tb, 3)).toEqual({ char: '★' });
    expect(getParagraphBullet(tb, 4)).toEqual({ autoNum: 'romanLcPeriod' });
    expect(getParagraphBullet(tb, 5)).toEqual({ autoNum: 'arabicPeriod', startAt: 5 });
    expect(getParagraphBulletPropertiesEffective(pres, tb, 3).font).toBe('Arial');
    expect(getParagraphBulletPropertiesEffective(pres, tb, 4).font).toBe('Aptos');

    const rebuilt = await loadPresentation(await savePresentation(pres));
    const rebuiltText = findShapeByText(getSlides(rebuilt)[0]!, 'A')!;
    expect(getParagraphBullet(rebuiltText, 5)).toEqual({ autoNum: 'arabicPeriod', startAt: 5 });
    expect(getParagraphBulletPropertiesEffective(rebuilt, rebuiltText, 3).font).toBe('Arial');
    expect(getParagraphBulletPropertiesEffective(rebuilt, rebuiltText, 4).font).toBe('Aptos');
  });

  it('writes picture bullets, shares identical media, and collects stale relationships', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const tb = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: 'First\nSecond',
    });

    setParagraphBulletImage(tb, 0, PNG);
    setParagraphBulletImage(tb, 1, PNG);
    expect(getParagraphBulletImageBytes(tb, 0)).toEqual(PNG);
    expect(getParagraphBulletImageBytes(tb, 1)).toEqual(PNG);
    expect(getParagraphIndent(tb, 0)).toEqual({
      leftEmu: 342900,
      rightEmu: null,
      firstLineEmu: -342900,
    });
    expect(getMediaParts(pres)).toHaveLength(1);

    const rebuilt = await loadPresentation(await savePresentation(pres));
    const rebuiltText = findShapeByText(getSlides(rebuilt)[0]!, 'First')!;
    expect(getParagraphBulletImageBytes(rebuiltText, 0)).toEqual(PNG);
    expect(getParagraphBulletImageBytes(rebuiltText, 1)).toEqual(PNG);

    setParagraphBulletImage(rebuiltText, 0, ALT_PNG);
    expect(getMediaParts(rebuilt)).toHaveLength(2);
    setParagraphBullet(rebuiltText, 1, 'bullet');
    expect(getMediaParts(rebuilt).map((media) => media.data)).toEqual([ALT_PNG]);
    setParagraphBullet(rebuiltText, 0, 'none');
    expect(getMediaParts(rebuilt)).toEqual([]);
  });

  it('rejects undetectable picture-bullet bytes before mutating the paragraph', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const tb = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Stable marker',
    });
    setParagraphBullet(tb, 0, { char: '★' });

    expect(() => setParagraphBulletImage(tb, 0, new Uint8Array([0, 1, 2]))).toThrow(/format/u);
    expect(getParagraphBullet(tb, 0)).toEqual({ char: '★' });
    expect(getMediaParts(pres)).toEqual([]);
  });

  it('rejects automatic-number starts outside the DrawingML range', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(1),
      text: 'Invalid start',
    });

    expect(() => setParagraphBullet(tb, 0, { autoNum: 'arabicPeriod', startAt: 0 })).toThrow(
      'automatic-number startAt must be an integer from 1 to 32767',
    );
    expect(() => setParagraphBullet(tb, 0, { autoNum: 'arabicPeriod', startAt: 32768 })).toThrow(
      'automatic-number startAt must be an integer from 1 to 32767',
    );
  });

  it('rejects an empty marker font before mutating the paragraph', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const tb = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: 'A',
    });
    setParagraphBullet(tb, 0, { char: '•', font: 'Arial' });

    expect(() => setParagraphBullet(tb, 0, { char: '•', font: '   ' })).toThrow(
      /bullet font must be a non-empty string/u,
    );
    expect(getParagraphBullet(tb, 0)).toBe('bullet');
    expect(getParagraphBulletPropertiesEffective(pres, tb, 0).font).toBe('Arial');
  });

  it('writes editable marker color plus relative or fixed size in schema order', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const tb = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: 'Relative\nFixed',
    });
    setParagraphBullet(tb, 0, { char: '▪', color: '#F26B5B', sizePct: 0.8 });
    setParagraphBullet(tb, 1, { char: '▪', color: '#3659E380', sizePts: 12 });

    expect(getParagraphBulletStyle(pres, tb, 0)).toEqual({
      color: '#F26B5B',
      sizePct: 0.8,
      sizePts: null,
      font: null,
    });
    expect(getParagraphBulletStyle(pres, tb, 1)).toEqual({
      color: '#3659E380',
      sizePct: null,
      sizePts: 12,
      font: null,
    });

    const pkg = _internalPackageOf(pres);
    const slidePartName = getSlidePartName(slide) as Parameters<typeof pkg.getPart>[0];
    const xml = new TextDecoder().decode(pkg.getPart(slidePartName)!.data);
    expect(xml).toContain(
      '<a:buClr><a:srgbClr val="F26B5B"/></a:buClr><a:buSzPct val="80000"/><a:buChar char="▪"/>',
    );
    expect(xml).toContain(
      '<a:buClr><a:srgbClr val="3659E3"><a:alpha val="50196"/></a:srgbClr></a:buClr><a:buSzPts val="1200"/><a:buChar char="▪"/>',
    );

    const rebuilt = await loadPresentation(await savePresentation(pres));
    const rebuiltText = findShapeByText(getSlides(rebuilt)[0]!, 'Relative')!;
    expect(getParagraphBulletStyle(rebuilt, rebuiltText, 0).sizePct).toBe(0.8);
    expect(getParagraphBulletStyle(rebuilt, rebuiltText, 1).sizePts).toBe(12);
  });

  it('rejects invalid marker sizes atomically', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const tb = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: 'Stable',
    });
    setParagraphBullet(tb, 0, { char: '•', color: '#F26B5B', sizePct: 0.8 });

    expect(() => setParagraphBullet(tb, 0, { char: '▪', sizePct: 0.2 })).toThrow(/0\.25 to 4/u);
    expect(() => setParagraphBullet(tb, 0, { char: '▪', sizePts: 4001 })).toThrow(
      /bullet sizePts/u,
    );
    expect(() => setParagraphBullet(tb, 0, { char: '▪', sizePct: 1, sizePts: 12 })).toThrow(
      /mutually exclusive/u,
    );
    expect(getParagraphBullet(tb, 0)).toBe('bullet');
    expect(getParagraphBulletStyle(pres, tb, 0)).toMatchObject({ color: '#F26B5B', sizePct: 0.8 });
  });

  it('resolves marker overrides independently from a direct bullet identity', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(1),
      text: 'Inherited marker',
    });
    setParagraphBullet(tb, 0, { char: '▪' });
    const pkg = _internalPackageOf(pres);
    const slidePartName = getSlidePartName(slide) as Parameters<typeof pkg.getPart>[0];
    const slidePart = pkg.getPart(slidePartName)!;
    const xml = new TextDecoder().decode(slidePart.data);
    slidePart.data = new TextEncoder().encode(
      xml.replace(
        '<a:lstStyle/>',
        '<a:lstStyle><a:lvl1pPr><a:buSzPct val="80%"/><a:buFont typeface="Wingdings"/></a:lvl1pPr></a:lstStyle>',
      ),
    );

    const rebuilt = await loadPresentation(await savePresentation(pres));
    const shape = findShapeByText(getSlides(rebuilt)[0]!, 'Inherited marker')!;
    expect(getParagraphBulletPropertiesEffective(rebuilt, shape, 0)).toMatchObject({
      bullet: { char: '▪' },
      picture: false,
      sizePct: 0.8,
      font: 'Wingdings',
    });
  });

  it('resolves picture bullets inherited from a text-body list style', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(1),
      text: 'Inherited picture',
    });
    const pkg = _internalPackageOf(pres);
    const slidePartName = getSlidePartName(slide) as Parameters<typeof pkg.getPart>[0];
    const slidePart = pkg.getPart(slidePartName)!;
    const xml = new TextDecoder().decode(slidePart.data);
    slidePart.data = new TextEncoder().encode(
      xml.replace(
        '<a:lstStyle/>',
        '<a:lstStyle><a:lvl1pPr><a:buBlip><a:blip r:embed="rId999"/></a:buBlip></a:lvl1pPr></a:lstStyle>',
      ),
    );

    const rebuilt = await loadPresentation(await savePresentation(pres));
    const shape = findShapeByText(getSlides(rebuilt)[0]!, 'Inherited picture')!;
    expect(getParagraphBulletPropertiesEffective(rebuilt, shape, 0)).toMatchObject({
      bullet: null,
      picture: true,
    });
  });
});
