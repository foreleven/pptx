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
  getParagraphBullet,
  getParagraphBulletPropertiesEffective,
  getSlidePartName,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphBullet,
} from '../src/api/index.ts';

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
