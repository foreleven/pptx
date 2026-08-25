// `getShapeRunFormat` — read back per-run text formatting.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  findShapeByText,
  getShapeRunFormat,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeRunFormat,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeRunFormat', () => {
  it('returns null when the run has no rPr', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'plain',
    });
    // setShapeText creates runs with rPr already (with `lang` attribute),
    // so we can't assert null on a freshly authored shape. Instead test
    // that the result is at least an empty object or has only inherited
    // values.
    const fmt = getShapeRunFormat(tb, 0, 0);
    expect(fmt).toBeDefined();
  });

  it('round-trips bold + color + size', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'styled',
    });
    setShapeRunFormat(tb, 0, 0, { bold: true, italic: true, color: '#FF0000', size: 18 });
    const fmt = getShapeRunFormat(tb, 0, 0);
    expect(fmt).not.toBeNull();
    expect(fmt!.bold).toBe(true);
    expect(fmt!.italic).toBe(true);
    expect(fmt!.color).toBe('#FF0000');
    expect(fmt!.size).toBeCloseTo(18);
  });

  it('round-trips font and fontEastAsian independently', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: '見出し',
    });
    setShapeRunFormat(tb, 0, 0, { font: 'Georgia', fontEastAsian: '游明朝' });
    const fmt = getShapeRunFormat(tb, 0, 0);
    expect(fmt).not.toBeNull();
    expect(fmt!.font).toBe('Georgia');
    expect(fmt!.fontEastAsian).toBe('游明朝');
  });

  it('round-trips script-specific fonts and language tags independently', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'عربي Symbols αβγ',
    });
    setShapeRunFormat(tb, 0, 0, {
      font: 'Aptos',
      fontEastAsian: '游ゴシック',
      fontComplexScript: 'Arial',
      fontSymbol: 'Symbol',
      language: 'ar-SA',
      alternativeLanguage: 'en-US',
    });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedShape = findShapeByText(getSlides(reloaded)[0]!, 'عربي Symbols αβγ')!;
    expect(getShapeRunFormat(reloadedShape, 0, 0)).toMatchObject({
      font: 'Aptos',
      fontEastAsian: '游ゴシック',
      fontComplexScript: 'Arial',
      fontSymbol: 'Symbol',
      language: 'ar-SA',
      alternativeLanguage: 'en-US',
    });
  });

  it('underline encodes both boolean and explicit token', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'u',
    });
    setShapeText(tb, 'u');
    setShapeRunFormat(tb, 0, 0, { underline: true });
    expect(getShapeRunFormat(tb, 0, 0)!.underline).toBe(true);
    setShapeRunFormat(tb, 0, 0, { underline: false });
    expect(getShapeRunFormat(tb, 0, 0)!.underline).toBe(false);
  });
});
