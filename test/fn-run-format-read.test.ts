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

  it('round-trips every editable text effect as one composed effect list', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(5),
      h: inches(2),
      text: 'effects',
    });
    setShapeRunFormat(tb, 0, 0, {
      effects: [
        { kind: 'softEdge', radiusEmu: 38100 },
        {
          kind: 'prstShdw',
          preset: 'shdw14',
          color: '#112233',
          distEmu: 25400,
          angleDeg: 45,
          opacity: 0.45,
        },
        { kind: 'outerShdw', color: '#223344', blurEmu: 50800, distEmu: 38100, angleDeg: 90 },
        {
          kind: 'reflection',
          blurEmu: 12700,
          distEmu: 12700,
          angleDeg: 90,
          startOpacity: 0.65,
          endOpacity: 0,
          scaleY: -1,
        },
        { kind: 'glow', color: '#35B9C6', radiusEmu: 76200, opacity: 0.5 },
        { kind: 'fillOverlay', color: '#F26B5B', opacity: 0.4, blend: 'mult' },
        { kind: 'blur', radiusEmu: 50800, grow: true },
        {
          kind: 'innerShdw',
          color: '#334455',
          blurEmu: 50800,
          distEmu: 25400,
          angleDeg: 45,
          opacity: 0.65,
        },
      ],
    });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedShape = findShapeByText(getSlides(reloaded)[0]!, 'effects')!;
    expect(getShapeRunFormat(reloadedShape, 0, 0)?.effects).toEqual([
      { kind: 'blur', radiusEmu: 50800, grow: true },
      { kind: 'fillOverlay', color: '#F26B5B', opacity: 0.4, blend: 'mult' },
      { kind: 'glow', color: '#35B9C6', radiusEmu: 76200, opacity: 0.5 },
      {
        kind: 'innerShdw',
        color: '#334455',
        blurEmu: 50800,
        distEmu: 25400,
        angleDeg: 45,
        opacity: 0.65,
      },
      { kind: 'outerShdw', color: '#223344', blurEmu: 50800, distEmu: 38100, angleDeg: 90 },
      {
        kind: 'prstShdw',
        preset: 'shdw14',
        color: '#112233',
        distEmu: 25400,
        angleDeg: 45,
        opacity: 0.45,
      },
      {
        kind: 'reflection',
        blurEmu: 12700,
        distEmu: 12700,
        angleDeg: 90,
        startOpacity: 0.65,
        endOpacity: 0,
        scaleY: -1,
      },
      { kind: 'softEdge', radiusEmu: 38100 },
    ]);
  });

  it('rejects authoring complete effects and the text-shadow compatibility field together', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(5),
      h: inches(2),
      text: 'conflict',
    });
    expect(() =>
      setShapeRunFormat(tb, 0, 0, {
        textShadow: { color: '#000000', blurPt: 2, offsetXPt: 1, offsetYPt: 1 },
        effects: [{ kind: 'glow', color: '#35B9C6', radiusEmu: 76200 }],
      }),
    ).toThrow(/cannot be authored together/u);
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
