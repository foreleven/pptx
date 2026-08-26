// WordArt identity and preset text warp are body-level text properties with no CSS equivalent.

import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  TEXT_WARP_PRESET_ADJUSTMENTS,
  TEXT_WARP_PRESETS,
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getShapeTextFromWordArt,
  getShapeTextFlatTextZ,
  getShapeTextWarp,
  getShapeTextWarpRaw,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeTextFromWordArt,
  setShapeTextFlatTextZ,
  setShapeTextWarp,
} from '../src/api/index.ts';

interface AdjustmentRange {
  readonly min: number;
  readonly max: number;
}

/** Read guide domains independently from the checked-in ECMA preset definitions. */
const readEcmaAdjustmentRanges = async (): Promise<
  Readonly<Record<string, Readonly<Record<string, AdjustmentRange>>>>
> => {
  const xml = await readFile(
    new URL(
      '../references/ecma-376-5th/ECMA-376/OfficeOpenXML-DrawingMLGeometries/presetTextWarpDefinitions.xml',
      import.meta.url,
    ),
    'utf8',
  );
  const result: Record<string, Record<string, AdjustmentRange>> = {};
  for (const preset of TEXT_WARP_PRESETS) {
    const presetBody = new RegExp(`<${preset}>([\\s\\S]*?)</${preset}>`, 'u').exec(xml)?.[1];
    if (preset === 'textNoShape' && presetBody === undefined) {
      result[preset] = {};
      continue;
    }
    if (presetBody === undefined) throw new Error(`Missing ECMA text-warp preset ${preset}.`);
    const handles = /<ahLst[^>]*>([\s\S]*?)<\/ahLst>/u.exec(presetBody)?.[1] ?? '';
    const ranges: Record<string, AdjustmentRange> = {};
    for (const handle of handles.matchAll(/<ah(?:XY|Polar)\s+([^>]*)>/gu)) {
      const attributes = Object.fromEntries(
        [...handle[1]!.matchAll(/(\w+)="([^"]*)"/gu)].map((match) => [match[1]!, match[2]!]),
      );
      for (const axis of ['X', 'Y', 'R', 'Ang'] as const) {
        const name = attributes[`gdRef${axis}`];
        const min = Number(attributes[`min${axis}`]);
        const max = Number(attributes[`max${axis}`]);
        if (name !== undefined && Number.isSafeInteger(min) && Number.isSafeInteger(max)) {
          ranges[name] = { min, max };
        }
      }
    }
    result[preset] = ranges;
  }
  return result;
};

describe('fn API: WordArt preset text warp', () => {
  it('round-trips flat text Z while preserving absence and explicit zero', async () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Flat text',
    });

    expect(getShapeTextFlatTextZ(shape)).toBeNull();
    setShapeTextFlatTextZ(shape, 0);
    expect(getShapeTextFlatTextZ(shape)).toBe(0);
    setShapeTextFlatTextZ(shape, -12_700);
    expect(getShapeTextFlatTextZ(shape)).toBe(-12_700);

    const reloaded = await loadPresentation(await savePresentation(presentation));
    const reloadedShape = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getShapeTextFlatTextZ(reloadedShape)).toBe(-12_700);

    setShapeTextFlatTextZ(reloadedShape, null);
    expect(getShapeTextFlatTextZ(reloadedShape)).toBeNull();
  });

  it('validates flat text Z and rejects the mutually exclusive sp3d branch before mutation', async () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Flat text bounds',
    });
    setShapeTextFlatTextZ(shape, 12_700);
    for (const value of [1.5, Number.NaN, 27_273_042_316_901, -27_273_042_329_601]) {
      expect(() => setShapeTextFlatTextZ(shape, value)).toThrow(/safe integer|ST_Coordinate/u);
      expect(getShapeTextFlatTextZ(shape)).toBe(12_700);
    }

    const entries = unzipSync(await savePresentation(presentation));
    const slidePart = 'ppt/slides/slide1.xml';
    entries[slidePart] = strToU8(
      strFromU8(entries[slidePart]!).replace('<a:flatTx z="12700"/>', '<a:sp3d/>'),
    );
    const imported = await loadPresentation(zipSync(entries));
    const importedShape = getSlideShapes(getSlides(imported)[0]!)[0]!;
    expect(getShapeTextFlatTextZ(importedShape)).toBeNull();
    expect(() => setShapeTextFlatTextZ(importedShape, 0)).toThrow(/sp3d/u);
    expect(getShapeTextFlatTextZ(importedShape)).toBeNull();
  });

  it('round-trips every preset plus deterministic constant-value guides', async () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Editable WordArt',
    });

    expect(getShapeTextFromWordArt(shape)).toBeNull();
    expect(getShapeTextWarp(shape)).toBeNull();
    for (const preset of TEXT_WARP_PRESETS) {
      setShapeTextWarp(shape, { preset });
      expect(getShapeTextWarp(shape)).toEqual({ preset });
    }

    setShapeTextFromWordArt(shape, true);
    setShapeTextWarp(shape, {
      preset: 'textWave1',
      adjustments: { adj2: -2500, adj1: 12500 },
    });
    expect(getShapeTextFromWordArt(shape)).toBe(true);
    expect(getShapeTextWarp(shape)).toEqual({
      preset: 'textWave1',
      adjustments: { adj1: 12500, adj2: -2500 },
    });
    expect(getShapeTextWarpRaw(shape)).toEqual({
      preset: 'textWave1',
      guides: [
        { name: 'adj1', formula: 'val 12500' },
        { name: 'adj2', formula: 'val -2500' },
      ],
      exact: true,
    });

    const reloaded = await loadPresentation(await savePresentation(presentation));
    const reloadedShape = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getShapeTextFromWordArt(reloadedShape)).toBe(true);
    expect(getShapeTextWarp(reloadedShape)).toEqual({
      preset: 'textWave1',
      adjustments: { adj1: 12500, adj2: -2500 },
    });

    setShapeTextFromWordArt(reloadedShape, false);
    expect(getShapeTextFromWordArt(reloadedShape)).toBe(false);
    setShapeTextFromWordArt(reloadedShape, null);
    setShapeTextWarp(reloadedShape, null);
    expect(getShapeTextFromWordArt(reloadedShape)).toBeNull();
    expect(getShapeTextWarpRaw(reloadedShape)).toBeNull();
  });

  it('enforces every ECMA preset adjustment boundary before mutation', async () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Boundary WordArt',
    });
    const ecmaRanges = await readEcmaAdjustmentRanges();
    expect(TEXT_WARP_PRESET_ADJUSTMENTS).toEqual(ecmaRanges);

    for (const preset of TEXT_WARP_PRESETS) {
      const ranges = ecmaRanges[preset]!;
      for (const [name, range] of Object.entries(ranges)) {
        setShapeTextWarp(shape, { preset, adjustments: { [name]: range.min } });
        expect(getShapeTextWarp(shape)).toEqual({ preset, adjustments: { [name]: range.min } });
        setShapeTextWarp(shape, { preset, adjustments: { [name]: range.max } });
        expect(getShapeTextWarp(shape)).toEqual({ preset, adjustments: { [name]: range.max } });
        expect(() =>
          setShapeTextWarp(shape, { preset, adjustments: { [name]: range.min - 1 } }),
        ).toThrow(/must be between/u);
        expect(() =>
          setShapeTextWarp(shape, { preset, adjustments: { [name]: range.max + 1 } }),
        ).toThrow(/must be between/u);
        expect(getShapeTextWarp(shape)).toEqual({ preset, adjustments: { [name]: range.max } });
      }
      expect(() => setShapeTextWarp(shape, { preset, adjustments: { unsupported: 0 } })).toThrow(
        /does not support adjustment/u,
      );
    }
  });

  it('does not expose imported guides outside the editable preset domain', async () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Imported WordArt',
    });
    setShapeTextWarp(shape, { preset: 'textInflate', adjustments: { adj: 18750 } });
    const entries = unzipSync(await savePresentation(presentation));
    const slidePart = 'ppt/slides/slide1.xml';
    entries[slidePart] = strToU8(
      strFromU8(entries[slidePart]!).replace('fmla="val 18750"', 'fmla="val 20001"'),
    );
    const imported = await loadPresentation(zipSync(entries));
    const importedShape = getSlideShapes(getSlides(imported)[0]!)[0]!;

    expect(getShapeTextWarpRaw(importedShape)).toMatchObject({
      preset: 'textInflate',
      guides: [{ name: 'adj', formula: 'val 20001' }],
      exact: true,
    });
    expect(getShapeTextWarp(importedShape)).toBeNull();
  });

  it('rejects invalid presets and guides before replacing the current warp', () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Validated WordArt',
    });
    setShapeTextWarp(shape, { preset: 'textArchUp', adjustments: { adj: 50000 } });

    expect(() => setShapeTextWarp(shape, { preset: 'not-a-preset' as never })).toThrow(
      /Unsupported preset text warp/u,
    );
    expect(() =>
      setShapeTextWarp(shape, { preset: 'textArchUp', adjustments: { 'bad name': 1 } }),
    ).toThrow(/without whitespace/u);
    expect(() =>
      setShapeTextWarp(shape, { preset: 'textArchUp', adjustments: { adj: 1.5 } }),
    ).toThrow(/safe integer/u);
    expect(getShapeTextWarp(shape)).toEqual({
      preset: 'textArchUp',
      adjustments: { adj: 50000 },
    });
  });
});
