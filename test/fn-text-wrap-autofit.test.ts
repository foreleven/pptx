// Text wrap + auto-fit on bodyPr.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeTextAutoFit,
  getShapeTextAutoFitParams,
  getShapeTextAutoFitParamsRaw,
  getShapeTextWrap,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeTextAutoFit,
  setShapeTextWrap,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setShapeTextWrap', () => {
  it('round-trips both wrap modes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A',
    });
    setShapeTextWrap(tb, 'none');
    expect(getShapeTextWrap(tb)).toBe('none');
    setShapeTextWrap(tb, 'square');
    expect(getShapeTextWrap(tb)).toBe('square');
  });
});

describe('fn API: setShapeTextAutoFit', () => {
  it('round-trips every mode', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A',
    });
    for (const mode of ['none', 'normal', 'shape'] as const) {
      setShapeTextAutoFit(tb, mode);
      expect(getShapeTextAutoFit(tb)).toBe(mode);
    }
    setShapeTextAutoFit(tb, 'normal');
    expect(getShapeTextAutoFitParams(tb)).toBeNull();
  });

  it('replaces the prior auto-fit child each call', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A',
    });
    setShapeTextAutoFit(tb, 'normal');
    setShapeTextAutoFit(tb, 'shape');
    expect(getShapeTextAutoFit(tb)).toBe('shape');
  });

  it('round-trips normal autofit font and line-spacing reductions', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A',
    });

    setShapeTextAutoFit(tb, 'normal', { fontScale: 0.72, lnSpcReduction: 0.18 });
    expect(getShapeTextAutoFit(tb)).toBe('normal');
    expect(getShapeTextAutoFitParams(tb)).toEqual({ fontScale: 0.72, lnSpcReduction: 0.18 });
  });

  it('reads Strict percent lexemes and writes the inclusive zero boundary schema-validly', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A',
    });

    setShapeTextAutoFit(tb, 'normal', { fontScale: 0.72, lnSpcReduction: 0.18 });
    const strictEntries = unzipSync(await savePresentation(pres));
    const slidePart = strictEntries['ppt/slides/slide1.xml']!;
    strictEntries['ppt/slides/slide1.xml'] = strToU8(
      strFromU8(slidePart)
        .replace('fontScale="72000"', 'fontScale="72%"')
        .replace('lnSpcReduction="18000"', 'lnSpcReduction="18%"'),
    );
    const strict = await loadPresentation(zipSync(strictEntries));
    expect(getShapeTextAutoFitParams(getSlideShapes(getSlides(strict)[0]!).at(-1)!)).toEqual({
      fontScale: 0.72,
      lnSpcReduction: 0.18,
    });

    setShapeTextAutoFit(tb, 'normal', { fontScale: 0, lnSpcReduction: 0 });
    const zeroEntries = unzipSync(await savePresentation(pres));
    const zeroXml = strFromU8(zeroEntries['ppt/slides/slide1.xml']!);
    expect(zeroXml).toContain('<a:normAutofit fontScale="0%" lnSpcReduction="0"/>');
    if (isSchemaValidationAvailable()) expectSchemaValid(zeroXml, 'pml');
  });

  it('separates out-of-contract native ratios from safe authored parameters', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A',
    });
    setShapeTextAutoFit(tb, 'normal', { fontScale: 0.72, lnSpcReduction: 0.18 });

    const entries = unzipSync(await savePresentation(pres));
    entries['ppt/slides/slide1.xml'] = strToU8(
      strFromU8(entries['ppt/slides/slide1.xml']!)
        .replace('fontScale="72000"', 'fontScale="110%"')
        .replace('lnSpcReduction="18000"', 'lnSpcReduction="120%"'),
    );
    const strict = await loadPresentation(zipSync(entries));
    const strictShape = getSlideShapes(getSlides(strict)[0]!).at(-1)!;
    expect(getShapeTextAutoFitParamsRaw(strictShape)).toEqual({
      fontScale: 1.1,
      lnSpcReduction: 1.2,
    });
    expect(getShapeTextAutoFitParams(strictShape)).toBeNull();
  });

  it('uses schema defaults for malformed native autofit ratio lexemes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A',
    });
    setShapeTextAutoFit(tb, 'normal', { fontScale: 0.72, lnSpcReduction: 0.18 });
    const sourceEntries = unzipSync(await savePresentation(pres));

    for (const malformed of ['', ' ', '0x10', '1e5', '.72%', '72.%']) {
      const entries = { ...sourceEntries };
      entries['ppt/slides/slide1.xml'] = strToU8(
        strFromU8(entries['ppt/slides/slide1.xml']!).replace(
          'fontScale="72000"',
          `fontScale="${malformed}"`,
        ),
      );
      const loaded = await loadPresentation(zipSync(entries));
      const loadedShape = getSlideShapes(getSlides(loaded)[0]!).at(-1)!;
      expect(getShapeTextAutoFitParamsRaw(loadedShape), malformed).toEqual({
        fontScale: 1,
        lnSpcReduction: 0.18,
      });
    }
  });

  it('applies XML Schema whitespace collapse only to Transitional integer ratios', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A',
    });
    setShapeTextAutoFit(tb, 'normal', { fontScale: 0.72, lnSpcReduction: 0.18 });

    const entries = unzipSync(await savePresentation(pres));
    entries['ppt/slides/slide1.xml'] = strToU8(
      strFromU8(entries['ppt/slides/slide1.xml']!)
        .replace('fontScale="72000"', 'fontScale=" 72000 "')
        .replace('lnSpcReduction="18000"', 'lnSpcReduction=" 18% "'),
    );
    const loaded = await loadPresentation(zipSync(entries));
    const loadedShape = getSlideShapes(getSlides(loaded)[0]!).at(-1)!;
    expect(getShapeTextAutoFitParamsRaw(loadedShape)).toEqual({
      fontScale: 0.72,
      lnSpcReduction: 0,
    });
  });

  it('rejects invalid reduction parameters before mutating the current mode', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'A',
    });
    setShapeTextAutoFit(tb, 'normal', { fontScale: 0.72, lnSpcReduction: 0.18 });

    expect(() => setShapeTextAutoFit(tb, 'none', { fontScale: 0.5, lnSpcReduction: 0.1 })).toThrow(
      /only for normal mode/u,
    );
    for (const params of [
      { fontScale: -0.01, lnSpcReduction: 0.1 },
      { fontScale: 1.01, lnSpcReduction: 0.1 },
      { fontScale: 0.5, lnSpcReduction: Number.NaN },
    ]) {
      expect(() => setShapeTextAutoFit(tb, 'normal', params)).toThrow(
        /finite ratio from 0 through 1/u,
      );
    }
    expect(getShapeTextAutoFit(tb)).toBe('normal');
    expect(getShapeTextAutoFitParams(tb)).toEqual({ fontScale: 0.72, lnSpcReduction: 0.18 });
  });
});
