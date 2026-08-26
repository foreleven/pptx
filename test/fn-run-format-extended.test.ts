// Extended run-format attributes: strike, spc (character spacing),
// kern (kerning threshold), baseline (super / sub), cap, highlight.
// These are part of ECMA-376's CT_TextCharacterProperties surface
// (§17.18.83) and round-trip through `setShapeRunFormat` /
// `getShapeRunFormat`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlideTextBox,
  findShapeByText,
  getShapeRunFormat,
  getShapeRunState,
  getSlidePartName,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeParagraphElements,
  setShapeRunFormat,
} from '../src/api/index.ts';
import { applyLineStyle } from '../src/internal/drawingml/index.ts';
import { NS, attr, elem, getAttrValue, qname } from '../src/internal/xml/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: extended run-format properties', () => {
  it('updates an unqualified line width without deleting same-local-name extension attributes', () => {
    const line = elem(qname('a', 'ln', NS.dml), {
      attrs: [
        attr(qname('', 'w', ''), '12700'),
        attr(qname('future', 'w', 'urn:office-kit:test'), 'opaque'),
      ],
    });

    applyLineStyle(line, { widthEmu: 25_400 });

    expect(getAttrValue(line, qname('', 'w', ''))).toBe('25400');
    expect(getAttrValue(line, qname('future', 'w', 'urn:office-kit:test'))).toBe('opaque');
  });

  it('accepts every ST_TextUnderlineType token and rejects unknown values', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'underline',
    });
    const tokens = [
      'none',
      'words',
      'sng',
      'dbl',
      'heavy',
      'dotted',
      'dottedHeavy',
      'dash',
      'dashHeavy',
      'dashLong',
      'dashLongHeavy',
      'dotDash',
      'dotDashHeavy',
      'dotDotDash',
      'dotDotDashHeavy',
      'wavy',
      'wavyHeavy',
      'wavyDbl',
    ] as const;
    for (const token of tokens) {
      setShapeRunFormat(tb, 0, 0, { underline: token });
      expect(getShapeRunFormat(tb, 0, 0)?.underline).toBe(
        token === 'none' ? false : token === 'sng' ? true : token,
      );
    }
    expect(() => setShapeRunFormat(tb, 0, 0, { underline: 'single' })).toThrow(
      /valid ST_TextUnderlineType token/,
    );
  });

  it('round-trips editable underline line and fill choices', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'follow explicit none',
    });
    setShapeParagraphElements(tb, 0, [
      {
        kind: 'r',
        text: 'follow ',
        format: {
          underline: 'words',
          underlineLine: { kind: 'followText' },
          underlineFill: { kind: 'followText' },
        },
      },
      {
        kind: 'r',
        text: 'explicit ',
        format: {
          underline: 'dashHeavy',
          underlineLine: {
            kind: 'solid',
            color: '#3659E3',
            widthPt: 1.25,
            cap: 'rnd',
            dash: 'dashDot',
            join: 'bevel',
            compound: 'dbl',
            alignment: 'ctr',
          },
          underlineFill: { kind: 'solid', color: '#F26B5B80' },
        },
      },
      {
        kind: 'r',
        text: 'none',
        format: { underline: 'wavyDbl', underlineFill: { kind: 'none' } },
      },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedShape = findShapeByText(getSlides(reloaded)[0]!, 'follow explicit none')!;
    expect(getShapeRunFormat(reloadedShape, 0, 0)).toMatchObject({
      underline: 'words',
      underlineLine: { kind: 'followText' },
      underlineFill: { kind: 'followText' },
    });
    expect(getShapeRunFormat(reloadedShape, 0, 1)).toMatchObject({
      underline: 'dashHeavy',
      underlineLine: {
        kind: 'solid',
        color: '#3659E3',
        widthPt: 1.25,
        cap: 'rnd',
        dash: 'dashDot',
        join: 'bevel',
        compound: 'dbl',
        alignment: 'ctr',
      },
      underlineFill: { kind: 'solid', color: '#F26B5B80' },
    });
    expect(getShapeRunFormat(reloadedShape, 0, 2)).toMatchObject({
      underline: 'wavyDbl',
      underlineFill: { kind: 'none' },
    });
  });

  it('round-trips strike, spc, kern, baseline, cap, highlight', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'extended',
    });
    setShapeRunFormat(tb, 0, 0, {
      strike: true,
      spc: 200,
      kern: 1200,
      baseline: 0.3,
      cap: 'all',
      highlight: '#FFFF00',
    });
    const fmt = getShapeRunFormat(tb, 0, 0);
    expect(fmt).not.toBeNull();
    expect(fmt!.strike).toBe(true);
    expect(fmt!.spc).toBe(200);
    expect(fmt!.kern).toBe(1200);
    expect(fmt!.baseline).toBeCloseTo(0.3, 4);
    expect(fmt!.cap).toBe('all');
    expect(fmt!.highlight).toBe('#FFFF00');
  });

  it('strike accepts both boolean shorthand and explicit dblStrike', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 's',
    });
    setShapeRunFormat(tb, 0, 0, { strike: true });
    expect(getShapeRunFormat(tb, 0, 0)!.strike).toBe(true);
    setShapeRunFormat(tb, 0, 0, { strike: 'dblStrike' });
    expect(getShapeRunFormat(tb, 0, 0)!.strike).toBe('dblStrike');
    setShapeRunFormat(tb, 0, 0, { strike: false });
    expect(getShapeRunFormat(tb, 0, 0)!.strike).toBe(false);
  });

  it('round-trips solid and explicit no text outlines', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'solidnone',
    });
    setShapeParagraphElements(tb, 0, [
      {
        kind: 'r',
        text: 'solid',
        format: {
          outline: {
            kind: 'solid',
            color: '#3659E3',
            widthPt: 1.5,
            cap: 'sq',
            dash: 'lgDashDot',
            join: 'round',
            compound: 'thickThin',
            alignment: 'in',
          },
        },
      },
      { kind: 'r', text: 'none', format: { outline: { kind: 'none' } } },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedShape = findShapeByText(getSlides(reloaded)[0]!, 'solidnone')!;
    expect(getShapeRunFormat(reloadedShape, 0, 0)?.outline).toEqual({
      kind: 'solid',
      color: '#3659E3',
      widthPt: 1.5,
      cap: 'sq',
      dash: 'lgDashDot',
      join: 'round',
      compound: 'thickThin',
      alignment: 'in',
    });
    expect(getShapeRunFormat(reloadedShape, 0, 1)?.outline).toEqual({ kind: 'none' });
  });

  it('rejects a kerning threshold outside ST_TextNonNegativePoint', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'kern',
    });
    expect(() => setShapeRunFormat(tb, 0, 0, { kern: 400001 })).toThrow(
      /out of range for textNonNegativePoint \(0\.\.400000\)/,
    );
  });

  it('reads direct normalize, proofing, dirty, error, and smart-tag run state', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'diagnostic state',
    });

    const pkg = _internalPackageOf(pres);
    const slidePartName = getSlidePartName(slide) as Parameters<typeof pkg.getPart>[0];
    const slidePart = pkg.getPart(slidePartName)!;
    const xml = new TextDecoder().decode(slidePart.data);
    const withState = xml.replace(
      /<a:r>(?:<a:rPr[^>]*\/>)?<a:t>diagnostic state<\/a:t><\/a:r>/u,
      '<a:r><a:rPr normalizeH="1" noProof="1" dirty="1" err="1" smtClean="1" smtId="42"/><a:t>diagnostic state</a:t></a:r>',
    );
    expect(withState).not.toBe(xml);
    slidePart.data = new TextEncoder().encode(withState);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedShape = findShapeByText(getSlides(reloaded)[0]!, 'diagnostic state')!;
    expect(getShapeRunState(reloadedShape, 0, 0)).toEqual({
      normalizeHeight: true,
      noProof: true,
      dirty: true,
      error: true,
      smartTagClean: true,
      smartTagId: 42,
    });
  });

  it('round-trips authored direct run state while preserving explicit false values', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'authored state',
    });
    setShapeParagraphElements(tb, 0, [
      {
        kind: 'r',
        text: 'authored state',
        state: {
          normalizeHeight: true,
          noProof: false,
          dirty: false,
          error: true,
          smartTagClean: false,
          smartTagId: 42,
        },
      },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedShape = findShapeByText(getSlides(reloaded)[0]!, 'authored state')!;
    expect(getShapeRunState(reloadedShape, 0, 0)).toEqual({
      normalizeHeight: true,
      noProof: false,
      dirty: false,
      error: true,
      smartTagClean: false,
      smartTagId: 42,
    });
  });
});
