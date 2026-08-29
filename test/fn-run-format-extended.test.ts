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
  getShapeRunFillImageBytes,
  getShapeRunFormat,
  getShapeRunUnderlineFillImageBytes,
  getShapeRunState,
  getSlidePartName,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeParagraphElements,
  setShapeRunFillImage,
  setShapeRunFormat,
  setShapeRunUnderlineFillImage,
} from '../src/api/index.ts';
import { applyLineStyle } from '../src/internal/drawingml/index.ts';
import { PATTERN_PRESETS } from '../src/internal/drawingml/fill.ts';
import { NS, attr, elem, getAttrValue, qname } from '../src/internal/xml/index.ts';
import { buildPng } from './lib/build-png.ts';

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

  it('round-trips complex underline paints and line geometry', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(5),
      h: inches(2),
      text: 'gradientpattern',
    });
    setShapeParagraphElements(tb, 0, [
      {
        kind: 'r',
        text: 'gradient',
        format: {
          underline: true,
          underlineFill: {
            kind: 'gradient',
            stops: [
              { offset: 0, color: '#3659E3' },
              { offset: 1, color: '#F26B5B' },
            ],
            angleDeg: 90,
          },
          underlineLine: {
            kind: 'bare',
            customDash: [
              { dash: 100_000, space: 25_000 },
              { dash: 15_000, space: 45_000 },
            ],
            head: { type: 'triangle', width: 'lg', length: 'lg' },
            tail: { type: 'oval', width: 'lg', length: 'lg' },
          },
        },
      },
      {
        kind: 'r',
        text: 'pattern',
        format: {
          underline: true,
          underlineFill: {
            kind: 'pattern',
            preset: 'pct20',
            foreground: '#3659E3',
            background: '#FFFFFF',
          },
        },
      },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const shape = findShapeByText(getSlides(reloaded)[0]!, 'gradientpattern')!;
    expect(getShapeRunFormat(shape, 0, 0)?.underlineFill).toEqual({
      kind: 'gradient',
      stops: [
        { offset: 0, color: '#3659E3' },
        { offset: 1, color: '#F26B5B' },
      ],
      angleDeg: 90,
    });
    expect(getShapeRunFormat(shape, 0, 0)?.underlineLine).toEqual({
      kind: 'bare',
      customDash: [
        { dash: 100_000, space: 25_000 },
        { dash: 15_000, space: 45_000 },
      ],
      head: { type: 'triangle', width: 'lg', length: 'lg' },
      tail: { type: 'oval', width: 'lg', length: 'lg' },
    });
    expect(getShapeRunFormat(shape, 0, 1)?.underlineFill).toEqual({
      kind: 'pattern',
      preset: 'pct20',
      foreground: '#3659E3',
      background: '#FFFFFF',
    });
  });

  it('embeds and resolves a picture underline fill relationship', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'picture',
    });
    const image = buildPng(2, 2, [54, 89, 227]);
    setShapeRunFormat(tb, 0, 0, { underline: true });
    setShapeRunUnderlineFillImage(tb, 0, 0, image);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const shape = findShapeByText(getSlides(reloaded)[0]!, 'picture')!;
    expect(getShapeRunFormat(shape, 0, 0)?.underlineFill).toEqual(
      expect.objectContaining({
        kind: 'picture',
        relationshipId: expect.stringMatching(/^rId\d+$/u),
      }),
    );
    expect(getShapeRunUnderlineFillImageBytes(shape, 0, 0)).toEqual(image);
  });

  it('round-trips an editable direct pattern text fill', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'pattern fill',
    });

    setShapeRunFormat(tb, 0, 0, {
      patternFill: {
        preset: 'pct50',
        foreground: '#3659E3',
        background: '#FFFFFF',
      },
    });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const shape = findShapeByText(getSlides(reloaded)[0]!, 'pattern fill')!;
    expect(getShapeRunFormat(shape, 0, 0)?.patternFill).toEqual({
      preset: 'pct50',
      foreground: '#3659E3',
      background: '#FFFFFF',
    });
  });

  it('round-trips every ST_PresetPatternVal as an editable direct text fill', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(8),
      h: inches(4),
      text: PATTERN_PRESETS.join(' '),
    });
    setShapeParagraphElements(
      tb,
      0,
      PATTERN_PRESETS.map((preset) => ({
        kind: 'r' as const,
        text: `${preset} `,
        format: { patternFill: { preset, foreground: '#3659E3', background: '#FFFFFF' } },
      })),
    );

    const reloaded = await loadPresentation(await savePresentation(pres));
    const shape = findShapeByText(getSlides(reloaded)[0]!, PATTERN_PRESETS[0])!;
    expect(PATTERN_PRESETS).toHaveLength(54);
    for (const [index, preset] of PATTERN_PRESETS.entries()) {
      expect(getShapeRunFormat(shape, 0, index)?.patternFill, preset).toEqual({
        preset,
        foreground: '#3659E3',
        background: '#FFFFFF',
      });
    }
  });

  it('rejects conflicting direct text fills before mutating the run', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'exclusive fill',
    });
    setShapeRunFormat(tb, 0, 0, { color: '#17233C' });

    expect(() =>
      setShapeRunFormat(tb, 0, 0, {
        color: '#17233C',
        patternFill: { preset: 'pct50', foreground: '#3659E3', background: '#FFFFFF' },
      }),
    ).toThrow(/at most one direct text fill/u);
    expect(getShapeRunFormat(tb, 0, 0)?.color).toBe('#17233C');
  });

  it('embeds and resolves a direct picture text fill relationship', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'picture fill',
    });
    const image = buildPng(2, 2, [54, 89, 227]);

    setShapeRunFillImage(tb, 0, 0, image);

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const shape = findShapeByText(getSlides(reloaded)[0]!, 'picture fill')!;
    expect(getShapeRunFormat(shape, 0, 0)?.pictureFill).toEqual({
      relationshipId: expect.stringMatching(/^rId\d+$/u),
    });
    expect(getShapeRunFillImageBytes(shape, 0, 0)).toEqual(image);
    const xml = getSlideXmlString(getSlides(reloaded)[0]!);
    expect(xml).toContain('<a:blipFill>');
    expect(xml).toContain('<a:stretch><a:fillRect/></a:stretch>');
    expect(xml).not.toContain('<a:tile');
  });

  it('keeps an imported picture relationship editable while diagnosing crop and tile normalization', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      text: 'complex picture fill',
    });
    const image = buildPng(2, 2, [242, 107, 91]);
    setShapeRunFillImage(tb, 0, 0, image);

    const pkg = _internalPackageOf(pres);
    const slidePartName = getSlidePartName(slide) as Parameters<typeof pkg.getPart>[0];
    const slidePart = pkg.getPart(slidePartName)!;
    const xml = new TextDecoder().decode(slidePart.data);
    const complex = xml.replace(
      '<a:stretch><a:fillRect/></a:stretch>',
      '<a:srcRect l="1000"/><a:tile tx="12700"/><a:stretch><a:fillRect l="3000"/></a:stretch>',
    );
    expect(complex).not.toBe(xml);
    slidePart.data = new TextEncoder().encode(complex);

    const imported = await loadPresentation(await savePresentation(pres));
    const shape = findShapeByText(getSlides(imported)[0]!, 'complex picture fill')!;
    expect(getShapeRunFormat(shape, 0, 0)?.pictureFill).toMatchObject({
      relationshipId: expect.stringMatching(/^rId\d+$/u),
      unsupported: expect.stringContaining('text picture fill child srcRect'),
    });
    expect(getShapeRunFormat(shape, 0, 0)?.pictureFill).toMatchObject({
      unsupported: expect.stringContaining('text picture fill fillRect attribute l'),
    });
    expect(getShapeRunFillImageBytes(shape, 0, 0)).toEqual(image);

    setShapeRunFillImage(shape, 0, 0, image);
    expect(getShapeRunFormat(shape, 0, 0)?.pictureFill).toEqual({
      relationshipId: expect.stringMatching(/^rId\d+$/u),
    });
  });

  it('round-trips gradient and pattern text outlines with custom dash and arrows', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(5),
      h: inches(2),
      text: 'gradientpattern',
    });
    setShapeParagraphElements(tb, 0, [
      {
        kind: 'r',
        text: 'gradient',
        format: {
          outline: {
            kind: 'gradient',
            stops: [
              { offset: 0, color: '#3659E3' },
              { offset: 1, color: '#F26B5B' },
            ],
            angleDeg: 90,
            customDash: [{ dash: 100_000, space: 25_000 }],
            head: { type: 'triangle', width: 'lg', length: 'lg' },
            tail: { type: 'oval', width: 'lg', length: 'lg' },
          },
        },
      },
      {
        kind: 'r',
        text: 'pattern',
        format: {
          outline: {
            kind: 'pattern',
            preset: 'pct20',
            foreground: '#3659E3',
            background: '#FFFFFF',
          },
        },
      },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const shape = findShapeByText(getSlides(reloaded)[0]!, 'gradientpattern')!;
    expect(getShapeRunFormat(shape, 0, 0)?.outline).toEqual({
      kind: 'gradient',
      stops: [
        { offset: 0, color: '#3659E3' },
        { offset: 1, color: '#F26B5B' },
      ],
      angleDeg: 90,
      customDash: [{ dash: 100_000, space: 25_000 }],
      head: { type: 'triangle', width: 'lg', length: 'lg' },
      tail: { type: 'oval', width: 'lg', length: 'lg' },
    });
    expect(getShapeRunFormat(shape, 0, 1)?.outline).toEqual({
      kind: 'pattern',
      preset: 'pct20',
      foreground: '#3659E3',
      background: '#FFFFFF',
    });
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
