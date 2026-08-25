// `getShapeParagraphElements` — yields runs, field placeholders, and
// line breaks in document order so renderers can faithfully reproduce
// the paragraph's full visible content. The strict <a:r>-only
// `getShapeRunCount` would skip the fld / br children.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeParagraphElements,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeParagraphElements,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeParagraphElements', () => {
  it('emits a single run for a one-word text box', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'hi',
    });
    const els = getShapeParagraphElements(tb, 0);
    expect(els).toHaveLength(1);
    expect(els[0]!.kind).toBe('r');
    if (els[0]!.kind === 'r') expect(els[0]!.text).toBe('hi');
  });

  it('returns an empty array for an out-of-range paragraph index', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'hi',
    });
    expect(() => getShapeParagraphElements(tb, 99)).toThrow();
  });

  it('writes runs and native line breaks in document order', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'FirstSecond',
    });

    setShapeParagraphElements(tb, 0, [
      { kind: 'r', text: 'First', format: { bold: true } },
      { kind: 'br' },
      { kind: 'r', text: 'Second' },
    ]);

    expect(getShapeParagraphElements(tb, 0)).toEqual([
      { kind: 'r', text: 'First', format: { bold: true } },
      { kind: 'br', format: null },
      { kind: 'r', text: 'Second', format: null },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedShape = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
    expect(getShapeParagraphElements(reloadedShape, 0).map((element) => element.kind)).toEqual([
      'r',
      'br',
      'r',
    ]);
  });

  it('writes and reloads editable fields with direct run and paragraph properties', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: '7Q1',
    });

    setShapeParagraphElements(tb, 0, [
      {
        kind: 'fld',
        id: '{aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee}',
        type: 'futurePowerPointToken',
        text: '7',
        format: { bold: true, color: '#3659E3' },
        state: { normalizeHeight: true, noProof: true },
        paragraph: {
          align: 'justLow',
          fontAlign: 'baseline',
          level: 3,
          marL: inches(0.25),
          marR: inches(0.125),
          indent: -inches(0.125),
          lineSpacing: { kind: 'pct', value: 1.25 },
          spcBefPts: 6,
          spcAftPts: 4,
          rtl: true,
          defaultTabSize: inches(0.5),
          tabStops: [{ positionEmu: inches(1), alignment: 'decimal' }],
        },
      },
      {
        kind: 'fld',
        id: '{11111111-2222-3333-4444-555555555555}',
        text: 'Q1',
      },
    ]);

    const expected = [
      {
        kind: 'fld',
        id: '{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}',
        type: 'futurePowerPointToken',
        text: '7',
        format: { bold: true, color: '#3659E3' },
        state: { normalizeHeight: true, noProof: true },
        paragraph: {
          align: 'justLow',
          fontAlign: 'baseline',
          level: 3,
          marL: inches(0.25),
          marR: inches(0.125),
          indent: -inches(0.125),
          lineSpacing: { kind: 'pct', value: 1.25 },
          spcBefPts: 6,
          spcAftPts: 4,
          rtl: true,
          defaultTabSize: inches(0.5),
          tabStops: [{ positionEmu: inches(1), alignment: 'decimal' }],
        },
      },
      {
        kind: 'fld',
        id: '{11111111-2222-3333-4444-555555555555}',
        type: null,
        text: 'Q1',
        format: null,
        state: null,
        paragraph: null,
      },
    ];
    expect(getShapeParagraphElements(tb, 0)).toEqual(expected);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedShape = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
    expect(getShapeParagraphElements(reloadedShape, 0)).toEqual(expected);
  });

  it('rejects invalid field identity, type, and paragraph ranges', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: '7',
    });
    const field = {
      kind: 'fld',
      id: '{11111111-2222-3333-4444-555555555555}',
      type: 'slidenum',
      text: '7',
    } as const;

    expect(() => setShapeParagraphElements(tb, 0, [{ ...field, id: 'not-a-guid' }])).toThrow(
      'text field id: "not-a-guid" is not a GUID',
    );
    expect(() =>
      setShapeParagraphElements(tb, 0, [{ ...field, type: 7 as unknown as string }]),
    ).toThrow('text field type must be a string');
    expect(() => setShapeParagraphElements(tb, 0, [{ ...field, paragraph: { level: 9 } }])).toThrow(
      'text field paragraph level must be an integer in [0, 8]',
    );
    expect(() =>
      setShapeParagraphElements(tb, 0, [{ ...field, paragraph: { spcBefPts: -1 } }]),
    ).toThrow('text field paragraph spcBef must be a non-negative number');
    expect(() =>
      setShapeParagraphElements(tb, 0, [
        {
          ...field,
          paragraph: {
            tabStops: Array.from({ length: 33 }, (_, index) => ({
              positionEmu: index,
              alignment: 'left' as const,
            })),
          },
        },
      ]),
    ).toThrow('text field paragraph tab stops must contain at most 32 entries');
  });

  it('reports unsupported field and field-paragraph semantics to editable importers', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: '7',
    });
    setShapeParagraphElements(tb, 0, [
      {
        kind: 'fld',
        id: '{11111111-2222-3333-4444-555555555555}',
        type: 'slidenum',
        text: '7',
        paragraph: { align: 'left', lineSpacing: { kind: 'pts', value: 18 } },
      },
    ]);

    const entries = unzipSync(await savePresentation(pres));
    const slidePart = 'ppt/slides/slide1.xml';
    const original = strFromU8(entries[slidePart]!);
    const changed = original
      .replace(
        'type="slidenum"',
        'type="slidenum" xmlns:future="urn:office-kit:test" future:opaque="1"',
      )
      .replace('algn="l"', 'algn="l" hangingPunct="1"')
      .replace('</a:pPr>', '<a:buNone/></a:pPr>')
      .replace('</a:fld>', '<future:payload/></a:fld>');
    expect(changed).not.toBe(original);
    entries[slidePart] = strToU8(changed);

    const reloaded = await loadPresentation(zipSync(entries));
    const reloadedShape = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
    const field = getShapeParagraphElements(reloadedShape, 0)[0];
    expect(field).toEqual(
      expect.objectContaining({
        kind: 'fld',
        unsupported: expect.arrayContaining([
          'text field attribute opaque',
          'text field child payload',
          'text field paragraph attribute hangingPunct',
          'text field paragraph child buNone',
        ]),
      }),
    );
  });
});
