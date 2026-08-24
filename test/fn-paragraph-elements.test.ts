// `getShapeParagraphElements` — yields runs, field placeholders, and
// line breaks in document order so renderers can faithfully reproduce
// the paragraph's full visible content. The strict <a:r>-only
// `getShapeRunCount` would skip the fld / br children.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
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
});
