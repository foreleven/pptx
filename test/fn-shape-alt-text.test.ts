// getShapeDescription / setShapeDescription / getShapeAltTitle /
// setShapeAltTitle — alt-text accessors for screen readers.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  addSlideTextBox,
  getShapeAltTitle,
  getShapeDescription,
  getShapeDecorative,
  inspectShapeDecorative,
  getShapeXmlString,
  getSlideShapes,
  getSlides,
  groupShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeAltTitle,
  setShapeDescription,
  setShapeDecorative,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: alt-text accessors', () => {
  it('round-trips a description', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const s = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    expect(getShapeDescription(s)).toBeNull();
    setShapeDescription(s, 'A red rectangle marking the focus area.');
    expect(getShapeDescription(s)).toBe('A red rectangle marking the focus area.');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const shapes = getSlideShapes(getSlides(reloaded)[0]!);
    expect(getShapeDescription(shapes[shapes.length - 1]!)).toBe(
      'A red rectangle marking the focus area.',
    );
  });

  it('round-trips an alt title', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const s = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'hi',
    });
    setShapeAltTitle(s, 'Greeting');
    expect(getShapeAltTitle(s)).toBe('Greeting');
  });

  it('passing null clears the attribute', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const s = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setShapeDescription(s, 'set');
    setShapeDescription(s, null);
    expect(getShapeDescription(s)).toBeNull();

    setShapeAltTitle(s, 'set');
    setShapeAltTitle(s, null);
    expect(getShapeAltTitle(s)).toBeNull();
  });

  it('preserves explicit empty alt-text attributes separately from missing attributes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });

    setShapeAltTitle(shape, '');
    setShapeDescription(shape, '');

    expect(getShapeAltTitle(shape)).toBe('');
    expect(getShapeDescription(shape)).toBe('');
    expect(getShapeXmlString(shape)).toContain(' title=""');
    expect(getShapeXmlString(shape)).toContain(' descr=""');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedShape = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
    expect(getShapeAltTitle(reloadedShape)).toBe('');
    expect(getShapeDescription(reloadedShape)).toBe('');
  });

  it('reads and writes accessibility metadata on group shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const first = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    const second = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    const group = groupShapes([first, second]);

    setShapeAltTitle(group, 'Metric group');
    setShapeDescription(group, 'Two metric tiles');
    setShapeDecorative(group, false);

    expect(getShapeAltTitle(group)).toBe('Metric group');
    expect(getShapeDescription(group)).toBe('Two metric tiles');
    expect(getShapeDecorative(group)).toBe(false);
  });

  it('round-trips the Office decorative accessibility extension', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(1),
      text: 'Decorative flourish',
    });

    expect(getShapeDecorative(shape)).toBeNull();
    setShapeDecorative(shape, true);
    expect(getShapeDecorative(shape)).toBe(true);
    expect(getShapeXmlString(shape)).toContain(
      '<a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext>',
    );
    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedShape = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
    expect(getShapeDecorative(reloadedShape)).toBe(true);
    setShapeDecorative(reloadedShape, false);
    expect(getShapeDecorative(reloadedShape)).toBe(false);
    setShapeDecorative(reloadedShape, null);
    expect(getShapeDecorative(reloadedShape)).toBeNull();
  });

  it('distinguishes absent, valid, malformed, and duplicate decorative metadata', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    expect(inspectShapeDecorative(shape)).toEqual({ kind: 'absent' });

    setShapeDecorative(shape, true);
    expect(inspectShapeDecorative(shape)).toEqual({ kind: 'valid', value: true });
    const entries = unzipSync(await savePresentation(pres));
    const slidePart = 'ppt/slides/slide1.xml';
    const validXml = strFromU8(entries[slidePart]!);

    entries[slidePart] = strToU8(validXml.replace('val="1"', 'val="maybe"'));
    const malformed = await loadPresentation(zipSync(entries));
    expect(inspectShapeDecorative(getSlideShapes(getSlides(malformed)[0]!).at(-1)!)).toEqual({
      kind: 'invalid',
      reason: 'invalid-value',
    });

    const extension = /<a:ext uri="\{C183D7F6-B498-43B3-948B-1728B52AA6E4\}">.*?<\/a:ext>/u.exec(
      validXml,
    )?.[0];
    if (!extension) throw new Error('Expected decorative extension in generated slide XML.');
    entries[slidePart] = strToU8(validXml.replace(extension, `${extension}${extension}`));
    const duplicate = await loadPresentation(zipSync(entries));
    expect(inspectShapeDecorative(getSlideShapes(getSlides(duplicate)[0]!).at(-1)!)).toEqual({
      kind: 'invalid',
      reason: 'duplicate-extension',
    });
  });
});
