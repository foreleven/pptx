// Free-function per-run text editing API.
//
// Lets callers target a single paragraph + run by index, in contrast to
// `setShapeText` / `setShapeTextFormat` which rewrite or bulk-format
// every run.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { partName } from '../src/internal/opc/index.ts';
import {
  _internalPackageOf,
  getMediaParts,
  getShapeParagraphCount,
  getShapeRunCount,
  getShapeRunText,
  getSlideShapes,
  getSlidePartName,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeRunFormat,
  setShapeRunHyperlinkDescriptor,
  setShapeParagraphElements,
  setShapeRunText,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

describe('fn API: per-run text editing', () => {
  it('counts paragraphs and runs accurately after setShapeText', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeText(shape, 'first line\nsecond line');
    expect(getShapeParagraphCount(shape)).toBe(2);
    expect(getShapeRunCount(shape, 0)).toBe(1);
    expect(getShapeRunCount(shape, 1)).toBe(1);
    expect(getShapeRunText(shape, 0, 0)).toBe('first line');
    expect(getShapeRunText(shape, 1, 0)).toBe('second line');
  });

  it('setShapeRunFormat scopes formatting to one run only', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeText(shape, 'one\ntwo');
    setShapeRunFormat(shape, 0, 0, { bold: true, color: '#FF0000' });

    const xml = await slideXml(await savePresentation(pres), 0);
    // First run carries b="1" + the red srgbClr in its rPr.
    expect(xml).toMatch(/<a:rPr[^>]*b="1"/);
    expect(xml).toContain('FF0000');
    // The second run's <a:r> should NOT carry b="1" — easiest sanity
    // check is that the second `<a:t>two</a:t>` is preceded by a plain
    // <a:r> with no rPr.
    expect(xml).toMatch(/<a:r>\s*<a:t>two<\/a:t>/);
  });

  it('setShapeParagraphElements replaces one paragraph with ordered formatted runs', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeParagraphElements(shape, 0, [
      { kind: 'r', text: '3,309', format: { font: 'Arial', size: 32, bold: true } },
      {
        kind: 'r',
        text: ' 亿港元',
        format: { fontEastAsian: 'PingFang SC', size: 17, baseline: -0.08 },
      },
    ]);

    expect(getShapeRunCount(shape, 0)).toBe(2);
    expect(getShapeRunText(shape, 0, 0)).toBe('3,309');
    expect(getShapeRunText(shape, 0, 1)).toBe(' 亿港元');
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toMatch(/<a:rPr[^>]*sz="3200"[^>]*b="1"/);
    expect(xml).toContain('typeface="PingFang SC"');
    expect(xml).toContain('baseline="-8000"');
    expect(xml).not.toContain('xml:space');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedShape = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getShapeRunText(reloadedShape, 0, 0)).toBe('3,309');
    expect(getShapeRunText(reloadedShape, 0, 1)).toBe(' 亿港元');
  });

  it('setShapeParagraphElements collects relationships and sound parts owned by replaced runs', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeRunHyperlinkDescriptor(shape, 0, 0, 'click', {
      target: { kind: 'url', url: 'https://obsolete.example/' },
      sound: { data: new Uint8Array([1, 2, 3]), contentType: 'audio/wav' },
    });
    expect(getMediaParts(pres).filter((part) => part.contentType === 'audio/wav')).toHaveLength(1);

    setShapeParagraphElements(shape, 0, [{ kind: 'r', text: 'replacement' }]);

    expect(getMediaParts(pres).filter((part) => part.contentType === 'audio/wav')).toHaveLength(0);
    expect(
      _internalPackageOf(pres)
        .getRels(partName(getSlidePartName(slide)))
        ?.items.some((relationship) => relationship.target === 'https://obsolete.example/'),
    ).toBe(false);
  });

  it('setShapeRunText replaces visible characters but preserves rPr', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeText(shape, 'first\nsecond');
    setShapeRunFormat(shape, 0, 0, { italic: true });
    setShapeRunText(shape, 0, 0, 'PRIMARY');

    expect(getShapeRunText(shape, 0, 0)).toBe('PRIMARY');
    const xml = await slideXml(await savePresentation(pres), 0);
    // The italicized run still carries i="1" after the text swap.
    expect(xml).toMatch(/<a:rPr[^>]+i="1"[^>]*\/?>\s*<a:t>PRIMARY/);
  });

  it('throws RangeError on out-of-range paragraph or run', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeText(shape, 'only one para');
    expect(() => setShapeRunFormat(shape, 5, 0, { bold: true })).toThrow(RangeError);
    expect(() => setShapeRunText(shape, 0, 5, 'x')).toThrow(RangeError);
  });
});
