// setShapeRunHyperlink — per-run hyperlink setter.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { partName } from '../src/internal/opc/index.ts';
import {
  _internalPackageOf,
  addSlideTextBox,
  getShapeRunHyperlinkDescriptor,
  getShapeRunHyperlink,
  getMediaParts,
  getSlidePartName,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeRunHyperlink,
  setShapeRunHyperlinkDescriptor,
  setShapeRunText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const hyperlinkTargets = (
  pres: Parameters<typeof _internalPackageOf>[0],
  slide: Parameters<typeof getSlidePartName>[0],
): string[] =>
  (_internalPackageOf(pres).getRels(partName(getSlidePartName(slide)))?.items ?? [])
    .filter((relationship) => relationship.type.endsWith('/hyperlink'))
    .map((relationship) => relationship.target)
    .sort();

describe('fn API: setShapeRunHyperlink', () => {
  it('writes the empty relationship ID PowerPoint requires for action-only interactions', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(1),
      text: 'action only',
    });

    setShapeRunHyperlinkDescriptor(tb, 0, 0, 'click', {
      action: 'ppaction://hlinkshowjump?jump=nextslide',
    });

    const entries = unzipSync(await savePresentation(pres));
    const slideXml = strFromU8(entries['ppt/slides/slide1.xml']!);
    expect(slideXml).toContain(
      '<a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=nextslide"/>',
    );
  });

  it('round-trips complete click and hover hyperlink descriptors', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const targetSlide = getSlides(pres)[1]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(1),
      text: 'click\nhover',
    });

    setShapeRunHyperlinkDescriptor(tb, 0, 0, 'click', {
      target: { kind: 'url', url: 'https://example.com/click' },
      action: 'ppaction://program',
      invalidUrl: 'https://invalid.example/',
      targetFrame: '_blank',
      tooltip: 'Click tip',
      history: false,
      highlightClick: true,
      endSound: true,
      sound: {
        data: new Uint8Array([82, 73, 70, 70, 1, 2, 3, 4]),
        contentType: 'audio/wav',
        name: 'Click sound',
      },
    });
    setShapeRunHyperlinkDescriptor(tb, 1, 0, 'hover', {
      target: { kind: 'slide', slide: targetSlide },
      action: 'ppaction://hlinksldjump',
      tooltip: 'Hover tip',
    });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const shapes = (await import('../src/api/index.ts')).getSlideShapes(getSlides(reloaded)[0]!);
    const rebuilt = shapes[shapes.length - 1]!;
    expect(getShapeRunHyperlinkDescriptor(rebuilt, 0, 0, 'click')).toEqual({
      target: { kind: 'url', url: 'https://example.com/click' },
      action: 'ppaction://program',
      invalidUrl: 'https://invalid.example/',
      targetFrame: '_blank',
      tooltip: 'Click tip',
      history: false,
      highlightClick: true,
      endSound: true,
      sound: {
        data: new Uint8Array([82, 73, 70, 70, 1, 2, 3, 4]),
        contentType: 'audio/wav',
        name: 'Click sound',
      },
    });
    const hover = getShapeRunHyperlinkDescriptor(rebuilt, 1, 0, 'hover');
    expect(hover?.target?.kind).toBe('slide');
    expect(hover?.target?.kind === 'slide' ? getSlidePartName(hover.target.slide) : null).toBe(
      getSlidePartName(getSlides(reloaded)[1]!),
    );
    expect(hover?.action).toBe('ppaction://hlinksldjump');
    expect(hover?.tooltip).toBe('Hover tip');
  });

  it('replaces or clears one activation without disturbing the other', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(1),
      text: 'both',
    });

    setShapeRunHyperlinkDescriptor(tb, 0, 0, 'click', {
      target: { kind: 'url', url: 'https://click.example/' },
    });
    setShapeRunHyperlinkDescriptor(tb, 0, 0, 'hover', {
      target: { kind: 'url', url: 'https://hover.example/' },
      action: 'vendor://unknown-action',
    });
    setShapeRunHyperlinkDescriptor(tb, 0, 0, 'click', null);

    expect(getShapeRunHyperlinkDescriptor(tb, 0, 0, 'click')).toBeNull();
    expect(getShapeRunHyperlinkDescriptor(tb, 0, 0, 'hover')).toEqual({
      target: { kind: 'url', url: 'https://hover.example/' },
      action: 'vendor://unknown-action',
    });
    expect(hyperlinkTargets(pres, slide)).toEqual(['https://hover.example/']);
  });

  it('garbage-collects sound relationships and owned media parts on replacement and clear', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(1),
      text: 'sound',
    });
    setShapeRunHyperlinkDescriptor(tb, 0, 0, 'click', {
      action: 'ppaction://noaction',
      sound: { data: new Uint8Array([1, 2, 3]), contentType: 'audio/wav' },
    });
    expect(getMediaParts(pres).filter((part) => part.contentType === 'audio/wav')).toHaveLength(1);

    setShapeRunHyperlinkDescriptor(tb, 0, 0, 'click', {
      action: 'ppaction://noaction',
      sound: { data: new Uint8Array([4, 5, 6]), contentType: 'audio/wav' },
    });
    const audio = getMediaParts(pres).filter((part) => part.contentType === 'audio/wav');
    expect(audio).toHaveLength(1);
    expect(audio[0]?.data).toEqual(new Uint8Array([4, 5, 6]));

    setShapeRunHyperlinkDescriptor(tb, 0, 0, 'click', null);
    expect(getMediaParts(pres).filter((part) => part.contentType === 'audio/wav')).toHaveLength(0);
  });

  it('links one run and leaves others untouched', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(1),
      // Two-line text gives us two paragraphs / two distinct runs.
      text: 'first\nsecond',
    });

    setShapeRunHyperlink(tb, 0, 0, 'https://first.example/');
    expect(getShapeRunHyperlink(tb, 0, 0)).toBe('https://first.example/');
    // Different paragraph, different run — should stay unlinked.
    expect(getShapeRunHyperlink(tb, 1, 0)).toBeNull();
  });

  it('round-trips through save / reload', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(1),
      text: 'click',
    });
    setShapeRunHyperlink(tb, 0, 0, 'https://example.com/');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const slide2 = getSlides(reloaded)[0]!;
    // Find the textbox we added — the last shape on the slide.
    const { getSlideShapes } = await import('../src/api/index.ts');
    const shapes = getSlideShapes(slide2);
    const last = shapes[shapes.length - 1]!;
    expect(getShapeRunHyperlink(last, 0, 0)).toBe('https://example.com/');
  });

  it('passing null clears the link', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(1),
      text: 'click',
    });
    setShapeRunHyperlink(tb, 0, 0, 'https://example.com/');
    expect(getShapeRunHyperlink(tb, 0, 0)).toBe('https://example.com/');
    setShapeRunHyperlink(tb, 0, 0, null);
    expect(getShapeRunHyperlink(tb, 0, 0)).toBeNull();
  });

  it('reuses an existing rel when the same URL is set on another run', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(1),
      text: 'a\nb',
    });
    // Sanity: ensure both runs exist.
    setShapeRunText(tb, 0, 0, 'a');
    setShapeRunText(tb, 1, 0, 'b');

    setShapeRunHyperlink(tb, 0, 0, 'https://shared.example/');
    setShapeRunHyperlink(tb, 1, 0, 'https://shared.example/');
    expect(getShapeRunHyperlink(tb, 0, 0)).toBe('https://shared.example/');
    expect(getShapeRunHyperlink(tb, 1, 0)).toBe('https://shared.example/');
  });

  it('removes replaced and cleared hyperlink relationships without deleting shared references', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(1),
      text: 'a\nb',
    });

    setShapeRunHyperlink(tb, 0, 0, 'https://shared.example/');
    setShapeRunHyperlink(tb, 1, 0, 'https://shared.example/');
    expect(hyperlinkTargets(pres, slide)).toContain('https://shared.example/');

    setShapeRunHyperlink(tb, 0, 0, 'https://replacement.example/');
    expect(hyperlinkTargets(pres, slide)).toEqual([
      'https://replacement.example/',
      'https://shared.example/',
    ]);

    setShapeRunHyperlink(tb, 1, 0, null);
    expect(hyperlinkTargets(pres, slide)).toEqual(['https://replacement.example/']);

    setShapeRunHyperlink(tb, 0, 0, null);
    expect(hyperlinkTargets(pres, slide)).toEqual([]);
  });
});
