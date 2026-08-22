import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideChart,
  addSlideComment,
  addSlideImage,
  addSlideMedia,
  addSlideOpaqueObject,
  addSlideShape,
  createPresentation,
  extractOpaqueObjectFragment,
  getCommentText,
  getShapeClickAction,
  getShapeAnimation,
  getShapeChartSpec,
  getShapeImagePartName,
  getShapeMedia,
  getSlideComments,
  getSlideBackground,
  getSlideNotes,
  getSlidePartName,
  getSlideSections,
  getSlideShapes,
  getSlideText,
  getSlideTopLevelShapes,
  getSlideTransition,
  getSlides,
  inches,
  listPackageParts,
  loadPresentation,
  removeSlide,
  replaceSlideContents,
  savePresentation,
  setShapeClickAction,
  setShapeAnimation,
  setSlideBackground,
  setSlideNotes,
  setSlideSections,
  setSlideTransition,
  validatePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe('replaceSlideContents', () => {
  it('replaces the target body while preserving its slide identity and order', async () => {
    const presentation = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [target, replacement] = getSlides(presentation);
    if (!target || !replacement) throw new Error('expected two slides');

    const targetPartName = getSlidePartName(target);
    const replacementPartName = getSlidePartName(replacement);
    const replacementText = getSlideText(replacement);

    replaceSlideContents(target, replacement);

    expect(getSlidePartName(target)).toBe(targetPartName);
    expect(getSlideText(target)).toBe(replacementText);
    expect(getSlides(presentation).map(getSlidePartName)).toEqual([
      targetPartName,
      replacementPartName,
    ]);

    const reloaded = await loadPresentation(await savePresentation(presentation));
    expect(getSlides(reloaded).map(getSlidePartName)).toEqual([
      targetPartName,
      replacementPartName,
    ]);
    expect(getSlideText(getSlides(reloaded)[0]!)).toBe(replacementText);
  });

  it('adopts replacement relationships and removes parts owned only by the old contents', async () => {
    const presentation = createPresentation();
    const target = addBlankSlide(presentation);
    const replacement = addBlankSlide(presentation);
    const targetPartName = getSlidePartName(target);

    const oldPicture = addSlideImage(target, PNG, {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    setSlideNotes(target, 'old notes');
    addSlideComment(target, { author: { name: 'Old reviewer' }, text: 'old comment' });
    const oldOwnedParts = listPackageParts(presentation)
      .map((part) => part.name)
      .filter(
        (name) =>
          name.startsWith('/ppt/media/') ||
          name.startsWith('/ppt/notesSlides/') ||
          name.startsWith('/ppt/comments/'),
      );

    const replacementPicture = addSlideImage(replacement, PNG, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    setSlideNotes(replacement, 'new notes');
    addSlideComment(replacement, { author: { name: 'New reviewer' }, text: 'new comment' });
    const replacementImagePart = getShapeImagePartName(replacementPicture);
    expect(replacementImagePart).not.toBeNull();

    replaceSlideContents(target, replacement);
    removeSlide(presentation, replacement);

    const reloaded = await loadPresentation(await savePresentation(presentation));
    const [result] = getSlides(reloaded);
    if (!result) throw new Error('expected target slide');
    expect(getSlides(reloaded)).toHaveLength(1);
    expect(getSlidePartName(result)).toBe(targetPartName);
    expect(getSlideNotes(result)).toBe('new notes');
    expect(getSlideComments(result).map(getCommentText)).toEqual(['new comment']);
    expect(getShapeImagePartName(getSlideShapes(result)[0]!)).toBe(replacementImagePart);

    const remainingNames = new Set(listPackageParts(reloaded).map((part) => part.name));
    for (const oldOwnedPart of oldOwnedParts) expect(remainingNames.has(oldOwnedPart)).toBe(false);
    void oldPicture;
  });

  it('keeps sections and incoming slide links while adopting replacement slide links', async () => {
    const presentation = createPresentation();
    const target = addBlankSlide(presentation);
    const destination = addBlankSlide(presentation);
    const replacement = addBlankSlide(presentation);
    const incomingLink = addSlideShape(destination, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'Back to target',
    });
    setShapeClickAction(incomingLink, { kind: 'slide', slide: target });
    const outgoingLink = addSlideShape(replacement, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'Continue',
    });
    setShapeClickAction(outgoingLink, { kind: 'slide', slide: destination });
    setSlideSections(presentation, [{ name: 'Main', slides: [target, destination] }]);

    replaceSlideContents(target, replacement);
    removeSlide(presentation, replacement);

    const reloaded = await loadPresentation(await savePresentation(presentation));
    const [result, reloadedDestination] = getSlides(reloaded);
    if (!result || !reloadedDestination) throw new Error('expected target and destination');
    expect(
      getSlideSections(reloaded).map((section) => ({
        name: section.name,
        slides: section.slides.map(getSlidePartName),
      })),
    ).toEqual([
      { name: 'Main', slides: [getSlidePartName(result), getSlidePartName(reloadedDestination)] },
    ]);

    const incoming = getShapeClickAction(getSlideShapes(reloadedDestination)[0]!);
    expect(incoming?.kind).toBe('slide');
    if (incoming?.kind === 'slide')
      expect(getSlidePartName(incoming.slide)).toBe(getSlidePartName(result));

    const outgoing = getShapeClickAction(getSlideShapes(result)[0]!);
    expect(outgoing?.kind).toBe('slide');
    if (outgoing?.kind === 'slide') {
      expect(getSlidePartName(outgoing.slide)).toBe(getSlidePartName(reloadedDestination));
    }
  });

  it('preserves charts, media, opaque objects, backgrounds, transitions, and animations', async () => {
    const fragmentSource = createPresentation();
    const fragmentSlide = addBlankSlide(fragmentSource);
    const fragmentChart = addSlideChart(fragmentSlide, {
      spec: {
        kind: 'line',
        categories: ['Q1', 'Q2'],
        series: [{ name: 'Trend', values: [4, 9] }],
      },
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
    });
    const opaqueFragment = extractOpaqueObjectFragment(fragmentChart);

    const presentation = createPresentation();
    const target = addBlankSlide(presentation);
    const replacement = addBlankSlide(presentation);
    const animated = addSlideShape(replacement, {
      preset: 'ellipse',
      x: inches(0.5),
      y: inches(0.5),
      w: inches(1),
      h: inches(1),
    });
    setShapeAnimation(animated, { effect: 'fadeIn' });
    addSlideChart(replacement, {
      spec: {
        kind: 'column',
        categories: ['North', 'South'],
        series: [{ name: 'Revenue', values: [12, 18] }],
      },
      x: inches(2),
      y: inches(0.5),
      w: inches(4),
      h: inches(2.5),
    });
    addSlideMedia(replacement, Uint8Array.from([0, 0, 0, 24, 102, 116, 121, 112]), {
      kind: 'video',
      contentType: 'video/mp4',
      posterBytes: PNG,
      posterFormat: 'png',
      x: inches(0.5),
      y: inches(3.5),
      w: inches(3),
      h: inches(2),
    });
    addSlideOpaqueObject(replacement, opaqueFragment);
    setSlideBackground(replacement, '#123456');
    setSlideTransition(replacement, { effect: 'fade', speed: 'slow' });

    replaceSlideContents(target, replacement);
    removeSlide(presentation, replacement);

    const reloaded = await loadPresentation(await savePresentation(presentation));
    const result = getSlides(reloaded)[0]!;
    const shapes = getSlideTopLevelShapes(result);
    expect(getShapeAnimation(shapes[0]!)).toBe('fadeIn');
    expect(getShapeChartSpec(shapes[1]!)).toMatchObject({
      kind: 'column',
      series: [{ name: 'Revenue', values: [12, 18] }],
    });
    expect(getShapeMedia(shapes[2]!)).toMatchObject({
      kind: 'video',
      contentType: 'video/mp4',
    });
    expect(getShapeChartSpec(shapes[3]!)).toMatchObject({
      kind: 'line',
      series: [{ name: 'Trend', values: [4, 9] }],
    });
    expect(getSlideBackground(result)).toMatchObject({ color: '#123456' });
    expect(getSlideTransition(result)).toEqual({ effect: 'fade', speed: 'slow' });
    expect(validatePresentation(reloaded)).toEqual([]);
  });
});
