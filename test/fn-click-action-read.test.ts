// getShapeClickAction — read-back parity for setShapeClickAction.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { partName } from '../src/internal/opc/index.ts';
import {
  _internalPackageOf,
  getShapeClickAction,
  getSlidePartName,
  getSlideShapes,
  getSlides,
  loadPresentation,
  setShapeClickAction,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const actionRelationshipTargets = (
  pres: Parameters<typeof _internalPackageOf>[0],
  slide: Parameters<typeof getSlidePartName>[0],
): string[] =>
  (_internalPackageOf(pres).getRels(partName(getSlidePartName(slide)))?.items ?? [])
    .filter(
      (relationship) =>
        relationship.type.endsWith('/hyperlink') || relationship.type.endsWith('/slide'),
    )
    .map((relationship) => relationship.target)
    .sort();

describe('fn API: getShapeClickAction', () => {
  it('returns null when no click action is set', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    expect(getShapeClickAction(shape)).toBeNull();
  });

  it('round-trips a URL action', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeClickAction(shape, { kind: 'url', url: 'https://example.com/' });
    expect(getShapeClickAction(shape)).toEqual({
      kind: 'url',
      url: 'https://example.com/',
    });
  });

  it('round-trips a slide jump', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    const target = slides[1]!;
    const shape = getSlideShapes(slides[0]!)[0]!;
    setShapeClickAction(shape, { kind: 'slide', slide: target });
    const got = getShapeClickAction(shape);
    expect(got?.kind).toBe('slide');
    // Identity compares structurally — slide opaque handle is rebuilt
    // on read, so compare on a stable text proxy.
    if (got?.kind === 'slide') {
      const { getSlideText } = await import('../src/api/index.ts');
      expect(getSlideText(got.slide)).toBe(getSlideText(target));
    }
  });

  it('round-trips preset navigation actions', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    for (const k of ['nextSlide', 'prevSlide', 'firstSlide', 'lastSlide'] as const) {
      setShapeClickAction(shape, { kind: k });
      expect(getShapeClickAction(shape)).toEqual({ kind: k });
    }
  });

  it('removes relationships made unreachable by action replacement and clearing', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    const slide = slides[0]!;
    const shape = getSlideShapes(slide)[0]!;

    setShapeClickAction(shape, { kind: 'url', url: 'https://action.example/' });
    expect(actionRelationshipTargets(pres, slide)).toContain('https://action.example/');

    setShapeClickAction(shape, { kind: 'slide', slide: slides[1]! });
    expect(actionRelationshipTargets(pres, slide)).not.toContain('https://action.example/');
    expect(actionRelationshipTargets(pres, slide)).toContain('../slides/slide2.xml');

    setShapeClickAction(shape, { kind: 'nextSlide' });
    expect(actionRelationshipTargets(pres, slide)).not.toContain('../slides/slide2.xml');

    setShapeClickAction(shape, null);
    expect(actionRelationshipTargets(pres, slide)).not.toContain('https://action.example/');
  });
});
