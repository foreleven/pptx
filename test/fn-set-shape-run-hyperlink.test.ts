// setShapeRunHyperlink — per-run hyperlink setter.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { partName } from '../src/internal/opc/index.ts';
import {
  _internalPackageOf,
  addSlideTextBox,
  getShapeRunHyperlink,
  getSlidePartName,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeRunHyperlink,
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
