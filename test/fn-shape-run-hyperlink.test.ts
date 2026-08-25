// getShapeRunHyperlink — per-run hyperlink lookup. Per-run setting
// goes through the existing setShapeHyperlink (applies to every run)
// for now; this reader makes it observable per run.

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
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeRunHyperlink', () => {
  const hyperlinkTargets = (
    pres: Parameters<typeof _internalPackageOf>[0],
    slide: Parameters<typeof getSlidePartName>[0],
  ): string[] =>
    (_internalPackageOf(pres).getRels(partName(getSlidePartName(slide)))?.items ?? [])
      .filter((relationship) => relationship.type.endsWith('/hyperlink'))
      .map((relationship) => relationship.target)
      .sort();

  it('returns the URL applied via setShapeHyperlink', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(1),
      text: 'click me',
    });
    setShapeHyperlink(tb, 'https://example.com/');
    expect(getShapeRunHyperlink(tb, 0, 0)).toBe('https://example.com/');
  });

  it('returns null on a run with no hyperlink', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(1),
      text: 'plain',
    });
    expect(getShapeRunHyperlink(tb, 0, 0)).toBeNull();
  });

  it('returns null after the hyperlink is cleared', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(1),
      text: 'toggle',
    });
    setShapeHyperlink(tb, 'https://example.com/');
    setShapeHyperlink(tb, null);
    expect(getShapeRunHyperlink(tb, 0, 0)).toBeNull();
    expect(hyperlinkTargets(pres, slide)).toEqual([]);
  });

  it('removes only unreachable relationships when replacing a shape-wide link', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const first = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(1),
      text: 'first',
    });
    const second = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(1),
      w: inches(3),
      h: inches(1),
      text: 'second',
    });
    setShapeHyperlink(first, 'https://shared.example/');
    setShapeHyperlink(second, 'https://shared.example/');

    setShapeHyperlink(first, 'https://replacement.example/');
    expect(hyperlinkTargets(pres, slide)).toEqual([
      'https://replacement.example/',
      'https://shared.example/',
    ]);
    setShapeHyperlink(second, null);
    expect(hyperlinkTargets(pres, slide)).toEqual(['https://replacement.example/']);
  });
});
