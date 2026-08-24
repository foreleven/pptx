// Slide sections (p14:sectionLst).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  readPackagePart,
  getSlideSections,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideSections,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: slide sections', () => {
  it('reports empty array when no sectionLst is present', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getSlideSections(pres)).toEqual([]);
  });

  it('writes + reads back named sections', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    setSlideSections(pres, [
      { name: 'Intro', slides: [slides[0]!] },
      { name: 'Main', slides: [slides[1]!] },
    ]);
    const sections = getSlideSections(pres);
    expect(sections.map((s) => s.name)).toEqual(['Intro', 'Main']);
    expect(sections.map((s) => s.id)).toEqual([
      '{625B007C-F2BA-51F0-8F2F-F5F2C8B573EC}',
      '{9E366A52-986F-5FA6-8DE1-65DC565CE9E2}',
    ]);
    expect(sections[0]!.slides).toHaveLength(1);
    expect(sections[1]!.slides).toHaveLength(1);

    // Round-trip through save → reload.
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    expect(getSlideSections(reloaded).map((s) => s.name)).toEqual(['Intro', 'Main']);

    setSlideSections(reloaded, [
      { name: 'Main', slides: [getSlides(reloaded)[1]!] },
      { name: 'Intro', slides: [getSlides(reloaded)[0]!] },
    ]);
    expect(getSlideSections(reloaded).map((s) => s.id)).toEqual([sections[1]!.id, sections[0]!.id]);
  });

  it('preserves caller-supplied section IDs', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    const id = '{12345678-1234-4ABC-8DEF-1234567890AB}';
    setSlideSections(pres, [{ id, name: 'Stable', slides }]);
    expect(getSlideSections(await loadPresentation(await savePresentation(pres)))[0]!.id).toBe(id);
  });

  it('normalizes explicit IDs and resolves generated-ID collisions deterministically', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    setSlideSections(pres, [
      { id: '{12345678-1234-4abc-8def-1234567890ab}', name: 'Explicit', slides },
      { name: 'Repeated', slides: [] },
      { name: 'Repeated', slides: [] },
    ]);

    const sections = getSlideSections(pres);
    expect(sections[0]!.id).toBe('{12345678-1234-4ABC-8DEF-1234567890AB}');
    expect(sections[1]!.id).not.toBe(sections[2]!.id);

    const second = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setSlideSections(second, [
      { name: 'Repeated', slides: [] },
      { name: 'Repeated', slides: [] },
    ]);
    expect(getSlideSections(second).map((section) => section.id)).toEqual(
      sections.slice(1).map((section) => section.id),
    );
  });

  it('empty section list drops the extension entirely', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setSlideSections(pres, [{ name: 'X', slides: [] }]);
    expect(getSlideSections(pres)).toHaveLength(1);
    setSlideSections(pres, []);
    expect(getSlideSections(pres)).toEqual([]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const presBytes = readPackagePart(reloaded, '/ppt/presentation.xml');
    expect(presBytes).not.toBeNull();
    const xml = new TextDecoder().decode(presBytes!);
    expect(xml).not.toContain('521415D9-36F7-43E2-AB2F-B90AF26B5E84');
  });
});
