// Slide-level introspection: `getSlideTransition` + `getSlideBackground`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  clearSlideBackground,
  clearSlideTransition,
  createPresentation,
  getSlideBackground,
  getSlideName,
  getSlideShowMasterPlaceholderAnimations,
  getSlideShowMasterShapes,
  getSlideTransition,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideBackground,
  setSlideName,
  setSlideShowMasterPlaceholderAnimations,
  setSlideShowMasterShapes,
  setSlideTransition,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: slide introspection', () => {
  it('getSlideTransition returns null before any set, and the configured effect after', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideTransition(slide)).toBeNull();

    setSlideTransition(slide, { effect: 'fade', speed: 'fast' });
    const got = getSlideTransition(slide);
    expect(got?.effect).toBe('fade');
    expect(got?.speed).toBe('fast');

    clearSlideTransition(slide);
    expect(getSlideTransition(slide)).toBeNull();
  });

  it('reads a no-effect transition when it carries slide-advance controls', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);

    setSlideTransition(slide, {
      effect: 'none',
      advanceOnClick: false,
      advanceAfterMs: 2500,
    });

    expect(getSlideTransition(slide)).toEqual({
      effect: 'none',
      advanceOnClick: false,
      advanceAfterMs: 2500,
    });
  });

  it('getSlideBackground reports inherit when no <p:bg> is set', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideBackground(slide).kind).toBe('inherit');
  });

  it('getSlideBackground reads back a solid color after setSlideBackground', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideBackground(slide, '#ABCDEF');
    expect(getSlideBackground(slide)).toEqual({ kind: 'solid', color: '#ABCDEF' });

    clearSlideBackground(slide);
    expect(getSlideBackground(slide).kind).toBe('inherit');
  });

  it('round-trips slide name and master-visibility flags with schema defaults', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);

    expect(getSlideName(slide)).toBe('');
    expect(getSlideShowMasterShapes(slide)).toBe(true);
    expect(getSlideShowMasterPlaceholderAnimations(slide)).toBe(true);

    setSlideName(slide, 'Quarterly appendix');
    setSlideShowMasterShapes(slide, false);
    setSlideShowMasterPlaceholderAnimations(slide, false);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedSlide = getSlides(reloaded)[0]!;
    expect(getSlideName(reloadedSlide)).toBe('Quarterly appendix');
    expect(getSlideShowMasterShapes(reloadedSlide)).toBe(false);
    expect(getSlideShowMasterPlaceholderAnimations(reloadedSlide)).toBe(false);

    setSlideName(reloadedSlide, '');
    setSlideShowMasterShapes(reloadedSlide, true);
    setSlideShowMasterPlaceholderAnimations(reloadedSlide, true);
    const xml = getSlideXmlString(reloadedSlide);
    expect(xml).not.toContain('name="Quarterly appendix"');
    expect(xml).not.toContain('showMasterSp=');
    expect(xml).not.toContain('showMasterPhAnim=');
  });

  it('reads every xsd:boolean lexical form for default-true slide flags', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    setSlideShowMasterShapes(slide, false);
    setSlideShowMasterPlaceholderAnimations(slide, false);

    const entries = unzipSync(await savePresentation(pres));
    const slideEntry = entries['ppt/slides/slide1.xml']!;
    const xml = new TextDecoder()
      .decode(slideEntry)
      .replace('showMasterSp="0"', 'showMasterSp="false"')
      .replace('showMasterPhAnim="0"', 'showMasterPhAnim="true"');
    entries['ppt/slides/slide1.xml'] = new TextEncoder().encode(xml);

    const reloaded = await loadPresentation(zipSync(entries));
    const reloadedSlide = getSlides(reloaded)[0]!;
    expect(getSlideShowMasterShapes(reloadedSlide)).toBe(false);
    expect(getSlideShowMasterPlaceholderAnimations(reloadedSlide)).toBe(true);

    const invalidEntries = unzipSync(await savePresentation(pres));
    const invalidSlideEntry = invalidEntries['ppt/slides/slide1.xml']!;
    invalidEntries['ppt/slides/slide1.xml'] = new TextEncoder().encode(
      new TextDecoder().decode(invalidSlideEntry).replace('showMasterSp="0"', 'showMasterSp="yes"'),
    );
    const invalid = await loadPresentation(zipSync(invalidEntries));
    expect(() => getSlideShowMasterShapes(getSlides(invalid)[0]!)).toThrow(
      'showMasterSp: invalid xsd:boolean value "yes"',
    );
  });
});
