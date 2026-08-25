// `getShapeRunFormatEffective` — resolves a run's character properties by
// walking the ECMA-376 §21.1.2.4.7 inheritance chain (run → endParaRPr →
// defRPr → lstStyle → layout placeholder → master placeholder → master
// txStyles → theme fontScheme). The literal `getShapeRunFormat` only
// reports what's authored on the run itself, so any deck that relies on
// placeholder defaults (i.e. the vast majority of real PPTX content)
// needs this resolver to render faithfully.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addBlankSlide,
  addSlideTextBox,
  addTitleSlide,
  createPresentation,
  findSlidePlaceholder,
  getParagraphPropertiesEffective,
  getShapeRunFormat,
  getShapeRunFormatEffective,
  getSlides,
  inches,
  loadPresentation,
  setPresentationFonts,
  setShapeRunFormat,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeRunFormatEffective', () => {
  it('returns the literal rPr value when one is set on the run', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'styled',
    });
    setShapeRunFormat(tb, 0, 0, { bold: true, color: '#FF0000', size: 18 });
    const fmt = getShapeRunFormatEffective(pres, tb, 0, 0);
    expect(fmt.bold).toBe(true);
    expect(fmt.color).toBe('#FF0000');
    expect(fmt.size).toBeCloseTo(18);
  });

  it('falls back to the theme font when the run does not set one', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'plain',
    });
    // Literal-only getter only sees the synthesized `<a:rPr lang=…/>`
    // attached during `setShapeText`, with no typeface child.
    expect(getShapeRunFormat(tb, 0, 0)?.font).toBeUndefined();
    // The cascade should resolve the typeface from the theme's
    // minor font (body fallback). `two-slides.pptx` carries the
    // Office theme, which sets both major and minor to Calibri.
    const fmt = getShapeRunFormatEffective(pres, tb, 0, 0);
    expect(fmt.font).toBe('Calibri');
  });

  it('falls back to the theme East Asian font independently of the Latin font', () => {
    const pres = createPresentation();
    setPresentationFonts(pres, { majorEastAsian: '游明朝', minorEastAsian: 'メイリオ' });
    const slide = addBlankSlide(pres);
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(2),
      text: 'plain',
    });
    expect(getShapeRunFormat(tb, 0, 0)?.fontEastAsian).toBeUndefined();
    const fmt = getShapeRunFormatEffective(pres, tb, 0, 0);
    expect(fmt.fontEastAsian).toBe('メイリオ');

    setShapeRunFormat(tb, 0, 0, { fontEastAsian: '游明朝' });
    const overridden = getShapeRunFormatEffective(pres, tb, 0, 0);
    expect(overridden.fontEastAsian).toBe('游明朝');
  });

  it('inherits placeholder size from the master title style', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const slide = addTitleSlide(pres, 'Hello');
    const title = findSlidePlaceholder(slide, 'title') ?? findSlidePlaceholder(slide, 'ctrTitle');
    expect(title).not.toBeNull();
    const fmt = getShapeRunFormatEffective(pres, title!, 0, 0);
    // The Office theme master sets the title to 44pt in its titleStyle.
    expect(fmt.size).toBeGreaterThan(28);
    // Title-class placeholders should pull the major font.
    expect(fmt.font).toBe('Calibri');
  });

  it('inherits a text outline from the master title style', async () => {
    const entries = unzipSync(await readFile(fixture('blank.pptx')));
    const masterPart = 'ppt/slideMasters/slideMaster1.xml';
    const masterXml = strFromU8(entries[masterPart]!);
    const outlinedMasterXml = masterXml.replace(
      '<a:defRPr sz="4400" kern="1200">',
      '<a:defRPr sz="4400" kern="1200"><a:ln w="19050"><a:solidFill><a:srgbClr val="3659E3"/></a:solidFill></a:ln>',
    );
    expect(outlinedMasterXml).not.toBe(masterXml);
    entries[masterPart] = strToU8(outlinedMasterXml);

    const pres = await loadPresentation(zipSync(entries));
    const slide = addTitleSlide(pres, 'Outlined title');
    const title = findSlidePlaceholder(slide, 'title') ?? findSlidePlaceholder(slide, 'ctrTitle');
    expect(title).not.toBeNull();
    expect(getShapeRunFormat(title!, 0, 0)?.outline).toBeUndefined();
    expect(getShapeRunFormatEffective(pres, title!, 0, 0).outline).toEqual({
      kind: 'solid',
      color: '#3659E3',
      widthPt: 1.5,
    });
  });

  it('resolves +mj-lt / +mn-lt theme tokens authored on the master', async () => {
    // PowerPoint masters typically set <a:latin typeface="+mj-lt"/> on
    // title placeholders. The cascade must re-write that token through
    // the theme so renderers get a real font family.
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const slide = addTitleSlide(pres, 'X');
    const title = findSlidePlaceholder(slide, 'title') ?? findSlidePlaceholder(slide, 'ctrTitle');
    expect(title).not.toBeNull();
    const fmt = getShapeRunFormatEffective(pres, title!, 0, 0);
    expect(fmt.font).toBeDefined();
    expect(fmt.font!.startsWith('+')).toBe(false);
  });

  it('does NOT inherit the master body size for a plain text box', async () => {
    // A text box (no <p:ph>) is not a placeholder, so it must not read the
    // master's bodyStyle. An unsized run resolves through the presentation's
    // 18pt default text style, not the master body size (which can be much
    // larger). Regression guard: previously phType===null conflated "text box"
    // with "body placeholder" and leaked the master body size.
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'unsized',
    });
    expect(getShapeRunFormat(tb, 0, 0)?.size).toBeUndefined();
    const fmt = getShapeRunFormatEffective(pres, tb, 0, 0);
    expect(fmt.size).toBe(18);
    // Font still resolves via the theme fontScheme fallback (not placeholder
    // inheritance), so this stays defined.
    expect(fmt.font).toBe('Calibri');
  });

  it('inherits presentation default run and paragraph styles for a plain text box', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const slide = addBlankSlide(pres);
    const tb = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'presentation defaults',
    });

    expect(getShapeRunFormat(tb, 0, 0)?.size).toBeUndefined();
    const effectiveRun = getShapeRunFormatEffective(pres, tb, 0, 0);
    expect(effectiveRun).toMatchObject({
      font: 'Calibri',
      size: 18,
      color: '#000000',
      kern: 1200,
    });
    expect(effectiveRun.fontEastAsian).toBeUndefined();
    expect(effectiveRun.fontComplexScript).toBeUndefined();
    expect(getParagraphPropertiesEffective(pres, tb, 0)).toMatchObject({
      align: 'left',
      marL: 0,
      rtl: false,
      defaultTabSize: inches(0.5),
    });
  });
});
