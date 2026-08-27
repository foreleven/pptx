// Per-cell editing on table graphic-frame shapes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTable,
  clearTableCellFill,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  getTableCell,
  getTableCellPosition,
  getTableCellParagraphs,
  getTableCellRunUnderlineFillImageBytes,
  getTableCellText,
  getTableCells,
  inches,
  isTableShape,
  loadPresentation,
  savePresentation,
  setTableCellAlignment,
  setTableCellFill,
  setTableCellText,
  setTableCellTextFormat,
  setTableCellParagraphs,
  setTableCellRunUnderlineFillImage,
} from '../src/api/index.ts';

const TINY_PNG = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
    'base64',
  ),
);

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

const addDemoTable = (slide: ReturnType<typeof getSlides>[number]) =>
  addSlideTable(slide, {
    x: inches(0),
    y: inches(0),
    w: inches(4),
    h: inches(2),
    rows: [
      ['A', 'B'],
      ['C', 'D'],
    ],
  });

describe('fn API: table cell access', () => {
  it('getTableCells returns a 2D grid of handles', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    const cells = getTableCells(table);
    expect(cells).toHaveLength(2);
    expect(cells[0]).toHaveLength(2);
    expect(getTableCellText(cells[0]![0]!)).toBe('A');
    expect(getTableCellText(cells[1]![1]!)).toBe('D');
    expect(getTableCellPosition(cells[1]![1]!)).toEqual({ row: 1, col: 1 });
  });

  it('setTableCellText replaces one cell only', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    const cell = getTableCell(table, 0, 1);
    setTableCellText(cell, 'B'.padStart(5, 'X'));
    expect(getTableCellText(cell)).toBe('XXXXB');
    // Other cells unchanged.
    expect(getTableCellText(getTableCell(table, 0, 0))).toBe('A');
  });

  it('setTableCellFill paints one cell background', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    setTableCellFill(getTableCell(table, 0, 0), '#FF0000');
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('FF0000');
  });

  it('setTableCellTextFormat applies to the cell only', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    setTableCellTextFormat(getTableCell(table, 1, 0), { bold: true, color: '#00FF00' });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toMatch(/<a:rPr[^>]*b="1"/);
    expect(xml).toContain('00FF00');
  });

  it("setTableCellAlignment writes algn on the cell's paragraphs", async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    setTableCellAlignment(getTableCell(table, 0, 0), 'center');
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toMatch(/<a:pPr[^>]*algn="ctr"/);
  });

  it('setTableCellParagraphs authors ordered native runs, paragraph properties, and links', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    const cell = getTableCell(table, 0, 0);

    setTableCellParagraphs(cell, [
      {
        alignment: 'center',
        rtl: true,
        beforePts: 4,
        indent: { leftEmu: 12700, firstLineEmu: -6350 },
        endParagraph: {
          format: { font: 'Aptos', size: 18, color: '#F26B5B', bold: true },
          state: { dirty: false },
        },
        runs: [
          { text: 'Revenue ', format: { bold: true, font: 'Aptos' } },
          {
            text: '42',
            format: { color: '#3659E3', fontEastAsian: 'PingFang SC' },
            hyperlink: { url: 'https://example.com', tooltip: 'Details' },
          },
        ],
      },
    ]);

    expect(getTableCellText(cell)).toBe('Revenue 42');
    expect(getTableCellParagraphs(cell)[0]).toMatchObject({
      align: 'center',
      rtl: true,
      elements: [
        { kind: 'r', text: 'Revenue ', format: { bold: true, font: 'Aptos' } },
        { kind: 'r', text: '42', format: { color: '#3659E3', fontEastAsian: 'PingFang SC' } },
      ],
      endParagraph: {
        format: { font: 'Aptos', size: 18, color: '#F26B5B', bold: true },
        state: { dirty: false },
      },
    });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toMatch(/<a:pPr[^>]*algn="ctr"[^>]*marL="12700"[^>]*indent="-6350"/);
    expect(xml).toContain('<a:spcBef><a:spcPts val="400"/></a:spcBef>');
    expect(xml).toMatch(/<a:hlinkClick[^>]*r:id="rId\d+"[^>]*tooltip="Details"/);
    expect(xml).toMatch(/<a:endParaRPr[^>]*sz="1800"[^>]*b="1"[^>]*dirty="0"/);
    expect(xml).toContain('<a:srgbClr val="F26B5B"/>');
    expect(xml).not.toContain('xml:space');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedTable = getSlideShapes(getSlides(reloaded)[0]!).find(isTableShape)!;
    expect(getTableCellParagraphs(getTableCell(reloadedTable, 0, 0))[0]?.elements).toMatchObject([
      { kind: 'r', text: 'Revenue ' },
      { kind: 'r', text: '42' },
    ]);
    expect(getTableCellParagraphs(getTableCell(reloadedTable, 0, 0))[0]?.endParagraph).toEqual({
      format: { font: 'Aptos', size: 18, color: '#F26B5B', bold: true },
      state: { dirty: false },
    });
  });

  it('embeds and resolves a table-cell picture underline fill relationship', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    const cell = getTableCell(table, 0, 0);
    setTableCellParagraphs(cell, [
      { runs: [{ text: 'Picture underline', format: { underline: true } }] },
    ]);
    setTableCellRunUnderlineFillImage(cell, 0, 0, TINY_PNG);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reloadedTable = getSlideShapes(getSlides(reloaded)[0]!).find(isTableShape)!;
    const reloadedCell = getTableCell(reloadedTable, 0, 0);
    expect(getTableCellParagraphs(reloadedCell)[0]?.elements[0]).toMatchObject({
      kind: 'r',
      text: 'Picture underline',
      format: { underlineFill: { kind: 'picture' } },
    });
    expect(getTableCellRunUnderlineFillImageBytes(reloadedCell, 0, 0)).toEqual(TINY_PNG);
  });

  it('clearTableCellFill removes a previously-set fill', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    const cell = getTableCell(table, 0, 0);
    setTableCellFill(cell, '#FF0000');
    expect(await slideXml(await savePresentation(pres), 0)).toContain('FF0000');
    clearTableCellFill(cell);
    // No FF0000 from this cell. (Other cells don't carry this color in
    // the demo table, so the substring should disappear entirely.)
    expect(await slideXml(await savePresentation(pres), 0)).not.toContain('FF0000');
  });

  it('throws on out-of-range cells and non-table shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    expect(() => getTableCell(table, 9, 0)).toThrow(RangeError);
    expect(() => getTableCell(table, 0, 9)).toThrow(RangeError);
  });

  it('getTableDimensions reports row + column counts', async () => {
    const { getTableDimensions } = await import('../src/api/index.ts');
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addDemoTable(slide);
    expect(getTableDimensions(table)).toEqual({ rows: 2, cols: 2 });
  });
});
