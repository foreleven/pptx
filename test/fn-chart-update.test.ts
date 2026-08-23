// setChartSpec — update an existing chart's data in place.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  getSlideCharts,
  getSlides,
  inches,
  listPackageParts,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setChartSpec,
} from '../src/api/index.ts';
import { readZip } from '../src/internal/opc/zip.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setChartSpec', () => {
  it('replaces chart data while preserving the shape geometry', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['Q1', 'Q2'],
        series: [{ name: 'Old', values: [1, 2] }],
      },
    });

    // Round-trip into a fresh handle, then update via setChartSpec.
    const reloaded = await loadPresentation(await savePresentation(pres));
    const charts = getSlideCharts(getSlides(reloaded)[0]!);
    expect(charts).toHaveLength(1);
    setChartSpec(charts[0]!, {
      kind: 'line',
      categories: ['Jan', 'Feb', 'Mar'],
      series: [{ name: 'New', values: [10, 20, 30] }],
      title: 'After',
    });

    // Re-load to verify persistence.
    const reread = await loadPresentation(await savePresentation(reloaded));
    const after = getSlideCharts(getSlides(reread)[0]!)[0]!.spec!;
    expect(after.kind).toBe('line');
    expect(after.categories).toEqual(['Jan', 'Feb', 'Mar']);
    expect(after.series[0]!.name).toBe('New');
    expect(after.series[0]!.values).toEqual([10, 20, 30]);
    expect(after.title).toBe('After');
  });

  it('also rewrites the embedded xlsx', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(3),
      spec: {
        kind: 'column',
        categories: ['A'],
        series: [{ name: 'orig', values: [1] }],
      },
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const chart = getSlideCharts(getSlides(reloaded)[0]!)[0]!;

    setChartSpec(chart, {
      kind: 'column',
      categories: ['Updated'],
      series: [{ name: 'renamed-series', values: [42] }],
    });

    const bytes = await savePresentation(reloaded);
    const after = await loadPresentation(bytes);
    const xlsx = listPackageParts(after).find((p) =>
      /^\/ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/.test(p.name),
    );
    expect(xlsx).toBeDefined();
    const chartBytes = readPackagePart(after, '/ppt/charts/chart1.xml');
    expect(chartBytes).not.toBeNull();
    expect(new TextDecoder().decode(chartBytes!)).toContain('renamed-series');
  });

  it('throws when invoked on a shape that is not a chart frame', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideCharts(slide)).toEqual([]);
    // Synthesize a "chart data" handle pointing at a non-chart shape.
    const { getSlideShapes } = await import('../src/api/index.ts');
    const shape = getSlideShapes(slide)[0]!;
    expect(() =>
      setChartSpec(
        { shape, spec: null },
        { kind: 'column', categories: ['x'], series: [{ name: 'y', values: [1] }] },
      ),
    ).toThrow(/not a chart/);
  });

  it('rejects radar authoring while scatter and bubble use native xy(z) channels', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(() =>
      addSlideChart(slide, {
        x: inches(0),
        y: inches(0),
        w: inches(4),
        h: inches(3),
        spec: { kind: 'radar', categories: ['A'], series: [{ name: 'S', values: [1] }] },
      }),
    ).toThrow(/read-only/);
  });

  it('rewrites scatter and bubble chart XML, workbooks, and reload semantics', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(3),
      spec: { kind: 'column', categories: ['A'], series: [{ name: 'S', values: [1] }] },
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const chart = getSlideCharts(getSlides(reloaded)[0]!)[0]!;
    setChartSpec(chart, {
      kind: 'scatter',
      categories: [],
      scatterStyle: 'marker',
      series: [{ name: 'Observations', values: [10, 20], xValues: [2, 4] }],
    });

    const scatterBytes = await savePresentation(reloaded);
    const scatterReloaded = await loadPresentation(scatterBytes);
    expect(getSlideCharts(getSlides(scatterReloaded)[0]!)[0]!.spec).toMatchObject({
      kind: 'scatter',
      scatterStyle: 'marker',
      series: [{ name: 'Observations', values: [10, 20], xValues: [2, 4] }],
    });
    const scatterChartXml = decodePart(scatterReloaded, '/ppt/charts/chart1.xml');
    expect(scatterChartXml).toContain('<c:scatterChart>');
    expect(scatterChartXml).toContain('<c:xVal>');
    expect(scatterChartXml).toContain('<c:f>Sheet1!$A$2:$A$3</c:f>');
    expect(scatterChartXml).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>');
    const scatterSheetXml = decodeWorkbookSheet(scatterReloaded);
    expect(scatterSheetXml).toContain('<t>Observations X</t>');
    expect(scatterSheetXml).toContain('<c r="A3"><v>4</v></c>');
    expect(scatterSheetXml).toContain('<c r="B3"><v>20</v></c>');

    setChartSpec(getSlideCharts(getSlides(scatterReloaded)[0]!)[0]!, {
      kind: 'bubble',
      categories: [],
      bubbleScale: 130,
      bubbleSizeRepresents: 'width',
      series: [
        {
          name: 'Markets',
          values: [5, 7, 11],
          xValues: [3, 6, 9],
          bubbleSizes: [4, 16, 36],
        },
      ],
    });

    const bubbleReloaded = await loadPresentation(await savePresentation(scatterReloaded));
    expect(getSlideCharts(getSlides(bubbleReloaded)[0]!)[0]!.spec).toMatchObject({
      kind: 'bubble',
      bubbleScale: 130,
      bubbleSizeRepresents: 'width',
      series: [
        {
          name: 'Markets',
          values: [5, 7, 11],
          xValues: [3, 6, 9],
          bubbleSizes: [4, 16, 36],
        },
      ],
    });
    const bubbleChartXml = decodePart(bubbleReloaded, '/ppt/charts/chart1.xml');
    expect(bubbleChartXml).toContain('<c:bubbleChart>');
    expect(bubbleChartXml).toContain('<c:bubbleSize>');
    expect(bubbleChartXml).toContain('<c:f>Sheet1!$C$2:$C$4</c:f>');
    const bubbleSheetXml = decodeWorkbookSheet(bubbleReloaded);
    expect(bubbleSheetXml).toContain('<t>Markets Size</t>');
    expect(bubbleSheetXml).toContain('<c r="C4"><v>36</v></c>');
  });

  it('setChartSpec rejects switching an existing chart to radar', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideChart(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(3),
      spec: { kind: 'column', categories: ['A'], series: [{ name: 'S', values: [1] }] },
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const chart = getSlideCharts(getSlides(reloaded)[0]!)[0]!;
    expect(() =>
      setChartSpec(chart, {
        kind: 'radar',
        categories: ['A'],
        series: [{ name: 'S', values: [1] }],
      }),
    ).toThrow(/read-only/);
  });
});

const decodePart = (
  presentation: Awaited<ReturnType<typeof loadPresentation>>,
  partName: string,
): string => {
  const bytes = readPackagePart(presentation, partName);
  expect(bytes).not.toBeNull();
  return new TextDecoder().decode(bytes!);
};

const decodeWorkbookSheet = (
  presentation: Awaited<ReturnType<typeof loadPresentation>>,
): string => {
  const workbook = listPackageParts(presentation).find((part) =>
    /^\/ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/.test(part.name),
  );
  expect(workbook).toBeDefined();
  const bytes = readPackagePart(presentation, workbook!.name);
  expect(bytes).not.toBeNull();
  const zip = readZip(bytes!);
  const sheet = zip.entries.find((entry) => entry.name === 'xl/worksheets/sheet1.xml');
  expect(sheet).toBeDefined();
  return new TextDecoder().decode(sheet!.data);
};
