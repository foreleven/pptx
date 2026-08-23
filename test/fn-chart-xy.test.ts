import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideChart,
  createPresentation,
  getShapeChartSpec,
  getSlideShapes,
  getSlides,
  inches,
  isChartShape,
  loadPresentation,
  readPackagePart,
  savePresentation,
} from '../src/api/index.ts';
import { readZip } from '../src/internal/opc/zip.ts';

const decode = (bytes: Uint8Array | null): string => {
  expect(bytes).not.toBeNull();
  return new TextDecoder().decode(bytes!);
};

describe('fn API: xy chart authoring', () => {
  it('creates an editable scatter chart and preserves its x/y channels', async () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    addSlideChart(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'scatter',
        categories: [],
        scatterStyle: 'smoothMarker',
        categoryAxisHidden: true,
        valueAxisHidden: false,
        valueAxisMajorGridlines: true,
        valueAxisMajorGridlineColor: '#C44536',
        categoryAxisLabelStyle: { font: 'Aptos', sizePt: 11, color: '#274C77' },
        valueAxisLabelStyle: { font: 'Aptos Display', sizePt: 13, color: '#6096BA' },
        series: [
          {
            name: 'Observations',
            xValues: [1, 2.5, 4],
            values: [8, 3, 11],
            color: '#3659E3',
            markerSymbol: 'diamond',
            markerSizePt: 8,
            smooth: true,
          },
        ],
      },
    });

    const bytes = await savePresentation(presentation);
    const chartXml = decode(readPackagePart(presentation, '/ppt/charts/chart1.xml'));
    expect(chartXml).toContain('<c:scatterChart>');
    expect(chartXml).toContain('<c:scatterStyle val="smoothMarker"/>');
    expect(chartXml).toContain('<c:xVal>');
    expect(chartXml).toContain('<c:yVal>');
    expect(chartXml).not.toContain('<c:cat>');
    expect(chartXml).not.toContain('<c:catAx>');

    const xlsxBytes = readPackagePart(
      presentation,
      '/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx',
    );
    expect(xlsxBytes).not.toBeNull();
    const xlsx = readZip(xlsxBytes!);
    const sheetXml = decode(
      xlsx.entries.find((entry) => entry.name === 'xl/worksheets/sheet1.xml')?.data ?? null,
    );
    expect(sheetXml).toContain('<t>Observations X</t>');
    expect(sheetXml).toContain('<t>Observations</t>');
    expect(sheetXml).toContain('<c r="A2"><v>1</v></c>');
    expect(sheetXml).toContain('<c r="B2"><v>8</v></c>');

    const reloaded = await loadPresentation(bytes);
    const shape = getSlideShapes(getSlides(reloaded)[0]!).find(isChartShape)!;
    expect(getShapeChartSpec(shape)).toMatchObject({
      kind: 'scatter',
      scatterStyle: 'smoothMarker',
      categoryAxisHidden: true,
      valueAxisHidden: false,
      valueAxisMajorGridlines: true,
      valueAxisMajorGridlineColor: '#C44536',
      categoryAxisLabelStyle: { font: 'Aptos', sizePt: 11, color: '#274C77' },
      valueAxisLabelStyle: { font: 'Aptos Display', sizePt: 13, color: '#6096BA' },
      series: [
        {
          name: 'Observations',
          xValues: [1, 2.5, 4],
          values: [8, 3, 11],
          markerSymbol: 'diamond',
          markerSizePt: 8,
          smooth: true,
        },
      ],
    });
  });

  it('creates an editable bubble chart and preserves its x/y/size channels', async () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    addSlideChart(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'bubble',
        categories: [],
        bubbleScale: 125,
        bubbleSizeRepresents: 'width',
        categoryAxisHidden: false,
        valueAxisHidden: true,
        valueAxisMajorGridlines: false,
        categoryAxisLabelStyle: { font: 'Aptos Narrow', sizePt: 10, color: '#5F0F40' },
        valueAxisLabelStyle: { font: 'Aptos', sizePt: 14, color: '#0F4C5C' },
        series: [
          {
            name: 'Markets',
            xValues: [1, 3, 6],
            values: [4, 9, 5],
            bubbleSizes: [10, 35, 80],
            color: '#E86A33',
          },
        ],
      },
    });

    const bytes = await savePresentation(presentation);
    const chartXml = decode(readPackagePart(presentation, '/ppt/charts/chart1.xml'));
    expect(chartXml).toContain('<c:bubbleChart>');
    expect(chartXml).toContain('<c:xVal>');
    expect(chartXml).toContain('<c:yVal>');
    expect(chartXml).toContain('<c:bubbleSize>');
    expect(chartXml).toContain('<c:bubbleScale val="125"/>');
    expect(chartXml).toContain('<c:sizeRepresents val="w"/>');
    expect(chartXml).not.toContain('<c:catAx>');

    const xlsxBytes = readPackagePart(
      presentation,
      '/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx',
    );
    expect(xlsxBytes).not.toBeNull();
    const xlsx = readZip(xlsxBytes!);
    const sheetXml = decode(
      xlsx.entries.find((entry) => entry.name === 'xl/worksheets/sheet1.xml')?.data ?? null,
    );
    expect(sheetXml).toContain('<t>Markets X</t>');
    expect(sheetXml).toContain('<t>Markets</t>');
    expect(sheetXml).toContain('<t>Markets Size</t>');
    expect(sheetXml).toContain('<c r="C2"><v>10</v></c>');

    const reloaded = await loadPresentation(bytes);
    const shape = getSlideShapes(getSlides(reloaded)[0]!).find(isChartShape)!;
    expect(getShapeChartSpec(shape)).toMatchObject({
      kind: 'bubble',
      bubbleScale: 125,
      bubbleSizeRepresents: 'width',
      categoryAxisHidden: false,
      valueAxisHidden: true,
      valueAxisMajorGridlines: false,
      categoryAxisLabelStyle: { font: 'Aptos Narrow', sizePt: 10, color: '#5F0F40' },
      valueAxisLabelStyle: { font: 'Aptos', sizePt: 14, color: '#0F4C5C' },
      series: [
        {
          name: 'Markets',
          xValues: [1, 3, 6],
          values: [4, 9, 5],
          bubbleSizes: [10, 35, 80],
        },
      ],
    });
  });
});
