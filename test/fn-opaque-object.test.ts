import { describe, expect, it, vi } from 'vitest';
import {
  addBlankSlide,
  addSlideChart,
  addSlideOpaqueObject,
  addSlideShape,
  createPresentation,
  extractOpaqueObjectFragment,
  getGroupChildren,
  getShapeChartSpec,
  getShapeClickAction,
  getShapeId,
  getSlideTopLevelShapes,
  getSlides,
  groupShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeClickAction,
} from '../src/api/index.ts';

describe('fn API: opaque OOXML objects', () => {
  it('extracts byte-identical fragments regardless of the current clock', () => {
    const source = createPresentation();
    const sourceSlide = addBlankSlide(source);
    const chart = addSlideChart(sourceSlide, {
      spec: {
        kind: 'column',
        categories: ['North', 'South'],
        series: [{ name: 'Revenue', values: [12, 18] }],
      },
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(3),
    });

    vi.useFakeTimers();
    try {
      vi.setSystemTime(new Date('2026-01-01T00:00:00Z'));
      const first = extractOpaqueObjectFragment(chart);
      vi.setSystemTime(new Date('2026-08-23T00:00:00Z'));
      expect(extractOpaqueObjectFragment(chart)).toEqual(first);
    } finally {
      vi.useRealTimers();
    }
  });

  it('copies a shape and its complete internal relationship closure without fetching external targets', async () => {
    const source = createPresentation();
    const sourceSlide = addBlankSlide(source);
    const chart = addSlideChart(sourceSlide, {
      spec: {
        kind: 'column',
        categories: ['North', 'South'],
        series: [{ name: 'Revenue', values: [12, 18] }],
      },
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(3),
    });
    setShapeClickAction(chart, { kind: 'url', url: 'https://example.com/opaque-chart' });

    const fragment = extractOpaqueObjectFragment(chart);
    const target = createPresentation();
    const targetSlide = addBlankSlide(target);
    addSlideOpaqueObject(targetSlide, fragment);

    const loaded = await loadPresentation(await savePresentation(target));
    const rebuilt = getSlideTopLevelShapes(getSlides(loaded)[0]!)[0]!;
    expect(getShapeChartSpec(rebuilt)).toMatchObject({
      kind: 'column',
      categories: ['North', 'South'],
      series: [{ name: 'Revenue', values: [12, 18] }],
    });
    expect(getShapeClickAction(rebuilt)).toEqual({
      kind: 'url',
      url: 'https://example.com/opaque-chart',
    });
  });

  it('rejects malformed or unsafe fragment archives before insertion', () => {
    const target = createPresentation();
    const slide = addBlankSlide(target);
    expect(() => addSlideOpaqueObject(slide, Uint8Array.from([80, 75, 3, 4]))).toThrow(
      /safety limit|compression ratio/u,
    );
  });

  it('rejects a fragment whose slide relationship is absent from the target deck', () => {
    const source = createPresentation();
    const sourceSlide = addBlankSlide(source);
    const linkedSlide = addBlankSlide(source);
    const chart = addSlideChart(sourceSlide, {
      spec: { kind: 'column', categories: ['A'], series: [{ name: 'Value', values: [1] }] },
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(3),
    });
    setShapeClickAction(chart, { kind: 'slide', slide: linkedSlide });
    const fragment = extractOpaqueObjectFragment(chart);

    const target = createPresentation();
    const targetSlide = addBlankSlide(target);
    expect(() => addSlideOpaqueObject(targetSlide, fragment)).toThrow(
      /missing relationship closure.*slide2/u,
    );
  });

  it('assigns unique shape ids to an opaque group and every descendant', async () => {
    const source = createPresentation();
    const sourceSlide = addBlankSlide(source);
    const first = addSlideShape(sourceSlide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    });
    const second = addSlideShape(sourceSlide, {
      preset: 'ellipse',
      x: inches(4),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    });
    const fragment = extractOpaqueObjectFragment(groupShapes([first, second]));

    const target = createPresentation();
    const targetSlide = addBlankSlide(target);
    addSlideShape(targetSlide, {
      preset: 'rect',
      x: inches(0.25),
      y: inches(0.25),
      w: inches(0.5),
      h: inches(0.5),
    });
    addSlideOpaqueObject(targetSlide, fragment);

    const loaded = await loadPresentation(await savePresentation(target));
    const shapes = getSlideTopLevelShapes(getSlides(loaded)[0]!);
    const ids = [...shapes.map(getShapeId), ...getGroupChildren(shapes[1]!).map(getShapeId)];
    expect(new Set(ids).size).toBe(ids.length);
  });
});
