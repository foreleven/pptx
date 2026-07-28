import { describe, expect, test } from 'vitest';
import {
  PresentationObserver,
  addBlankSlide,
  addSlideShape,
  addSlideTextBox,
  createPresentation,
  getSlidePartName,
  inches,
  setShapeText,
  setShapeGradientFill,
  setCoreProperties,
  type PresentationMutationRecord,
} from '../src/api/index.ts';
import { PresentationRenderer } from '../packages/preview/src/index.ts';

const addTextSlide = (presentation: ReturnType<typeof createPresentation>, text: string) => {
  const slide = addBlankSlide(presentation);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    text,
  });
  return { slide, shape };
};

describe('PresentationRenderer', () => {
  test('renders all slides and rerenders only a slide changed by observer records', async () => {
    const presentation = createPresentation();
    const first = addTextSlide(presentation, 'First');
    const second = addTextSlide(presentation, 'Second');
    const deliveries: Array<readonly PresentationMutationRecord[]> = [];
    const observer = new PresentationObserver((records) => deliveries.push(records));
    const renderer = new PresentationRenderer();

    observer.connect(presentation);
    renderer.connect(presentation);

    const initial = renderer.renderAll();
    expect(initial.map((rendered) => rendered.slidePartName)).toEqual([
      getSlidePartName(first.slide),
      getSlidePartName(second.slide),
    ]);
    expect(initial[0]?.svg).toContain('First');
    expect(initial[1]?.svg).toContain('Second');

    setShapeText(second.shape, 'Updated');
    await Promise.resolve();

    const updates = renderer.renderChanges(deliveries[0]!);
    expect(updates).toHaveLength(1);
    expect(updates[0]?.slidePartName).toBe(getSlidePartName(second.slide));
    expect(updates[0]?.svg).toContain('Updated');
  });

  test('keeps definition ids deterministic across repeated renders', () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(2),
    });
    setShapeGradientFill(shape, {
      stops: [
        { offset: 0, color: '#0000FF' },
        { offset: 1, color: '#FF0000' },
      ],
      angleDeg: 0,
    });
    const renderer = new PresentationRenderer();
    renderer.connect(presentation);
    expect(renderer.renderAll()[0]?.svg).toBe(renderer.renderAll()[0]?.svg);
  });

  test('does not rerender slides for non-visual package metadata changes', async () => {
    const presentation = createPresentation();
    addTextSlide(presentation, 'Visible');
    const deliveries: Array<readonly PresentationMutationRecord[]> = [];
    const observer = new PresentationObserver((records) => deliveries.push(records));
    const renderer = new PresentationRenderer();
    observer.connect(presentation);
    renderer.connect(presentation);
    renderer.renderAll();

    setCoreProperties(presentation, { title: 'Metadata only' });
    await Promise.resolve();

    expect(deliveries[0]?.some((record) => record.partName === '/docProps/core.xml')).toBe(true);
    expect(renderer.renderChanges(deliveries[0]!)).toEqual([]);
  });
});
