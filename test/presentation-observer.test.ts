import { describe, expect, test, vi } from 'vitest';
import {
  PresentationObserver,
  addBlankSlide,
  addSlideTextBox,
  applyPresentationMutationRecords,
  createPresentation,
  emu,
  getShapeId,
  getShapeText,
  getSlideShapes,
  getSlides,
  loadPresentation,
  removeShape,
  savePresentation,
  sendShapeToBack,
  setShapeText,
  type PresentationMutationRecord,
} from '../src/index.ts';

const nextObserverDelivery = async (): Promise<void> => {
  await Promise.resolve();
};

describe('PresentationObserver', () => {
  test('reports package parts added by a deck-level mutation', async () => {
    const presentation = createPresentation();
    const callback = vi.fn<(records: readonly PresentationMutationRecord[]) => void>();
    const observer = new PresentationObserver(callback);

    observer.connect(presentation);
    addBlankSlide(presentation);
    await nextObserverDelivery();

    expect(callback).toHaveBeenCalledTimes(1);
    const records = callback.mock.calls[0]?.[0] ?? [];
    expect(records).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ type: 'part-put', partName: '/ppt/presentation.xml' }),
        expect.objectContaining({
          type: 'part-put',
          partName: '/ppt/slides/slide1.xml',
          beforeHash: null,
        }),
      ]),
    );
  });

  test('reports committed public API mutations and stops after disconnect', async () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const title = addSlideTextBox(slide, {
      x: emu(100),
      y: emu(200),
      w: emu(3_000),
      h: emu(400),
      text: 'Before',
    });
    const callback = vi.fn<(records: readonly PresentationMutationRecord[]) => void>();
    const observer = new PresentationObserver(callback);

    observer.connect(presentation);
    setShapeText(title, 'After');
    await nextObserverDelivery();

    expect(callback).toHaveBeenCalledTimes(1);
    const records = callback.mock.calls[0]?.[0];
    expect(records).toHaveLength(1);
    expect(records?.[0]).toMatchObject({
      type: 'part-put',
      partName: '/ppt/slides/slide1.xml',
      contentType: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
      beforeHash: expect.any(String),
      afterHash: expect.any(String),
      data: expect.any(Uint8Array),
    });

    observer.disconnect();
    setShapeText(title, 'Disconnected');
    await nextObserverDelivery();

    expect(callback).toHaveBeenCalledTimes(1);
  });

  test('applies records to a mirror without echoing them back', async () => {
    const initial = createPresentation();
    const slide = addBlankSlide(initial);
    addSlideTextBox(slide, {
      x: emu(100),
      y: emu(200),
      w: emu(3_000),
      h: emu(400),
      text: 'Before',
    });
    const initialBytes = await savePresentation(initial);
    const source = await loadPresentation(initialBytes);
    const mirror = await loadPresentation(initialBytes);
    const deliveries: Array<readonly PresentationMutationRecord[]> = [];
    const sourceObserver = new PresentationObserver((records) => deliveries.push(records));
    const mirrorCallback = vi.fn<(records: readonly PresentationMutationRecord[]) => void>();
    const mirrorObserver = new PresentationObserver(mirrorCallback);

    sourceObserver.connect(source);
    mirrorObserver.connect(mirror);
    const retainedMirrorShape = getSlideShapes(getSlides(mirror)[0]!)[0]!;
    setShapeText(getSlideShapes(getSlides(source)[0]!)[0]!, 'After');
    await nextObserverDelivery();

    expect(deliveries).toHaveLength(1);
    applyPresentationMutationRecords(mirror, deliveries[0]!, { notifyObservers: false });
    await nextObserverDelivery();

    const mirroredShape = getSlideShapes(getSlides(mirror)[0]!)[0]!;
    expect(getShapeText(mirroredShape)).toBe('After');
    expect(getShapeText(retainedMirrorShape)).toBe('After');
    expect(mirrorCallback).not.toHaveBeenCalled();
  });

  test('rejects a part body that does not match its advertised afterHash', async () => {
    const presentation = createPresentation();
    expect(() =>
      applyPresentationMutationRecords(presentation, [
        {
          type: 'part-put',
          partName: '/ppt/slides/slide99.xml',
          contentType: 'application/xml',
          beforeHash: null,
          afterHash: 'fnv1a64:0000000000000000',
          data: new TextEncoder().encode('<p:sld/>'),
        },
      ]),
    ).toThrow(/does not match/);
  });

  test('refreshes retained remote shape handles by stable shape id', async () => {
    const initial = createPresentation();
    const slide = addBlankSlide(initial);
    addSlideTextBox(slide, {
      x: emu(100),
      y: emu(100),
      w: emu(1_000),
      h: emu(300),
      text: 'First',
    });
    addSlideTextBox(slide, {
      x: emu(100),
      y: emu(500),
      w: emu(1_000),
      h: emu(300),
      text: 'Second',
    });
    const bytes = await savePresentation(initial);
    const source = await loadPresentation(bytes);
    const mirror = await loadPresentation(bytes);
    const deliveries: Array<readonly PresentationMutationRecord[]> = [];
    const observer = new PresentationObserver((records) => deliveries.push(records));
    observer.connect(source);

    const retained = getSlideShapes(getSlides(mirror)[0]!);
    const firstId = getShapeId(retained[0]!);
    const secondId = getShapeId(retained[1]!);
    sendShapeToBack(getSlideShapes(getSlides(source)[0]!)[1]!);
    await nextObserverDelivery();
    applyPresentationMutationRecords(mirror, deliveries.shift()!, { notifyObservers: false });

    const reordered = getSlideShapes(getSlides(mirror)[0]!);
    expect(reordered.map(getShapeId)).toEqual([secondId, firstId]);
    expect(reordered[0]).toBe(retained[1]);
    expect(reordered[1]).toBe(retained[0]);

    removeShape(
      getSlideShapes(getSlides(source)[0]!).find((shape) => getShapeId(shape) === firstId)!,
    );
    addSlideTextBox(getSlides(source)[0]!, {
      x: emu(100),
      y: emu(900),
      w: emu(1_000),
      h: emu(300),
      text: 'Replacement',
    });
    await nextObserverDelivery();
    applyPresentationMutationRecords(mirror, deliveries.shift()!, { notifyObservers: false });

    const replaced = getSlideShapes(getSlides(mirror)[0]!);
    expect(replaced).toContain(retained[1]);
    expect(replaced).not.toContain(retained[0]);
    expect(getShapeText(retained[0]!)).toBe('First');
    expect(replaced.map(getShapeText)).toEqual(['Second', 'Replacement']);
  });
});
