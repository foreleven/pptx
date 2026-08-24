import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getParagraphDefaultTabSize,
  getParagraphPropertiesEffective,
  getParagraphTabStops,
  getSlideXmlString,
  inches,
  pt,
  setParagraphBullet,
  setParagraphDefaultTabSize,
  setParagraphTabStops,
} from '../src/api/index.ts';

describe('fn API: paragraph tabs', () => {
  it('round-trips a default tab size and the complete tab-stop alignment catalog', () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(7),
      h: inches(2),
      text: 'Name\tQ1\tQ2',
    });
    setParagraphBullet(shape, 0, 'none');
    setParagraphDefaultTabSize(shape, 0, pt(36));
    setParagraphTabStops(shape, 0, [
      { positionEmu: pt(72), alignment: 'left' },
      { positionEmu: pt(144), alignment: 'center' },
      { positionEmu: pt(216), alignment: 'right' },
      { positionEmu: pt(288), alignment: 'decimal' },
    ]);

    expect(getParagraphDefaultTabSize(shape, 0)).toBe(pt(36));
    expect(getParagraphTabStops(shape, 0)).toEqual([
      { positionEmu: pt(72), alignment: 'left' },
      { positionEmu: pt(144), alignment: 'center' },
      { positionEmu: pt(216), alignment: 'right' },
      { positionEmu: pt(288), alignment: 'decimal' },
    ]);
    expect(getParagraphPropertiesEffective(presentation, shape, 0)).toMatchObject({
      defaultTabSize: pt(36),
      tabStops: [
        { positionEmu: pt(72), alignment: 'left' },
        { positionEmu: pt(144), alignment: 'center' },
        { positionEmu: pt(216), alignment: 'right' },
        { positionEmu: pt(288), alignment: 'decimal' },
      ],
    });

    const xml = getSlideXmlString(slide);
    expect(xml).toContain(`defTabSz="${String(pt(36))}"`);
    expect(xml).toMatch(
      /<a:buNone\/><a:tabLst><a:tab pos="914400" algn="l"\/><a:tab pos="1828800" algn="ctr"\/><a:tab pos="2743200" algn="r"\/><a:tab pos="3657600" algn="dec"\/><\/a:tabLst>/u,
    );
  });

  it('clears authored tab properties and validates schema boundaries', () => {
    const presentation = createPresentation();
    const shape = addSlideTextBox(addBlankSlide(presentation), {
      x: inches(1),
      y: inches(1),
      w: inches(7),
      h: inches(2),
      text: 'A\tB',
    });

    setParagraphDefaultTabSize(shape, 0, pt(36));
    setParagraphTabStops(shape, 0, [{ positionEmu: pt(72), alignment: 'left' }]);
    setParagraphDefaultTabSize(shape, 0, null);
    setParagraphTabStops(shape, 0, null);
    expect(getParagraphDefaultTabSize(shape, 0)).toBeNull();
    expect(getParagraphTabStops(shape, 0)).toBeNull();

    expect(() => setParagraphDefaultTabSize(shape, 0, -1)).toThrow(RangeError);
    expect(() =>
      setParagraphTabStops(shape, 0, [{ positionEmu: 2_147_483_648, alignment: 'left' }]),
    ).toThrow(RangeError);
    expect(() =>
      setParagraphTabStops(
        shape,
        0,
        Array.from({ length: 33 }, (_, index) => ({
          positionEmu: pt(index),
          alignment: 'left' as const,
        })),
      ),
    ).toThrow(RangeError);
    expect(() =>
      setParagraphTabStops(shape, 0, [{ positionEmu: pt(72), alignment: 'invalid' as never }]),
    ).toThrow(RangeError);
  });
});
