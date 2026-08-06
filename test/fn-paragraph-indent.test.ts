// Paragraph indentation writer — the counterpart to getParagraphIndent.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getParagraphIndent,
  inches,
  pt,
  setParagraphIndent,
} from '../src/api/index.ts';

describe('fn API: setParagraphIndent', () => {
  it('round-trips left, right, and first-line indentation independently', () => {
    const shape = addSlideTextBox(addBlankSlide(createPresentation()), {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: 'Indented paragraph',
    });

    setParagraphIndent(shape, 0, {
      leftEmu: pt(18),
      rightEmu: pt(9),
      firstLineEmu: pt(-6),
    });
    expect(getParagraphIndent(shape, 0)).toEqual({
      leftEmu: pt(18),
      rightEmu: pt(9),
      firstLineEmu: pt(-6),
    });

    setParagraphIndent(shape, 0, { rightEmu: null });
    expect(getParagraphIndent(shape, 0)).toEqual({
      leftEmu: pt(18),
      rightEmu: null,
      firstLineEmu: pt(-6),
    });
  });

  it('rejects negative paragraph margins but permits hanging first-line indentation', () => {
    const shape = addSlideTextBox(addBlankSlide(createPresentation()), {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(1),
      text: 'A',
    });

    expect(() => setParagraphIndent(shape, 0, { leftEmu: -1 })).toThrow(RangeError);
    expect(() => setParagraphIndent(shape, 0, { rightEmu: -1 })).toThrow(RangeError);
    expect(() => setParagraphIndent(shape, 0, { firstLineEmu: pt(-12) })).not.toThrow();
  });
});
