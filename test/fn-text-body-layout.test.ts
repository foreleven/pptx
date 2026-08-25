// Direct body-layout properties that have no paragraph-level equivalent.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getShapeTextAnchorCentering,
  getShapeTextCompatibilityLineSpacing,
  getShapeTextHorizontalOverflow,
  getShapeTextVerticalOverflow,
  inches,
  setShapeTextAnchorCentering,
  setShapeTextCompatibilityLineSpacing,
  setShapeTextHorizontalOverflow,
  setShapeTextVerticalOverflow,
} from '../src/api/index.ts';

describe('fn API: text body layout', () => {
  it('round-trips overflow and explicit compatibility booleans', () => {
    const presentation = createPresentation();
    const slide = addBlankSlide(presentation);
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: 'Body layout',
    });

    expect(getShapeTextHorizontalOverflow(shape)).toBeNull();
    expect(getShapeTextVerticalOverflow(shape)).toBeNull();
    expect(getShapeTextAnchorCentering(shape)).toBeNull();
    expect(getShapeTextCompatibilityLineSpacing(shape)).toBeNull();

    for (const value of ['overflow', 'clip'] as const) {
      setShapeTextHorizontalOverflow(shape, value);
      expect(getShapeTextHorizontalOverflow(shape)).toBe(value);
    }
    for (const value of ['overflow', 'ellipsis', 'clip'] as const) {
      setShapeTextVerticalOverflow(shape, value);
      expect(getShapeTextVerticalOverflow(shape)).toBe(value);
    }
    for (const value of [true, false] as const) {
      setShapeTextAnchorCentering(shape, value);
      expect(getShapeTextAnchorCentering(shape)).toBe(value);
      setShapeTextCompatibilityLineSpacing(shape, value);
      expect(getShapeTextCompatibilityLineSpacing(shape)).toBe(value);
    }

    setShapeTextHorizontalOverflow(shape, null);
    setShapeTextVerticalOverflow(shape, null);
    setShapeTextAnchorCentering(shape, null);
    setShapeTextCompatibilityLineSpacing(shape, null);
    expect(getShapeTextHorizontalOverflow(shape)).toBeNull();
    expect(getShapeTextVerticalOverflow(shape)).toBeNull();
    expect(getShapeTextAnchorCentering(shape)).toBeNull();
    expect(getShapeTextCompatibilityLineSpacing(shape)).toBeNull();
  });
});
