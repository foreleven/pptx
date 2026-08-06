// Picture cover placement — aspect-preserving crop inside the authored frame.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideImage,
  createPresentation,
  getShapeBounds,
  getShapeImageCrop,
  inches,
} from '../src/api/index.ts';
import { buildPng } from './lib/build-png.ts';

describe("fn API: addSlideImage fit='cover'", () => {
  it('keeps the requested geometry and writes a centered source crop', () => {
    const slide = addBlankSlide(createPresentation());
    const picture = addSlideImage(slide, buildPng(400, 200, [54, 89, 227]), {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
      fit: 'cover',
    });

    expect(getShapeBounds(picture)).toEqual({
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    expect(getShapeImageCrop(picture)).toEqual({ left: 0.25, top: 0, right: 0.25, bottom: 0 });
  });

  it('applies an explicit source crop before computing the cover crop', () => {
    const slide = addBlankSlide(createPresentation());
    const picture = addSlideImage(slide, buildPng(400, 200, [54, 89, 227]), {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
      fit: 'cover',
      crop: { left: 0.1 },
    });

    expect(getShapeImageCrop(picture)).toEqual({ left: 0.3, top: 0, right: 0.2, bottom: 0 });
  });
});
