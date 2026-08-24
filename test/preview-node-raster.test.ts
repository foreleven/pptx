// Smoke tests for the Node rasterization path (`@office-kit/pptx-preview/node`).
//
// The pipeline: renderSlideToSvg (svg text mode) → resvg → PNG / RGBA.
// We verify dimensions, the PNG magic bytes, that the output is not entirely
// blank white, and that two renders of the same slide are byte-identical.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideImage,
  addSlideShape,
  addSlideTextBox,
  findSlideLayout,
  getSlideSize,
  inches,
  loadPresentation,
  setShapeFill,
  setShapeImageBiLevel,
  setShapeImageFill,
  setShapeRunFormat,
  setShapeStroke,
} from '../src/api/index.ts';
import { FONT_DIR, renderSlideToImage, renderSlideToRgba } from '../packages/preview/src/node.ts';
import { buildPng } from './lib/build-png.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const buildTestSlide = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout not found');
  const slide = addSlide(pres, { layout });
  // Colored background rect so the rendered image is not entirely white.
  const bg = addSlideShape(slide, {
    preset: 'rect',
    x: inches(0),
    y: inches(0),
    w: inches(10),
    h: inches(7.5),
  });
  setShapeFill(bg, '#3B82F6');
  // A text box for additional content variety.
  addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(8),
    h: inches(1),
    text: 'raster test',
  });
  return { pres, slide };
};

// PNG magic signature: first 4 bytes are 89 50 4E 47.
const isPng = (bytes: Uint8Array): boolean =>
  bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47;

describe('renderSlideToRgba (Node)', () => {
  it('returns image with expected pixel dimensions and correct aspect ratio', async () => {
    const { pres, slide } = await buildTestSlide();
    const targetWidth = 320;
    const { image } = renderSlideToRgba(pres, slide, { width: targetWidth });

    expect(image.width).toBe(targetWidth);

    // The aspect ratio should match the slide (16:9 ≈ 1.777).
    const slideSize = getSlideSize(pres)!;
    const expectedHeight = Math.round((targetWidth * slideSize.height) / slideSize.width);
    // Allow ±1 pixel rounding difference between integer rounding approaches.
    expect(Math.abs(image.height - expectedHeight)).toBeLessThanOrEqual(1);

    // RGBA: 4 bytes per pixel.
    expect(image.data.length).toBe(image.width * image.height * 4);
  });

  it('the PNG output starts with the PNG signature bytes', async () => {
    const { pres, slide } = await buildTestSlide();
    const { png } = renderSlideToRgba(pres, slide, { width: 320 });
    expect(isPng(png)).toBe(true);
  });

  it('the RGBA buffer is not entirely white/blank', async () => {
    const { pres, slide } = await buildTestSlide();
    const { image } = renderSlideToRgba(pres, slide, { width: 320 });
    // Count pixels that differ from white (R=255, G=255, B=255).
    let nonWhite = 0;
    for (let i = 0; i < image.data.length; i += 4) {
      if (image.data[i] !== 255 || image.data[i + 1] !== 255 || image.data[i + 2] !== 255) {
        nonWhite++;
      }
    }
    expect(nonWhite).toBeGreaterThan(0);
  });

  it('paints an image-filled roundRect inside its geometry', async () => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('Blank layout not found');
    const slide = addSlide(pres, { layout });
    const shape = addSlideShape(slide, {
      preset: 'roundRect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
    });
    setShapeImageFill(shape, buildPng(8, 4, [38, 52, 82]), { format: 'png' });
    setShapeStroke(shape, { color: '#35B9C6', widthEmu: 19_050 });

    const { image } = renderSlideToRgba(pres, slide, { width: 320 });
    const centerX = 64;
    const centerY = 48;
    const offset = (centerY * image.width + centerX) * 4;
    expect(Array.from(image.data.slice(offset, offset + 4))).toEqual([38, 52, 82, 255]);
  });

  it('applies bi-level threshold to luminance instead of individual RGB channels', async () => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('Blank layout not found');
    const slide = addSlide(pres, { layout });
    const picture = addSlideImage(slide, buildPng(8, 4, [242, 107, 91]), {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      format: 'png',
    });
    setShapeImageBiLevel(picture, 0.65);

    const { image } = renderSlideToRgba(pres, slide, { width: 320 });
    const offset = (48 * image.width + 64) * 4;
    const pixel = Array.from(image.data.slice(offset, offset + 4));
    // Coral has luminance below 65%, so DrawingML biLevel maps it to black.
    // Per-channel thresholding would incorrectly preserve it as bright red.
    expect(pixel).toEqual([0, 0, 0, 255]);
  });

  it('render is byte-identical when called twice (determinism)', async () => {
    const { pres, slide } = await buildTestSlide();
    const opts = { width: 320 };
    const first = renderSlideToRgba(pres, slide, opts).png;
    const second = renderSlideToRgba(pres, slide, opts).png;
    // Compare via string encoding to get a useful diff on failure.
    expect(Buffer.from(first).toString('hex')).toBe(Buffer.from(second).toString('hex'));
  });

  it('renders distinct CJK glyphs instead of one missing-glyph box', async () => {
    /**
     * Renders one glyph per fresh copy of the same slide so geometry and styling stay
     * identical. If the Node renderer lacks a CJK font, resvg paints every character
     * as the same tofu box and all four PNG byte sequences collapse to one value.
     */
    const renderGlyph = async (text: string): Promise<string> => {
      const pres = await loadPresentation(await readFile(fixturePath));
      const layout = findSlideLayout(pres, 'Blank');
      if (!layout) throw new Error('Blank layout not found');
      const slide = addSlide(pres, { layout });
      const box = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(2),
        h: inches(1),
        text,
      });
      setShapeRunFormat(box, 0, 0, {
        font: 'Arial',
        fontEastAsian: 'Noto Sans CJK SC',
        size: 48,
      });
      return Buffer.from(renderSlideToRgba(pres, slide, { width: 320 }).png).toString('hex');
    };

    const renders = await Promise.all(['中', '文', '字', '体'].map(renderGlyph));
    expect(new Set(renders).size).toBe(renders.length);
  });
});

describe('renderSlideToImage (Node)', () => {
  it('returns a valid non-empty PNG', async () => {
    const { pres, slide } = await buildTestSlide();
    const png = renderSlideToImage(pres, slide, { width: 320 });
    expect(isPng(png)).toBe(true);
    expect(png.length).toBeGreaterThan(100);
  });

  it('output matches the png from renderSlideToRgba for the same input', async () => {
    // Both functions call the same internal rasterize(); their PNG outputs
    // should be identical for the same parameters.
    const { pres, slide } = await buildTestSlide();
    const opts = { width: 320 };
    const fromImage = renderSlideToImage(pres, slide, opts);
    const { png: fromRgba } = renderSlideToRgba(pres, slide, opts);
    expect(Buffer.from(fromImage).toString('hex')).toBe(Buffer.from(fromRgba).toString('hex'));
  });
});

describe('renderSlideToImage — rasterizer font options', () => {
  it('accepts extra fontFiles without changing an all-Latin render', async () => {
    // The strongest deterministic assertion available without shipping a CJK
    // fixture font: an extra face that adds no new glyph coverage must not
    // perturb the output, proving the option is plumbed through resvg rather
    // than replacing the bundled set.
    const { pres, slide } = await buildTestSlide();
    const baseline = renderSlideToImage(pres, slide, { width: 320 });
    const withExtra = renderSlideToImage(pres, slide, {
      width: 320,
      fontFiles: [`${FONT_DIR}Carlito-Regular.ttf`],
    });
    expect(Buffer.from(withExtra).toString('hex')).toBe(Buffer.from(baseline).toString('hex'));
  });

  it('tolerates a fontFiles path that does not exist (resvg skips it)', async () => {
    const { pres, slide } = await buildTestSlide();
    const png = renderSlideToImage(pres, slide, {
      width: 320,
      fontFiles: ['/nonexistent/no-such-font.ttf'],
    });
    expect(isPng(png)).toBe(true);
  });

  it('loadSystemFonts: true still produces a valid PNG', async () => {
    const { pres, slide } = await buildTestSlide();
    const png = renderSlideToImage(pres, slide, { width: 320, loadSystemFonts: true });
    expect(isPng(png)).toBe(true);
  });
});
