// Slide-level shape authoring.

import type { Emu } from '../units.ts';
import {
  coverImageCrop,
  croppedImageSize,
  type ImageCrop,
} from '../../internal/drawingml/index.ts';
import {
  contentTypeForFormat,
  detectImageFormat,
  emptyRels,
  extensionForFormat,
  type ImageFormat,
  nextRelId,
  partName,
  readImagePixelSize,
} from '../../internal/opc/index.ts';
import { type ImageFit, fitImageRect } from './shape-image.ts';
import {
  REL_TYPES,
  type PresetShape,
  buildConnector,
  buildPicture,
  buildShape,
  buildTable,
  buildTextBox,
} from '../../internal/presentationml/index.ts';
import {
  INTERNAL_PACKAGE,
  SLIDE_PART_NAME,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { appendAndReturnNewShape, nextShapeId } from './_helpers.ts';
import { resolveDeckBodyTextColor } from './color-map.ts';
// ---------------------------------------------------------------------------
// Slide-level shape authoring.
//
// Each `addXxx` builds an XML element via an internal builder, appends
// it to the slide's `<p:spTree>`, commits, rebuilds the typed view, and
// returns the new SlideShapeData.

/**
 * Adds a free-form text box to the slide. Returns the new shape.
 *
 * The box is a plain rectangle with no fill or outline carrying one
 * paragraph with one run. The shape id is allocated as one more than
 * the current max id.
 */
export const addSlideTextBox = (
  slide: SlideData,
  opts: { x: Emu; y: Emu; w: Emu; h: Emu; text: string; name?: string },
): SlideShapeData => {
  const sp = buildTextBox({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    text: opts.text,
  });
  return appendAndReturnNewShape(slide, sp);
};

/**
 * Adds a preset shape (rectangle, ellipse, arrow, ...) to the slide.
 * Optional `text` seeds a single run.
 */
export const addSlideShape = (
  slide: SlideData,
  opts: {
    preset: PresetShape | string;
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    text?: string;
    /** Vertical text anchor (`t` / `ctr` / `b`). Horizontal alignment is set
     * separately via `setShapeAlignment` / `setParagraphAlignment`. */
    textAnchor?: 'ctr' | 't' | 'b';
    name?: string;
  },
): SlideShapeData => {
  const sp = buildShape({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    preset: opts.preset,
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    ...(opts.text !== undefined ? { text: opts.text } : {}),
    ...(opts.textAnchor !== undefined ? { textAnchor: opts.textAnchor } : {}),
  });
  return appendAndReturnNewShape(slide, sp);
};

/** Adds a straight-line connector between two points. */
export const addSlideLine = (
  slide: SlideData,
  opts: {
    from: { x: Emu; y: Emu };
    to: { x: Emu; y: Emu };
    color?: string;
    widthEmu?: number;
    name?: string;
  },
): SlideShapeData => {
  const cxn = buildConnector({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    from: opts.from,
    to: opts.to,
    ...(opts.color !== undefined ? { color: opts.color } : {}),
    ...(opts.widthEmu !== undefined ? { widthEmu: opts.widthEmu } : {}),
  });
  return appendAndReturnNewShape(slide, cxn);
};

/**
 * Adds a table to the slide. Cells render as plain text; `firstRow` /
 * `bandRow` flags drive PowerPoint's banded-header look unless options say
 * otherwise.
 *
 * Cell text color is baked from the deck's resolved body-text color so the
 * table stays readable even on templates with an inverted color map
 * (`bg1="dk1" tx1="lt1"`), where the default `tx1` text token would otherwise
 * paint the text the same as the background. Override per cell afterwards with
 * `setTableCellTextFormat`.
 */
export const addSlideTable = (
  slide: SlideData,
  opts: {
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    rows: ReadonlyArray<ReadonlyArray<string>>;
    colWidths?: ReadonlyArray<Emu>;
    rowHeights?: ReadonlyArray<Emu>;
    firstRow?: boolean;
    bandRow?: boolean;
    name?: string;
  },
): SlideShapeData => {
  const textColorHex = resolveDeckBodyTextColor(slide);
  const frame = buildTable({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    x: opts.x,
    y: opts.y,
    w: opts.w,
    h: opts.h,
    rows: opts.rows,
    ...(opts.colWidths !== undefined ? { colWidths: opts.colWidths } : {}),
    ...(opts.rowHeights !== undefined ? { rowHeights: opts.rowHeights } : {}),
    ...(opts.firstRow !== undefined ? { firstRow: opts.firstRow } : {}),
    ...(opts.bandRow !== undefined ? { bandRow: opts.bandRow } : {}),
    ...(textColorHex !== null ? { textColorHex } : {}),
  });
  return appendAndReturnNewShape(slide, frame);
};

/**
 * Adds a picture to the slide from raw bytes. Returns the new shape.
 *
 * Allocates a `/ppt/media/imageN.<ext>` part, registers a Content_Types
 * Default if the extension isn't yet covered, allocates a slide→image
 * rel, and appends a `<p:pic>` element to the slide's `<p:spTree>`.
 *
 * Format is detected from magic bytes; pass `opts.format` to override.
 *
 * `opts.fit` controls how the image fills the `w × h` box. `'fill'` (the
 * default) stretches to the exact box, ignoring aspect ratio. `'contain'`
 * scales the image to fit inside the box preserving its aspect ratio and
 * centers it using the supported image header. For malformed headers,
 * `'contain'` falls back to `'fill'` rather than erroring. `'cover'` preserves the box
 * and center-crops a measurable image to fill it without distortion; unknown
 * natural size falls back to the uncropped `'fill'` geometry.
 * `opts.crop` is applied to the source first, including negative outset values;
 * contain/cover then fit that remaining source region into the authored box.
 */
export const addSlideImage = (
  slide: SlideData,
  bytes: Uint8Array,
  opts: {
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    format?: ImageFormat;
    name?: string;
    fit?: ImageFit;
    crop?: ImageCrop | null;
  },
): SlideShapeData => {
  const pkg = slide[INTERNAL_PACKAGE];
  const format = opts.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error(
      'addSlideImage: could not detect image format. Pass options.format explicitly.',
    );
  }
  const contentType = contentTypeForFormat(format);
  const extension = extensionForFormat(format);

  let nextN = 1;
  const mediaPattern = /^\/ppt\/media\/image(\d+)\./;
  for (const p of pkg.parts) {
    const m = p.name.match(mediaPattern);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
    }
  }
  const newMediaName = partName(`/ppt/media/image${nextN}.${extension}`);

  const hasDefault = pkg.contentTypes.defaults.some((d) => d.extension.toLowerCase() === extension);
  if (!hasDefault) {
    pkg.contentTypes.defaults.push({ extension, contentType });
  }
  pkg.addPart(newMediaName, contentType, bytes);

  const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const newRId = nextRelId(rels.items.map((r) => r.id));
  rels.items.push({
    id: newRId,
    type: REL_TYPES.image,
    target: `../media/image${nextN}.${extension}`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], rels);

  const naturalSize = readImagePixelSize(bytes);
  const fittedNaturalSize = naturalSize ? croppedImageSize(naturalSize, opts.crop ?? null) : null;
  const rect = fitImageRect(
    { x: opts.x, y: opts.y, w: opts.w, h: opts.h },
    opts.fit ?? 'fill',
    fittedNaturalSize,
  );
  const crop =
    opts.fit === 'cover' && naturalSize !== null
      ? coverImageCrop({ width: opts.w, height: opts.h }, naturalSize, opts.crop ?? null)
      : (opts.crop ?? null);
  const pic = buildPicture({
    id: nextShapeId(slide),
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    rEmbed: newRId,
    x: rect.x,
    y: rect.y,
    w: rect.w,
    h: rect.h,
    ...(crop ? { crop } : {}),
  });
  return appendAndReturnNewShape(slide, pic);
};

export type MediaKind = 'video' | 'audio';

/**
 * Adds an embedded click-to-play audio or video frame backed by a poster image.
 * The media bytes are stored once and referenced by both the legacy media-file
 * relationship and the Office 2010 `p14:media` extension.
 */
export const addSlideMedia = (
  slide: SlideData,
  bytes: Uint8Array,
  opts: {
    kind: MediaKind;
    contentType: string;
    posterBytes: Uint8Array;
    posterFormat: ImageFormat;
    x: Emu;
    y: Emu;
    w: Emu;
    h: Emu;
    name?: string;
  },
): SlideShapeData => {
  const pkg = slide[INTERNAL_PACKAGE];
  const mediaExtension = extensionForMediaContentType(opts.kind, opts.contentType);
  const mediaNumber = nextMediaNumber(
    pkg.parts.map((part) => part.name),
    'media',
  );
  const mediaName = partName(`/ppt/media/media${mediaNumber}.${mediaExtension}`);
  ensureContentTypeDefault(pkg, mediaExtension, opts.contentType);
  pkg.addPart(mediaName, opts.contentType, bytes);

  const posterExtension = extensionForFormat(opts.posterFormat);
  const posterContentType = contentTypeForFormat(opts.posterFormat);
  const posterNumber = nextMediaNumber(
    pkg.parts.map((part) => part.name),
    'image',
  );
  const posterName = partName(`/ppt/media/image${posterNumber}.${posterExtension}`);
  ensureContentTypeDefault(pkg, posterExtension, posterContentType);
  pkg.addPart(posterName, posterContentType, opts.posterBytes);

  const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const usedIds = rels.items.map((rel) => rel.id);
  const fileRelId = nextRelId(usedIds);
  usedIds.push(fileRelId);
  const mediaRelId = nextRelId(usedIds);
  usedIds.push(mediaRelId);
  const posterRelId = nextRelId(usedIds);
  const mediaTarget = `../media/media${mediaNumber}.${mediaExtension}`;
  rels.items.push(
    {
      id: fileRelId,
      type: opts.kind === 'video' ? REL_TYPES.video : REL_TYPES.audio,
      target: mediaTarget,
      targetMode: 'Internal',
    },
    {
      id: mediaRelId,
      type: REL_TYPES.media,
      target: mediaTarget,
      targetMode: 'Internal',
    },
    {
      id: posterRelId,
      type: REL_TYPES.image,
      target: `../media/image${posterNumber}.${posterExtension}`,
      targetMode: 'Internal',
    },
  );
  pkg.setRels(slide[SLIDE_PART_NAME], rels);

  return appendAndReturnNewShape(
    slide,
    buildPicture({
      id: nextShapeId(slide),
      ...(opts.name !== undefined ? { name: opts.name } : {}),
      rEmbed: posterRelId,
      x: opts.x,
      y: opts.y,
      w: opts.w,
      h: opts.h,
      media: { kind: opts.kind, fileRelId, mediaRelId },
    }),
  );
};

function extensionForMediaContentType(kind: MediaKind, contentType: string): string {
  const normalized = contentType.trim().toLowerCase();
  const known: Readonly<Record<string, string>> = {
    'video/mp4': 'mp4',
    'video/quicktime': 'mov',
    'video/webm': 'webm',
    'audio/mpeg': 'mp3',
    'audio/mp4': 'm4a',
    'audio/wav': 'wav',
    'audio/x-wav': 'wav',
  };
  const extension = known[normalized];
  if (!normalized.startsWith(`${kind}/`) || extension === undefined) {
    throw new Error(
      `addSlideMedia: unsupported ${kind} content type ${JSON.stringify(contentType)}`,
    );
  }
  return extension;
}

function nextMediaNumber(partNames: readonly string[], prefix: 'image' | 'media'): number {
  const pattern = new RegExp(`^/ppt/media/${prefix}(\\d+)\\.`, 'u');
  let next = 1;
  for (const name of partNames) {
    const value = pattern.exec(name)?.[1];
    if (value !== undefined) next = Math.max(next, Number.parseInt(value, 10) + 1);
  }
  return next;
}

function ensureContentTypeDefault(
  pkg: SlideData[typeof INTERNAL_PACKAGE],
  extension: string,
  contentType: string,
): void {
  if (
    !pkg.contentTypes.defaults.some(
      (entry) => entry.extension.toLowerCase() === extension.toLowerCase(),
    )
  ) {
    pkg.contentTypes.defaults.push({ extension, contentType });
  }
}
