// Shared image-crop primitives for picture shapes and shape image fills.

import { NS, type XmlElement, attr, elem, qname } from '../xml/index.ts';

/** Fractions removed from each source-image edge before it fills its target box; negatives add outset. */
export interface ImageCrop {
  readonly left?: number;
  readonly top?: number;
  readonly right?: number;
  readonly bottom?: number;
}

interface ImageSize {
  readonly width: number;
  readonly height: number;
}

type NormalizedImageCrop = Required<ImageCrop>;

/** Validate one source rectangle and materialize omitted edges as zero. */
const normalizeImageCrop = (crop: ImageCrop | null = null): NormalizedImageCrop => {
  const normalized = {
    left: crop?.left ?? 0,
    top: crop?.top ?? 0,
    right: crop?.right ?? 0,
    bottom: crop?.bottom ?? 0,
  };
  for (const value of Object.values(normalized)) {
    if (!Number.isFinite(value)) throw new RangeError(`crop fraction must be finite, got ${value}`);
  }
  if (normalized.left + normalized.right >= 1 || normalized.top + normalized.bottom >= 1) {
    throw new RangeError('opposite crop fractions must sum to less than 1');
  }
  return normalized;
};

/** Omit zero edges and collapse an unchanged source rectangle to `null`. */
const sparseImageCrop = (crop: NormalizedImageCrop): ImageCrop | null => {
  const result: ImageCrop = {
    ...(crop.left === 0 ? {} : { left: crop.left }),
    ...(crop.top === 0 ? {} : { top: crop.top }),
    ...(crop.right === 0 ? {} : { right: crop.right }),
    ...(crop.bottom === 0 ? {} : { bottom: crop.bottom }),
  };
  return Object.keys(result).length === 0 ? null : result;
};

/** Return the source dimensions after the authored inset/outset is applied. */
export const croppedImageSize = (source: ImageSize, crop: ImageCrop | null = null): ImageSize => {
  const normalized = normalizeImageCrop(crop);
  return {
    width: source.width * (1 - normalized.left - normalized.right),
    height: source.height * (1 - normalized.top - normalized.bottom),
  };
};

/** Compute the centered source crop needed after an authored crop so the remainder covers the target. */
export const coverImageCrop = (
  target: ImageSize,
  source: ImageSize,
  crop: ImageCrop | null = null,
): ImageCrop | null => {
  const normalized = normalizeImageCrop(crop);
  if (target.width <= 0 || target.height <= 0 || source.width <= 0 || source.height <= 0) {
    return sparseImageCrop(normalized);
  }
  const croppedSource = croppedImageSize(source, normalized);
  const targetAspect = target.width / target.height;
  const sourceAspect = croppedSource.width / croppedSource.height;
  if (Math.abs(targetAspect - sourceAspect) < Number.EPSILON) return sparseImageCrop(normalized);
  const widthFraction = 1 - normalized.left - normalized.right;
  const heightFraction = 1 - normalized.top - normalized.bottom;
  if (sourceAspect > targetAspect) {
    const side = (1 - targetAspect / sourceAspect) / 2;
    return sparseImageCrop({
      ...normalized,
      left: normalized.left + side * widthFraction,
      right: normalized.right + side * widthFraction,
    });
  }
  const side = (1 - sourceAspect / targetAspect) / 2;
  return sparseImageCrop({
    ...normalized,
    top: normalized.top + side * heightFraction,
    bottom: normalized.bottom + side * heightFraction,
  });
};

/** Build the DrawingML `<a:srcRect>` used by both picture and shape-fill XML. */
export const buildImageSourceRectangle = (crop: ImageCrop): XmlElement => {
  const checked = normalizeImageCrop(crop);
  const pairs: ReadonlyArray<['l' | 't' | 'r' | 'b', number]> = [
    ['l', checked.left],
    ['t', checked.top],
    ['r', checked.right],
    ['b', checked.bottom],
  ];
  const attrs = pairs.flatMap(([name, value]) => {
    return value === 0 ? [] : [attr(qname('', name, ''), String(Math.round(value * 100000)))];
  });
  return elem(qname('a', 'srcRect', NS.dml), { attrs });
};
