// Image format detection from raw bytes. Used by the picture-replacement
// path to pick the right content type and file extension without forcing the
// caller to spell them out.
//
// Detection is by magic bytes — sufficient for the formats PowerPoint
// accepts. We do not run a full validator (that's the image library's job).

export type ImageFormat = 'png' | 'jpeg' | 'gif' | 'bmp' | 'tiff' | 'webp' | 'svg';

const startsWith = (bytes: Uint8Array, signature: ReadonlyArray<number>): boolean => {
  if (bytes.length < signature.length) return false;
  for (let i = 0; i < signature.length; i++) {
    if (bytes[i] !== signature[i]) return false;
  }
  return true;
};

const decoder = new TextDecoder('utf-8', { fatal: false });

/**
 * Detects the image format from raw bytes. Returns `null` if no known
 * signature matches — callers should treat that as a hard error and surface
 * a message asking for a recognized format.
 */
export const detectImageFormat = (bytes: Uint8Array): ImageFormat | null => {
  if (startsWith(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return 'png';
  if (startsWith(bytes, [0xff, 0xd8, 0xff])) return 'jpeg';
  if (startsWith(bytes, [0x47, 0x49, 0x46, 0x38])) return 'gif';
  if (startsWith(bytes, [0x42, 0x4d])) return 'bmp';
  if (startsWith(bytes, [0x49, 0x49, 0x2a, 0x00])) return 'tiff';
  if (startsWith(bytes, [0x4d, 0x4d, 0x00, 0x2a])) return 'tiff';
  if (
    bytes.length >= 12 &&
    startsWith(bytes, [0x52, 0x49, 0x46, 0x46]) &&
    bytes[8] === 0x57 &&
    bytes[9] === 0x45 &&
    bytes[10] === 0x42 &&
    bytes[11] === 0x50
  ) {
    return 'webp';
  }
  // SVG is textual — sniff up to the first 1KB.
  const head = decoder.decode(bytes.subarray(0, Math.min(bytes.length, 1024)));
  if (/<svg[\s>]/.test(head)) return 'svg';
  return null;
};

/** Natural pixel dimensions of an image. */
export interface ImagePixelSize {
  readonly width: number;
  readonly height: number;
}

const readUint16Be = (bytes: Uint8Array, at: number): number => (bytes[at]! << 8) | bytes[at + 1]!;

const readUint16Le = (bytes: Uint8Array, at: number): number => bytes[at]! | (bytes[at + 1]! << 8);

const readUint32Be = (bytes: Uint8Array, at: number): number =>
  // `>>> 0` keeps the result an unsigned 32-bit int (a 4-byte PNG dimension
  // with the high bit set would otherwise read as negative).
  ((bytes[at]! << 24) | (bytes[at + 1]! << 16) | (bytes[at + 2]! << 8) | bytes[at + 3]!) >>> 0;

const readUint32Le = (bytes: Uint8Array, at: number): number =>
  (bytes[at]! | (bytes[at + 1]! << 8) | (bytes[at + 2]! << 16) | (bytes[at + 3]! << 24)) >>> 0;

/** Return only positive finite dimensions so malformed headers always fall back safely. */
const imageSize = (width: number, height: number): ImagePixelSize | null =>
  Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0
    ? { width, height }
    : null;

// PNG: the IHDR chunk is the first chunk and always at a fixed offset —
// 8-byte signature, 4-byte length, 4-byte "IHDR" tag, then width / height
// as big-endian uint32 (PNG spec §11.2.2).
const pngSize = (bytes: Uint8Array): ImagePixelSize | null => {
  if (bytes.length < 24) return null;
  const width = readUint32Be(bytes, 16);
  const height = readUint32Be(bytes, 20);
  if (width <= 0 || height <= 0) return null;
  return { width, height };
};

// JPEG: walk the marker segments until a Start-Of-Frame marker, whose
// payload carries the sample dimensions. SOF markers are 0xC0..0xCF except
// the non-frame markers 0xC4 (DHT), 0xC8 (JPG), 0xCC (DAC).
const JPEG_SOF_EXCLUDED = new Set([0xc4, 0xc8, 0xcc]);
const jpegSize = (bytes: Uint8Array): ImagePixelSize | null => {
  // Skip the SOI (0xFFD8); then each segment is 0xFF, marker, 2-byte length.
  let offset = 2;
  while (offset + 9 < bytes.length) {
    if (bytes[offset] !== 0xff) {
      offset++;
      continue;
    }
    const marker = bytes[offset + 1]!;
    // Padding / standalone markers (RSTn, SOI, EOI, TEM) carry no length.
    if (marker === 0xff || (marker >= 0xd0 && marker <= 0xd9) || marker === 0x01) {
      offset += 2;
      continue;
    }
    const segmentLength = readUint16Be(bytes, offset + 2);
    if (segmentLength < 2) return null;
    if (marker >= 0xc0 && marker <= 0xcf && !JPEG_SOF_EXCLUDED.has(marker)) {
      // SOF payload: 1-byte precision, 2-byte height, 2-byte width.
      const height = readUint16Be(bytes, offset + 5);
      const width = readUint16Be(bytes, offset + 7);
      if (width <= 0 || height <= 0) return null;
      return { width, height };
    }
    offset += 2 + segmentLength;
  }
  return null;
};

/** Read the GIF logical-screen width and height stored after its six-byte signature. */
const gifSize = (bytes: Uint8Array): ImagePixelSize | null =>
  bytes.length < 10 ? null : imageSize(readUint16Le(bytes, 6), readUint16Le(bytes, 8));

/** Read Windows or OS/2 BMP dimensions from the DIB header. */
const bmpSize = (bytes: Uint8Array): ImagePixelSize | null => {
  if (bytes.length < 26) return null;
  const dibSize = readUint32Le(bytes, 14);
  if (dibSize === 12) return imageSize(readUint16Le(bytes, 18), readUint16Le(bytes, 20));
  const unsignedWidth = readUint32Le(bytes, 18);
  const unsignedHeight = readUint32Le(bytes, 22);
  const width = unsignedWidth > 0x7fffffff ? unsignedWidth - 0x1_0000_0000 : unsignedWidth;
  const height = unsignedHeight > 0x7fffffff ? unsignedHeight - 0x1_0000_0000 : unsignedHeight;
  return imageSize(Math.abs(width), Math.abs(height));
};

/** Read scalar ImageWidth/ImageLength tags from the first TIFF image-file directory. */
const tiffSize = (bytes: Uint8Array): ImagePixelSize | null => {
  if (bytes.length < 10) return null;
  const littleEndian = startsWith(bytes, [0x49, 0x49, 0x2a, 0x00]);
  const bigEndian = startsWith(bytes, [0x4d, 0x4d, 0x00, 0x2a]);
  if (!littleEndian && !bigEndian) return null;
  const read16 = littleEndian ? readUint16Le : readUint16Be;
  const read32 = littleEndian ? readUint32Le : readUint32Be;
  const ifdOffset = read32(bytes, 4);
  if (ifdOffset + 2 > bytes.length) return null;
  const count = read16(bytes, ifdOffset);
  let width: number | null = null;
  let height: number | null = null;
  for (let index = 0; index < count; index++) {
    const entry = ifdOffset + 2 + index * 12;
    if (entry + 12 > bytes.length) return null;
    const tag = read16(bytes, entry);
    if (tag !== 256 && tag !== 257) continue;
    const type = read16(bytes, entry + 2);
    const values = read32(bytes, entry + 4);
    if (values !== 1 || (type !== 3 && type !== 4)) continue;
    const value = type === 3 ? read16(bytes, entry + 8) : read32(bytes, entry + 8);
    if (tag === 256) width = value;
    else height = value;
  }
  return width === null || height === null ? null : imageSize(width, height);
};

/** Read dimensions from the VP8X, VP8L, or VP8 payload header inside a WebP RIFF. */
const webpSize = (bytes: Uint8Array): ImagePixelSize | null => {
  if (bytes.length < 30) return null;
  const chunk = decoder.decode(bytes.subarray(12, 16));
  if (chunk === 'VP8X') {
    const width = 1 + bytes[24]! + (bytes[25]! << 8) + (bytes[26]! << 16);
    const height = 1 + bytes[27]! + (bytes[28]! << 8) + (bytes[29]! << 16);
    return imageSize(width, height);
  }
  if (chunk === 'VP8L' && bytes[20] === 0x2f) {
    const width = 1 + bytes[21]! + ((bytes[22]! & 0x3f) << 8);
    const height = 1 + (bytes[22]! >> 6) + (bytes[23]! << 2) + ((bytes[24]! & 0x0f) << 10);
    return imageSize(width, height);
  }
  if (chunk === 'VP8 ' && bytes[23] === 0x9d && bytes[24] === 0x01 && bytes[25] === 0x2a) {
    return imageSize(readUint16Le(bytes, 26) & 0x3fff, readUint16Le(bytes, 28) & 0x3fff);
  }
  return null;
};

const SVG_UNIT_TO_PX: Readonly<Record<string, number>> = {
  px: 1,
  pt: 96 / 72,
  pc: 16,
  in: 96,
  cm: 96 / 2.54,
  mm: 96 / 25.4,
};

/** Convert one absolute SVG length into CSS pixels; percentages need a viewBox and return null. */
const svgLength = (value: string): number | null => {
  const match = /^\s*(\d+(?:\.\d*)?|\.\d+)\s*(px|pt|pc|in|cm|mm)?\s*$/i.exec(value);
  if (!match?.[1]) return null;
  const amount = Number(match[1]);
  const scale = SVG_UNIT_TO_PX[match[2]?.toLowerCase() ?? 'px'];
  return Number.isFinite(amount) && scale !== undefined ? amount * scale : null;
};

/** Read an SVG's absolute width/height, falling back to the root viewBox aspect. */
const svgSize = (bytes: Uint8Array): ImagePixelSize | null => {
  const head = decoder.decode(bytes.subarray(0, Math.min(bytes.length, 4096)));
  const root = /<svg\b[^>]*>/is.exec(head)?.[0];
  if (!root) return null;
  const widthValue = /\bwidth\s*=\s*["']([^"']+)["']/i.exec(root)?.[1];
  const heightValue = /\bheight\s*=\s*["']([^"']+)["']/i.exec(root)?.[1];
  const width = widthValue === undefined ? null : svgLength(widthValue);
  const height = heightValue === undefined ? null : svgLength(heightValue);
  if (width !== null && height !== null) return imageSize(width, height);
  const viewBox = /\bviewBox\s*=\s*["']([^"']+)["']/i
    .exec(root)?.[1]
    ?.trim()
    .split(/[\s,]+/u)
    .map(Number);
  return viewBox?.length === 4 ? imageSize(viewBox[2]!, viewBox[3]!) : null;
};

/**
 * Reads an image's natural dimensions from the supported PNG, JPEG, GIF, BMP,
 * TIFF, WebP, or SVG header. Returns `null` for truncated or malformed bytes,
 * letting callers preserve the historical fill fallback instead of failing.
 */
export const readImagePixelSize = (bytes: Uint8Array): ImagePixelSize | null => {
  const format = detectImageFormat(bytes);
  if (format === 'png') return pngSize(bytes);
  if (format === 'jpeg') return jpegSize(bytes);
  if (format === 'gif') return gifSize(bytes);
  if (format === 'bmp') return bmpSize(bytes);
  if (format === 'tiff') return tiffSize(bytes);
  if (format === 'webp') return webpSize(bytes);
  if (format === 'svg') return svgSize(bytes);
  return null;
};

/**
 * Returns the conventional file-extension token (no leading dot) for the
 * given format. `jpeg` maps to `jpg` because that's what PowerPoint emits.
 */
export const extensionForFormat = (format: ImageFormat): string => {
  switch (format) {
    case 'jpeg':
      return 'jpg';
    case 'png':
      return 'png';
    case 'gif':
      return 'gif';
    case 'bmp':
      return 'bmp';
    case 'tiff':
      return 'tiff';
    case 'webp':
      return 'webp';
    case 'svg':
      return 'svg';
  }
};

/** Returns the IANA media type for the given image format. */
export const contentTypeForFormat = (format: ImageFormat): string => {
  switch (format) {
    case 'png':
      return 'image/png';
    case 'jpeg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'bmp':
      return 'image/bmp';
    case 'tiff':
      return 'image/tiff';
    case 'webp':
      return 'image/webp';
    case 'svg':
      return 'image/svg+xml';
  }
};
