/**
 * Rewrites selected attribute values without reserializing the surrounding XML.
 *
 * The scanner only enters XML start tags and only replaces an exact attribute
 * name. Comments, CDATA, processing instructions, declarations, element text,
 * quote style, whitespace, namespace declarations, and attribute ordering are
 * preserved byte-for-byte. Replacement lookup is single-pass, so mappings such
 * as rId1 -> rId2 and rId2 -> rId3 cannot cascade.
 */
export const rewriteAttributeValuesLossless = (
  source: string,
  attributeName: string,
  replacements: ReadonlyMap<string, string>,
): string => {
  if (replacements.size === 0 || source.length === 0) return source;

  let output = '';
  let cursor = 0;
  while (cursor < source.length) {
    const tagStart = source.indexOf('<', cursor);
    if (tagStart < 0) return output + source.slice(cursor);
    output += source.slice(cursor, tagStart);

    if (source.startsWith('<!--', tagStart)) {
      const end = source.indexOf('-->', tagStart + 4);
      if (end < 0) return output + source.slice(tagStart);
      output += source.slice(tagStart, end + 3);
      cursor = end + 3;
      continue;
    }
    if (source.startsWith('<![CDATA[', tagStart)) {
      const end = source.indexOf(']]>', tagStart + 9);
      if (end < 0) return output + source.slice(tagStart);
      output += source.slice(tagStart, end + 3);
      cursor = end + 3;
      continue;
    }
    if (
      source.startsWith('<?', tagStart) ||
      source.startsWith('<!', tagStart) ||
      source.startsWith('</', tagStart)
    ) {
      const end = tagEnd(source, tagStart + 2);
      if (end < 0) return output + source.slice(tagStart);
      output += source.slice(tagStart, end + 1);
      cursor = end + 1;
      continue;
    }

    const end = tagEnd(source, tagStart + 1);
    if (end < 0) return output + source.slice(tagStart);
    output += rewriteStartTag(source.slice(tagStart, end + 1), attributeName, replacements);
    cursor = end + 1;
  }
  return output;
};

const tagEnd = (source: string, start: number): number => {
  let quote: '"' | "'" | null = null;
  for (let index = start; index < source.length; index += 1) {
    const char = source[index]!;
    if (quote !== null) {
      if (char === quote) quote = null;
    } else if (char === '"' || char === "'") {
      quote = char;
    } else if (char === '>') {
      return index;
    }
  }
  return -1;
};

const rewriteStartTag = (
  tag: string,
  attributeName: string,
  replacements: ReadonlyMap<string, string>,
): string => {
  let scan = 1;
  while (scan < tag.length && isNameChar(tag[scan]!)) scan += 1;
  let emittedThrough = 0;
  let output = '';
  while (scan < tag.length) {
    while (isXmlWhitespace(tag[scan])) scan += 1;
    if (tag[scan] === '/' || tag[scan] === '>' || tag[scan] === undefined) break;
    if (!isNameStart(tag[scan]!)) {
      scan += 1;
      continue;
    }
    const nameStart = scan;
    const nameEnd = nextNameEnd(tag, nameStart);
    scan = nameEnd;
    while (isXmlWhitespace(tag[scan])) scan += 1;
    if (tag[scan] !== '=') continue;
    scan += 1;
    while (isXmlWhitespace(tag[scan])) scan += 1;
    const quote = tag[scan];
    if (quote !== '"' && quote !== "'") continue;
    const valueStart = scan + 1;
    const valueEnd = tag.indexOf(quote, valueStart);
    if (valueEnd < 0) break;
    if (tag.slice(nameStart, nameEnd) === attributeName) {
      const replacement = replacements.get(tag.slice(valueStart, valueEnd));
      if (replacement !== undefined) {
        output += tag.slice(emittedThrough, valueStart) + replacement;
        emittedThrough = valueEnd;
      }
    }
    scan = valueEnd + 1;
  }
  return output + tag.slice(emittedThrough);
};

const nextNameEnd = (source: string, start: number): number => {
  let index = start + 1;
  while (index < source.length && isNameChar(source[index]!)) index += 1;
  return index;
};

const isNameStart = (char: string): boolean => /[A-Za-z_:]/u.test(char);
const isNameChar = (char: string): boolean => /[A-Za-z0-9_.:-]/u.test(char);
const isXmlWhitespace = (char: string | undefined): boolean =>
  char === ' ' || char === '\t' || char === '\r' || char === '\n';
