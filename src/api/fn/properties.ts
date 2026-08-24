// Core properties and extended properties.

import { emptyRels, nextRelId, partName } from '../../internal/opc/index.ts';
import { REL_TYPES } from '../../internal/presentationml/index.ts';
import {
  type XmlElement,
  type XmlNode,
  attr,
  childElements,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
  serializeXml,
  text as textNode,
  textContent,
} from '../../internal/xml/index.ts';
import { INTERNAL_PACKAGE, type PresentationData } from '../_internal-symbols.ts';
import { decode, encode } from './_helpers.ts';

// ---------------------------------------------------------------------------
// Core properties (`/docProps/core.xml`).

const NS_CORE_PROPS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const NS_DC = 'http://purl.org/dc/elements/1.1/';
const NS_DCTERMS = 'http://purl.org/dc/terms/';
const NS_XSI = 'http://www.w3.org/2001/XMLSchema-instance';
const CORE_PROPS_PART_NAME = partName('/docProps/core.xml');

/**
 * Document-level metadata from `/docProps/core.xml` (Open Packaging
 * Conventions). Surfaces the fields PowerPoint, Keynote, and
 * everyone else exchange via OPC core-properties — these are the
 * values shown in PowerPoint's "File › Properties" / "Info" panel.
 */
export interface CoreProperties {
  readonly title: string | null;
  readonly subject: string | null;
  readonly creator: string | null;
  readonly keywords: string | null;
  readonly description: string | null;
  readonly lastModifiedBy: string | null;
  readonly revision: string | null;
  /** ISO-8601 timestamp string when set; `null` otherwise. */
  readonly created: string | null;
  /** ISO-8601 timestamp string when set; `null` otherwise. */
  readonly modified: string | null;
  readonly category: string | null;
}

/**
 * Reads `/docProps/core.xml`. Returns `null` when the package has
 * no core-properties part. Each field is `null` when the
 * corresponding element is absent or empty.
 */
/**
 * Convenience: bumps core-properties' `cp:revision` by one (treating
 * an unset / unparseable value as 0). Returns the new revision
 * number. Useful right before `savePresentation` so consumers can
 * tell decks apart.
 */
export const incrementRevision = (pres: PresentationData): number => {
  const props = getCoreProperties(pres);
  const current =
    props?.revision === null || props?.revision === undefined
      ? 0
      : Number.parseInt(props.revision, 10);
  const next = (Number.isFinite(current) ? current : 0) + 1;
  setCoreProperties(pres, { revision: String(next) });
  return next;
};

/**
 * Convenience: writes `new Date().toISOString()` to
 * `dcterms:modified`. Useful right before `savePresentation` so
 * "last edited" shows the actual save time. Pass an explicit
 * `Date` to set a specific value.
 */
export const touchModified = (pres: PresentationData, at: Date = new Date()): void => {
  setCoreProperties(pres, { modified: at.toISOString() });
};

/**
 * Convenience: the timestamp from core-properties' `dcterms:created`,
 * parsed as a `Date`. Returns `null` when no created field is set
 * or the value isn't a recognizable W3C-DTF / ISO-8601 string.
 */
export const getPresentationCreated = (pres: PresentationData): Date | null => {
  const props = getCoreProperties(pres);
  if (!props || props.created === null) return null;
  const d = new Date(props.created);
  return Number.isFinite(d.getTime()) ? d : null;
};

/**
 * Convenience: the timestamp from core-properties' `dcterms:modified`,
 * parsed as a `Date`. Returns `null` when no modified field is set
 * or the value isn't a recognizable W3C-DTF / ISO-8601 string.
 */
export const getPresentationModified = (pres: PresentationData): Date | null => {
  const props = getCoreProperties(pres);
  if (!props || props.modified === null) return null;
  const d = new Date(props.modified);
  return Number.isFinite(d.getTime()) ? d : null;
};

export const getCoreProperties = (pres: PresentationData): CoreProperties | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const part = pkg.getPart(CORE_PROPS_PART_NAME);
  if (!part) return null;
  const root = parseXml(decode(part.data)).root;
  const read = (uri: string, local: string): string | null => {
    const el = firstChildElement(root, qname('', local, uri));
    if (!el) return null;
    let s = '';
    for (const c of el.children) if (c.kind === 'text') s += c.data;
    return s.length === 0 ? null : s;
  };
  return {
    title: read(NS_DC, 'title'),
    subject: read(NS_DC, 'subject'),
    creator: read(NS_DC, 'creator'),
    keywords: read(NS_CORE_PROPS, 'keywords'),
    description: read(NS_DC, 'description'),
    lastModifiedBy: read(NS_CORE_PROPS, 'lastModifiedBy'),
    revision: read(NS_CORE_PROPS, 'revision'),
    created: read(NS_DCTERMS, 'created'),
    modified: read(NS_DCTERMS, 'modified'),
    category: read(NS_CORE_PROPS, 'category'),
  };
};

const CORE_PROPS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml';

const CORE_PROP_FIELDS: ReadonlyArray<{
  key: keyof CoreProperties;
  uri: string;
  prefix: string;
  local: string;
  w3cdtf?: true;
}> = [
  { key: 'title', uri: NS_DC, prefix: 'dc', local: 'title' },
  { key: 'subject', uri: NS_DC, prefix: 'dc', local: 'subject' },
  { key: 'creator', uri: NS_DC, prefix: 'dc', local: 'creator' },
  { key: 'keywords', uri: NS_CORE_PROPS, prefix: 'cp', local: 'keywords' },
  { key: 'description', uri: NS_DC, prefix: 'dc', local: 'description' },
  { key: 'lastModifiedBy', uri: NS_CORE_PROPS, prefix: 'cp', local: 'lastModifiedBy' },
  { key: 'revision', uri: NS_CORE_PROPS, prefix: 'cp', local: 'revision' },
  { key: 'created', uri: NS_DCTERMS, prefix: 'dcterms', local: 'created', w3cdtf: true },
  { key: 'modified', uri: NS_DCTERMS, prefix: 'dcterms', local: 'modified', w3cdtf: true },
  { key: 'category', uri: NS_CORE_PROPS, prefix: 'cp', local: 'category' },
];

const buildEmptyCorePropsRoot = (): XmlElement => {
  const prefixDecls = new Map<string, string>([
    ['cp', NS_CORE_PROPS],
    ['dc', NS_DC],
    ['dcterms', NS_DCTERMS],
    ['xsi', NS_XSI],
  ]);
  return {
    kind: 'element',
    name: qname('cp', 'coreProperties', NS_CORE_PROPS),
    attrs: [],
    prefixDecls,
    children: [],
  };
};

/** Mark an OPC date as W3C-DTF, which PowerPoint requires even though the element text is already ISO-8601. */
const setW3cdtfType = (root: XmlElement, dateElement: XmlElement): void => {
  root.prefixDecls.set('xsi', NS_XSI);
  const typeIndex = dateElement.attrs.findIndex(
    (candidate) => candidate.name.namespaceURI === NS_XSI && candidate.name.localName === 'type',
  );
  const type = attr(qname('xsi', 'type', NS_XSI), 'dcterms:W3CDTF');
  if (typeIndex >= 0) dateElement.attrs[typeIndex] = type;
  else dateElement.attrs.push(type);
};

/**
 * Writes selected fields on `/docProps/core.xml`. Unspecified fields
 * are left as-is; pass `null` to clear a field that's currently set.
 * Bootstraps the part (and the `/_rels/.rels` entry + content-type
 * override) if the package didn't have one.
 *
 * Setting `created` / `modified` requires an ISO-8601 timestamp string
 * (e.g. `'2026-05-15T12:34:56Z'`). These fields are serialized with
 * `xsi:type="dcterms:W3CDTF"`; PowerPoint treats the otherwise-valid
 * core-properties part as repairable when that required type is absent.
 */
export const setCoreProperties = (
  pres: PresentationData,
  values: Partial<CoreProperties>,
): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  let part = pkg.getPart(CORE_PROPS_PART_NAME);
  let root: XmlElement;
  let doc: ReturnType<typeof parseXml>;
  if (part) {
    doc = parseXml(decode(part.data));
    root = doc.root;
  } else {
    root = buildEmptyCorePropsRoot();
    doc = { kind: 'document', decl: null, prolog: [], root, epilog: [] };
  }

  for (const field of CORE_PROP_FIELDS) {
    if (!(field.key in values)) continue;
    const value = values[field.key] ?? null;
    const name = qname(field.prefix, field.local, field.uri);
    const existing = firstChildElement(root, name);
    if (value === null) {
      if (existing) {
        existing.children = [];
      }
      continue;
    }
    let target: XmlElement;
    if (existing) {
      existing.children = [textNode(value)];
      target = existing;
    } else {
      target = elem(name, { children: [textNode(value)] });
      root.children.push(target);
    }
    if (field.w3cdtf) setW3cdtfType(root, target);
  }

  const bytes = encode(serializeXml(doc));
  if (part) {
    part.data = bytes;
    return;
  }

  // Bootstrap: register override, add part, wire root rel.
  pkg.contentTypes.overrides.push({
    partName: CORE_PROPS_PART_NAME,
    contentType: CORE_PROPS_CONTENT_TYPE,
  });
  pkg.addPart(CORE_PROPS_PART_NAME, CORE_PROPS_CONTENT_TYPE, bytes);

  const rootRels = pkg.rootRels() ?? emptyRels();
  const rId = nextRelId(rootRels.items.map((r) => r.id));
  rootRels.items.push({
    id: rId,
    type: REL_TYPES.coreProperties,
    target: 'docProps/core.xml',
    targetMode: 'Internal',
  });
  pkg.setRootRels(rootRels);
};

// ---------------------------------------------------------------------------
// Extended properties (`/docProps/app.xml`).

const NS_EXT_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const NS_VARIANT_TYPES = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
const EXT_PROPS_PART_NAME = partName('/docProps/app.xml');

export interface ExtendedPropertyHeadingPair {
  readonly heading: string;
  readonly count: number;
}

const parseBoundedXsdInteger = (
  value: string,
  minimum: number,
  maximum: number,
  label: string,
): number => {
  const lexical = value
    .replaceAll('\t', ' ')
    .replaceAll('\n', ' ')
    .replaceAll('\r', ' ')
    .replace(/ +/gu, ' ')
    .replace(/^ | $/gu, '');
  if (!/^[+-]?[0-9]+$/u.test(lexical)) {
    throw new Error(`${label} is not an XML Schema integer`);
  }
  const parsed = Number(lexical);
  if (!Number.isSafeInteger(parsed) || parsed < minimum || parsed > maximum) {
    throw new Error(`${label} is outside its XML Schema integer range`);
  }
  return parsed;
};

const parseXsdInt = (value: string, label: string): number =>
  parseBoundedXsdInteger(value, -2_147_483_648, 2_147_483_647, label);

const parseXsdUnsignedInt = (value: string, label: string): number =>
  parseBoundedXsdInteger(value, 0, 4_294_967_295, label);

/**
 * Presentation-oriented fields from `/docProps/app.xml`
 * (extended-properties / "app props"). PowerPoint exposes these
 * under File › Info / Properties as the "Origin" and "Related
 * People" groups.
 *
 * The presentation-oriented count fields and the paired vectors used by
 * PowerPoint's Properties UI are exposed verbatim. They are metadata, not a
 * live view: consumers that edit slides must update them explicitly when
 * they need the values to reflect the new deck.
 */
export interface ExtendedProperties {
  readonly application: string | null;
  readonly appVersion: string | null;
  readonly company: string | null;
  readonly manager: string | null;
  readonly presentationFormat: string | null;
  readonly hyperlinkBase: string | null;
  readonly slides: number | null;
  readonly notes: number | null;
  readonly hiddenSlides: number | null;
  readonly headingPairs: readonly ExtendedPropertyHeadingPair[] | null;
  readonly titlesOfParts: readonly string[] | null;
}

/**
 * Reads `/docProps/app.xml`. Returns `null` if the package has no
 * extended-properties part. Each field is `null` when the
 * corresponding element is absent or empty.
 */
export const getExtendedProperties = (pres: PresentationData): ExtendedProperties | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const part = pkg.getPart(EXT_PROPS_PART_NAME);
  if (!part) return null;
  const root = parseXml(decode(part.data)).root;
  const element = (local: string): XmlElement | null =>
    firstChildElement(root, qname('', local, NS_EXT_PROPS));
  const read = (local: string): string | null => {
    const el = element(local);
    if (!el) return null;
    const s = textContent(el);
    return s.length === 0 ? null : s;
  };
  const readInteger = (local: string): number | null => {
    const field = element(local);
    if (field === null) return null;
    const value = parseXsdInt(textContent(field), `getExtendedProperties: ${local}`);
    if (value < 0) throw new Error(`getExtendedProperties: ${local} must be non-negative`);
    return value;
  };
  const readVector = (
    local: string,
    baseType: 'variant' | 'lpstr',
  ): readonly XmlElement[] | null => {
    const container = element(local);
    if (!container) return null;
    const vector = firstChildElement(container, qname('vt', 'vector', NS_VARIANT_TYPES));
    if (!vector || getAttrValue(vector, qname('', 'baseType', '')) !== baseType) {
      throw new Error(`getExtendedProperties: ${local} has an invalid vt:vector`);
    }
    const entries = childElements(vector);
    const size = getAttrValue(vector, qname('', 'size', ''));
    if (
      size === null ||
      parseXsdUnsignedInt(size, `getExtendedProperties: ${local} vt:vector size`) !==
        entries.length ||
      entries.length === 0
    ) {
      throw new Error(`getExtendedProperties: ${local} vt:vector size does not match its entries`);
    }
    return entries;
  };
  const headingPairEntries = readVector('HeadingPairs', 'variant');
  const headingPairs = (() => {
    if (headingPairEntries === null) return null;
    if (headingPairEntries.length % 2 !== 0) {
      throw new Error('getExtendedProperties: HeadingPairs requires heading/count pairs');
    }
    const pairs: ExtendedPropertyHeadingPair[] = [];
    for (let index = 0; index < headingPairEntries.length; index += 2) {
      const headingVariant = headingPairEntries[index]!;
      const countVariant = headingPairEntries[index + 1]!;
      const heading = childElements(headingVariant);
      const count = childElements(countVariant);
      if (
        headingVariant.name.namespaceURI !== NS_VARIANT_TYPES ||
        headingVariant.name.localName !== 'variant' ||
        countVariant.name.namespaceURI !== NS_VARIANT_TYPES ||
        countVariant.name.localName !== 'variant' ||
        heading.length !== 1 ||
        heading[0]!.name.namespaceURI !== NS_VARIANT_TYPES ||
        heading[0]!.name.localName !== 'lpstr' ||
        count.length !== 1 ||
        count[0]!.name.namespaceURI !== NS_VARIANT_TYPES ||
        count[0]!.name.localName !== 'i4'
      ) {
        throw new Error('getExtendedProperties: HeadingPairs contains an invalid variant pair');
      }
      const parsedCount = parseXsdInt(
        textContent(count[0]!),
        'getExtendedProperties: HeadingPairs count',
      );
      if (parsedCount < 0) {
        throw new Error('getExtendedProperties: HeadingPairs count must be non-negative');
      }
      pairs.push({ heading: textContent(heading[0]!), count: parsedCount });
    }
    return pairs;
  })();
  const titleEntries = readVector('TitlesOfParts', 'lpstr');
  const titlesOfParts =
    titleEntries === null
      ? null
      : titleEntries.map((entry) => {
          if (entry.name.namespaceURI !== NS_VARIANT_TYPES || entry.name.localName !== 'lpstr') {
            throw new Error('getExtendedProperties: TitlesOfParts contains a non-lpstr entry');
          }
          return textContent(entry);
        });
  return {
    application: read('Application'),
    appVersion: read('AppVersion'),
    company: read('Company'),
    manager: read('Manager'),
    presentationFormat: read('PresentationFormat'),
    hyperlinkBase: read('HyperlinkBase'),
    slides: readInteger('Slides'),
    notes: readInteger('Notes'),
    hiddenSlides: readInteger('HiddenSlides'),
    headingPairs,
    titlesOfParts,
  };
};

type ExtendedStringPropertyKey =
  | 'application'
  | 'appVersion'
  | 'company'
  | 'manager'
  | 'presentationFormat'
  | 'hyperlinkBase';

const EXT_PROP_STRING_FIELDS: ReadonlyArray<{
  key: ExtendedStringPropertyKey;
  local: string;
}> = [
  { key: 'application', local: 'Application' },
  { key: 'appVersion', local: 'AppVersion' },
  { key: 'company', local: 'Company' },
  { key: 'manager', local: 'Manager' },
  { key: 'presentationFormat', local: 'PresentationFormat' },
  { key: 'hyperlinkBase', local: 'HyperlinkBase' },
];

type ExtendedIntegerPropertyKey = 'slides' | 'notes' | 'hiddenSlides';

const EXT_PROP_INTEGER_FIELDS: ReadonlyArray<{
  key: ExtendedIntegerPropertyKey;
  local: string;
}> = [
  { key: 'slides', local: 'Slides' },
  { key: 'notes', local: 'Notes' },
  { key: 'hiddenSlides', local: 'HiddenSlides' },
];

/**
 * Writes selected fields on `/docProps/app.xml`. Throws when the
 * package has no extended-properties part. Bootstrapping remains a separate
 * package-level operation because it must also create the content-type and
 * root relationship; presentations created by this package already include
 * that complete relationship chain.
 *
 * Pass `null` to remove an existing field element. Unspecified keys
 * are left untouched.
 */
export const setExtendedProperties = (
  pres: PresentationData,
  values: Partial<ExtendedProperties>,
): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const part = pkg.getPart(EXT_PROPS_PART_NAME);
  if (!part) {
    throw new Error('setExtendedProperties: /docProps/app.xml not present; cannot bootstrap');
  }
  const doc = parseXml(decode(part.data));
  const setChildren = (local: string, children: XmlNode[] | null): void => {
    const name = qname('', local, NS_EXT_PROPS);
    const existing = firstChildElement(doc.root, name);
    if (children === null) {
      if (existing) doc.root.children = doc.root.children.filter((child) => child !== existing);
    } else if (existing) {
      existing.children = children;
    } else {
      doc.root.children.push(elem(name, { children }));
    }
  };
  for (const field of EXT_PROP_STRING_FIELDS) {
    if (!(field.key in values)) continue;
    const value = values[field.key] ?? null;
    setChildren(field.local, value === null ? null : [textNode(value)]);
  }
  for (const field of EXT_PROP_INTEGER_FIELDS) {
    if (!(field.key in values)) continue;
    const value = values[field.key] ?? null;
    if (value !== null && (!Number.isInteger(value) || value < 0 || value > 2_147_483_647)) {
      throw new RangeError(`setExtendedProperties: ${field.key} must be a non-negative xsd:int`);
    }
    setChildren(field.local, value === null ? null : [textNode(String(value))]);
  }
  if ('headingPairs' in values) {
    const pairs = values.headingPairs ?? null;
    if (pairs !== null && pairs.length === 0) {
      throw new RangeError('setExtendedProperties: headingPairs must be non-empty or null');
    }
    doc.root.prefixDecls.set('vt', NS_VARIANT_TYPES);
    const entries =
      pairs?.flatMap(({ heading, count }) => {
        if (!Number.isInteger(count) || count < 0 || count > 2_147_483_647) {
          throw new RangeError(
            'setExtendedProperties: heading-pair count must be a non-negative xsd:int',
          );
        }
        return [
          elem(qname('vt', 'variant', NS_VARIANT_TYPES), {
            children: [
              elem(qname('vt', 'lpstr', NS_VARIANT_TYPES), { children: [textNode(heading)] }),
            ],
          }),
          elem(qname('vt', 'variant', NS_VARIANT_TYPES), {
            children: [
              elem(qname('vt', 'i4', NS_VARIANT_TYPES), { children: [textNode(String(count))] }),
            ],
          }),
        ];
      }) ?? null;
    setChildren(
      'HeadingPairs',
      entries === null
        ? null
        : [
            elem(qname('vt', 'vector', NS_VARIANT_TYPES), {
              attrs: [
                attr(qname('', 'size', ''), String(entries.length)),
                attr(qname('', 'baseType', ''), 'variant'),
              ],
              children: entries,
            }),
          ],
    );
  }
  if ('titlesOfParts' in values) {
    const titles = values.titlesOfParts ?? null;
    if (titles !== null && titles.length === 0) {
      throw new RangeError('setExtendedProperties: titlesOfParts must be non-empty or null');
    }
    doc.root.prefixDecls.set('vt', NS_VARIANT_TYPES);
    setChildren(
      'TitlesOfParts',
      titles === null
        ? null
        : [
            elem(qname('vt', 'vector', NS_VARIANT_TYPES), {
              attrs: [
                attr(qname('', 'size', ''), String(titles.length)),
                attr(qname('', 'baseType', ''), 'lpstr'),
              ],
              children: titles.map((title) =>
                elem(qname('vt', 'lpstr', NS_VARIANT_TYPES), {
                  children: [textNode(title)],
                }),
              ),
            }),
          ],
    );
  }
  part.data = encode(serializeXml(doc));
};
