import { type ZipOptions, unzipSync, zipSync } from 'fflate';
import {
  emptyRels,
  nextRelId,
  partName,
  resolveTarget,
  type PartName,
  type Relationship,
} from '../../internal/opc/index.ts';
import {
  NS,
  type XmlElement,
  parseFragment,
  serializeFragment,
  walkElements,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SLIDE_PART_NAME,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { appendAndReturnNewShape, requireSpTree } from './_helpers.ts';

const FRAGMENT_CONTENT_TYPE = 'application/vnd.deer-presentation.ooxml-object';
const encoder = new TextEncoder();
const decoder = new TextDecoder();
const MAX_FRAGMENT_BYTES = 32 * 1024 * 1024;
const MAX_EXPANDED_BYTES = 128 * 1024 * 1024;
const MAX_ENTRY_BYTES = 64 * 1024 * 1024;
const MAX_ENTRIES = 1024;
const MAX_COMPRESSION_RATIO = 100;
const FRAGMENT_ZIP_MTIME = new Date(1980, 0, 1, 0, 0, 0, 0);

/** Classify the preserved object for registries and diagnostics without exposing its XML. */
export const getOpaqueObjectKind = (shape: SlideShapeData): string =>
  opaqueObjectKind(shape[SHAPE_ELEMENT]);

interface FragmentRelationship {
  id: string;
  type: string;
  targetMode: 'Internal' | 'External';
  target: string;
  targetPartName?: string;
}

interface FragmentPart {
  originalName: string;
  contentType: string;
  file: string;
  relationships: FragmentRelationship[];
}

interface OpaqueFragmentManifest {
  schemaVersion: 1;
  contentType: typeof FRAGMENT_CONTENT_TYPE;
  kind: string;
  sourceSlidePartName: string;
  shapeFile: 'shape.xml';
  relationships: FragmentRelationship[];
  parts: FragmentPart[];
}

/** Extract one shape plus the complete internal OPC relationship closure it references. */
export const extractOpaqueObjectFragment = (shape: SlideShapeData): Uint8Array => {
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const directIds = relationshipIds(shape[SHAPE_ELEMENT]);
  const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]);
  const relationships = directIds.map((id) => {
    const relationship = slideRels?.items.find((item) => item.id === id);
    if (!relationship)
      throw new Error(`extractOpaqueObjectFragment: missing slide relationship ${id}`);
    return fragmentRelationship(slide[SLIDE_PART_NAME], relationship);
  });

  const queued = relationships.flatMap((relationship) =>
    relationship.targetMode === 'Internal' &&
    relationship.targetPartName &&
    !isSlidePart(relationship.targetPartName)
      ? [partName(relationship.targetPartName)]
      : [],
  );
  const visited = new Set<string>();
  const parts: FragmentPart[] = [];
  const files: Record<string, Uint8Array> = {
    'shape.xml': encoder.encode(serializeFragment(selfContainedShape(shape[SHAPE_ELEMENT]))),
  };
  while (queued.length > 0) {
    const current = queued.shift()!;
    if (visited.has(current.toLowerCase())) continue;
    visited.add(current.toLowerCase());
    const part = pkg.getPart(current);
    if (!part) throw new Error(`extractOpaqueObjectFragment: missing related part ${current}`);
    const partRelationships = (pkg.getRels(current)?.items ?? []).map((relationship) =>
      fragmentRelationship(current, relationship),
    );
    for (const relationship of partRelationships) {
      if (
        relationship.targetMode === 'Internal' &&
        relationship.targetPartName &&
        !isSlidePart(relationship.targetPartName)
      ) {
        queued.push(partName(relationship.targetPartName));
      }
    }
    const file = `parts/${String(parts.length + 1).padStart(4, '0')}.bin`;
    files[file] = new Uint8Array(part.data);
    parts.push({
      originalName: part.name,
      contentType: part.contentType,
      file,
      relationships: partRelationships,
    });
  }
  const manifest: OpaqueFragmentManifest = {
    schemaVersion: 1,
    contentType: FRAGMENT_CONTENT_TYPE,
    kind: opaqueObjectKind(shape[SHAPE_ELEMENT]),
    sourceSlidePartName: slide[SLIDE_PART_NAME],
    shapeFile: 'shape.xml',
    relationships,
    parts,
  };
  files['manifest.json'] = encoder.encode(`${JSON.stringify(manifest, null, 2)}\n`);
  const deterministicFiles: Record<string, [Uint8Array, ZipOptions]> = {};
  for (const [name, bytes] of Object.entries(files)) {
    deterministicFiles[name] = [bytes, { level: 6, mtime: FRAGMENT_ZIP_MTIME }];
  }
  return zipSync(deterministicFiles);
};

/** Insert one validated opaque fragment into a slide, rebinding its package relationships. */
export const addSlideOpaqueObject = (
  slide: SlideData,
  fragmentBytes: Uint8Array,
): SlideShapeData => {
  validateZipEnvelope(fragmentBytes);
  const files = unzipSync(fragmentBytes);
  const manifestBytes = files['manifest.json'];
  const shapeBytes = files['shape.xml'];
  if (!manifestBytes || !shapeBytes)
    throw new Error('addSlideOpaqueObject: fragment is missing manifest.json or shape.xml');
  const manifest = parseManifest(decoder.decode(manifestBytes));
  const shape = parseFragment(decoder.decode(shapeBytes));
  if (
    shape.name.namespaceURI !== NS.pml ||
    !['sp', 'pic', 'cxnSp', 'graphicFrame', 'grpSp'].includes(shape.name.localName)
  ) {
    throw new Error('addSlideOpaqueObject: fragment root is not a PresentationML shape');
  }

  const pkg = slide[INTERNAL_PACKAGE];
  const partMap = new Map<string, PartName>();
  for (const entry of manifest.parts) {
    const original = partName(entry.originalName);
    const allocated = allocatePartName(
      pkg.parts.map((part) => part.name),
      original,
    );
    const data = files[entry.file];
    if (!data) throw new Error(`addSlideOpaqueObject: missing payload ${entry.file}`);
    pkg.addPart(allocated, entry.contentType, new Uint8Array(data));
    partMap.set(original.toLowerCase(), allocated);
  }
  const availableParts = new Set(pkg.parts.map((part) => part.name.toLowerCase()));
  for (const entry of manifest.parts) {
    const source = partMap.get(entry.originalName.toLowerCase())!;
    if (entry.relationships.length === 0) continue;
    const rels = emptyRels();
    rels.items.push(
      ...entry.relationships.map((relationship) =>
        reboundRelationship(source, relationship, partMap, availableParts),
      ),
    );
    pkg.setRels(source, rels);
  }

  const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const usedIds = slideRels.items.map((relationship) => relationship.id);
  const idMap = new Map<string, string>();
  for (const relationship of manifest.relationships) {
    const id = nextRelId(usedIds);
    usedIds.push(id);
    idMap.set(relationship.id, id);
    slideRels.items.push({
      ...reboundRelationship(slide[SLIDE_PART_NAME], relationship, partMap, availableParts),
      id,
    });
  }
  pkg.setRels(slide[SLIDE_PART_NAME], slideRels);
  rewriteRelationshipIds(shape, idMap);
  rewriteShapeIds(shape, slide);
  return appendAndReturnNewShape(slide, shape);
};

function fragmentRelationship(source: PartName, relationship: Relationship): FragmentRelationship {
  if (relationship.targetMode === 'External') return { ...relationship, targetMode: 'External' };
  const targetPartName = resolveTarget(source, relationship.target);
  return { ...relationship, targetMode: 'Internal', targetPartName };
}

function reboundRelationship(
  source: PartName,
  relationship: FragmentRelationship,
  partMap: ReadonlyMap<string, PartName>,
  availableParts: ReadonlySet<string>,
): Relationship {
  if (relationship.targetMode === 'External') {
    return {
      id: relationship.id,
      type: relationship.type,
      target: relationship.target,
      targetMode: 'External',
    };
  }
  if (!relationship.targetPartName)
    throw new Error(`addSlideOpaqueObject: relationship ${relationship.id} has no target part`);
  const original = partName(relationship.targetPartName);
  const target =
    partMap.get(original.toLowerCase()) ??
    (isSlidePart(original) && availableParts.has(original.toLowerCase()) ? original : null);
  if (!target)
    throw new Error(`addSlideOpaqueObject: missing relationship closure for ${original}`);
  return {
    id: relationship.id,
    type: relationship.type,
    target: relativeTarget(source, target),
    targetMode: 'Internal',
  };
}

function parseManifest(text: string): OpaqueFragmentManifest {
  const value = JSON.parse(text) as Partial<OpaqueFragmentManifest>;
  if (
    value.schemaVersion !== 1 ||
    value.contentType !== FRAGMENT_CONTENT_TYPE ||
    value.shapeFile !== 'shape.xml' ||
    typeof value.kind !== 'string' ||
    typeof value.sourceSlidePartName !== 'string' ||
    !Array.isArray(value.relationships) ||
    !Array.isArray(value.parts)
  ) {
    throw new Error('addSlideOpaqueObject: invalid fragment manifest');
  }
  return value as OpaqueFragmentManifest;
}

function relationshipIds(root: XmlElement): string[] {
  const ids = new Set<string>();
  walkElements(root, (element) => {
    for (const attribute of element.attrs) {
      if (
        attribute.name.namespaceURI === NS.officeDocRels &&
        ['id', 'embed', 'link'].includes(attribute.name.localName) &&
        attribute.value
      ) {
        ids.add(attribute.value);
      }
    }
  });
  return [...ids];
}

function selfContainedShape(source: XmlElement): XmlElement {
  const clone = (element: XmlElement): XmlElement => ({
    kind: 'element',
    name: element.name,
    attrs: element.attrs.map((attribute) => ({ name: attribute.name, value: attribute.value })),
    prefixDecls: new Map(element.prefixDecls),
    children: element.children.map((child) =>
      child.kind === 'element' ? clone(child) : { ...child },
    ),
  });
  const root = clone(source);
  walkElements(root, (element) => {
    if (element.name.prefix) root.prefixDecls.set(element.name.prefix, element.name.namespaceURI);
    for (const attribute of element.attrs) {
      if (attribute.name.prefix)
        root.prefixDecls.set(attribute.name.prefix, attribute.name.namespaceURI);
    }
  });
  return root;
}

function rewriteRelationshipIds(root: XmlElement, ids: ReadonlyMap<string, string>): void {
  walkElements(root, (element) => {
    element.attrs = element.attrs.map((attribute) =>
      attribute.name.namespaceURI === NS.officeDocRels && ids.has(attribute.value)
        ? { name: attribute.name, value: ids.get(attribute.value)! }
        : attribute,
    );
  });
}

function rewriteShapeIds(root: XmlElement, slide: SlideData): void {
  const usedIds = new Set<number>();
  walkElements(requireSpTree(slide), (element) => {
    if (element.name.namespaceURI !== NS.pml || element.name.localName !== 'cNvPr') return;
    const value = Number(
      element.attrs.find(
        (attribute) => attribute.name.namespaceURI === '' && attribute.name.localName === 'id',
      )?.value,
    );
    if (Number.isSafeInteger(value) && value >= 0) usedIds.add(value);
  });

  let nextId = 2;
  for (const id of usedIds) nextId = Math.max(nextId, id + 1);
  const idMap = new Map<string, string>();
  walkElements(root, (element) => {
    if (element.name.namespaceURI !== NS.pml || element.name.localName !== 'cNvPr') return;
    element.attrs = element.attrs.map((attribute) => {
      if (attribute.name.namespaceURI !== '' || attribute.name.localName !== 'id') {
        return attribute;
      }
      const allocated = String(nextId++);
      idMap.set(attribute.value, allocated);
      return { name: attribute.name, value: allocated };
    });
  });

  walkElements(root, (element) => {
    if (
      element.name.namespaceURI !== NS.dml ||
      !['stCxn', 'endCxn'].includes(element.name.localName)
    )
      return;
    element.attrs = element.attrs.map((attribute) =>
      attribute.name.namespaceURI === '' &&
      attribute.name.localName === 'id' &&
      idMap.has(attribute.value)
        ? { name: attribute.name, value: idMap.get(attribute.value)! }
        : attribute,
    );
  });
}

function opaqueObjectKind(root: XmlElement): string {
  let kind = root.name.localName === 'graphicFrame' ? 'graphic-frame' : root.name.localName;
  walkElements(root, (element) => {
    if (element.name.namespaceURI === NS.pml && element.name.localName === 'oleObj') kind = 'ole';
    if (element.name.namespaceURI === NS.dml && element.name.localName === 'graphicData') {
      const uri =
        element.attrs.find((attribute) => attribute.name.localName === 'uri')?.value ?? '';
      if (uri.includes('/diagram')) kind = 'smartart';
      if (uri.includes('/chart')) kind = 'chart';
    }
  });
  return kind;
}

function allocatePartName(existing: readonly string[], original: PartName): PartName {
  const occupied = new Set(existing.map((name) => name.toLowerCase()));
  if (!occupied.has(original.toLowerCase())) return original;
  const dot = original.lastIndexOf('.');
  const stem = dot > original.lastIndexOf('/') ? original.slice(0, dot) : original;
  const extension = dot > original.lastIndexOf('/') ? original.slice(dot) : '';
  for (let index = 1; ; index += 1) {
    const candidate = partName(`${stem}-opaque${index}${extension}`);
    if (!occupied.has(candidate.toLowerCase())) return candidate;
  }
}

function relativeTarget(source: PartName, target: PartName): string {
  const from = source.slice(1).split('/').slice(0, -1);
  const to = target.slice(1).split('/');
  while (from.length > 0 && to.length > 0 && from[0]!.toLowerCase() === to[0]!.toLowerCase()) {
    from.shift();
    to.shift();
  }
  return `${from.map(() => '..').join('/')}${from.length > 0 && to.length > 0 ? '/' : ''}${to.join('/')}`;
}

function isSlidePart(name: string): boolean {
  return /^\/ppt\/slides\/slide\d+\.xml$/iu.test(name);
}

function validateZipEnvelope(bytes: Uint8Array): void {
  if (bytes.byteLength === 0 || bytes.byteLength > MAX_FRAGMENT_BYTES) {
    throw new Error('addSlideOpaqueObject: fragment compressed size exceeds the safety limit');
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let entries = 0;
  let expanded = 0;
  for (let offset = 0; offset + 46 <= bytes.byteLength; offset += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) continue;
    const uncompressed = view.getUint32(offset + 24, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const end = offset + 46 + nameLength + extraLength + commentLength;
    if (end > bytes.byteLength)
      throw new Error('addSlideOpaqueObject: malformed ZIP central directory');
    const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
    if (!name || name.startsWith('/') || name.includes('\\') || name.split('/').includes('..')) {
      throw new Error(`addSlideOpaqueObject: unsafe ZIP entry ${JSON.stringify(name)}`);
    }
    entries += 1;
    expanded += uncompressed;
    if (entries > MAX_ENTRIES || uncompressed > MAX_ENTRY_BYTES || expanded > MAX_EXPANDED_BYTES) {
      throw new Error('addSlideOpaqueObject: fragment expanded size exceeds the safety limit');
    }
    offset = end - 1;
  }
  if (entries === 0 || expanded / bytes.byteLength > MAX_COMPRESSION_RATIO) {
    throw new Error('addSlideOpaqueObject: fragment compression ratio exceeds the safety limit');
  }
}
