// `OpcPackage` is the in-memory model of a `.pptx` archive at the OPC layer:
// an ordered list of parts plus a `[Content_Types].xml` map.
//
// Round-trip rules:
//
//   - Parts are kept in their original load order. New parts append at the
//     end. This matches what PowerPoint emits and keeps diffs small.
//   - `.rels` files are first-class parts. They are NOT auto-regenerated from
//     a parsed model on every save — that would force callers who only edit
//     content parts to pay the cost of round-tripping every relationship.
//     Higher-level edits that need a typed view go through `getRels` / `setRels`,
//     which parse on read and re-serialize on write back into the part.
//   - The package never invents content types. Callers must register an
//     override (or rely on the extension default) before adding a part whose
//     extension isn't already covered. Otherwise `validate()` will surface
//     the gap.

import {
  type ContentTypes,
  type PartName,
  type Relationships,
  type ZipEntry,
  emptyContentTypes,
  fromZipPath,
  lookupContentType,
  parseContentTypes,
  parseRels,
  partName,
  partNamesEqual,
  readZip,
  relsPartNameFor,
  serializeContentTypes,
  serializeRels,
  toZipPath,
  writeZip,
} from '../opc/index.ts';

const TEXT_DECODER = new TextDecoder();
const TEXT_ENCODER = new TextEncoder();

const decode = (b: Uint8Array): string => TEXT_DECODER.decode(b);
const encode = (s: string): Uint8Array => TEXT_ENCODER.encode(s);

/** A single OPC part: name, content type, and bytes. */
export interface Part {
  readonly name: PartName;
  contentType: string;
  data: Uint8Array;
}

const CONTENT_TYPES_PART = '[Content_Types].xml';

/**
 * In-memory model of an OPC package. Construct via `OpcPackage.load(bytes)`
 * to read an existing `.pptx`, or via `OpcPackage.empty()` to start from
 * scratch.
 */
export class OpcPackage {
  /**
   * Ordered list of parts. The order is the order parts will be written to
   * the ZIP central directory on `save()`. This is a readonly view: use
   * `addPart` / `removePart`, and replace a Part's `data` value rather than
   * mutating its Uint8Array in place, so package observers are notified.
   */
  get parts(): readonly Part[] {
    return this.orderedParts;
  }
  /**
   * Mutable Content_Types map for package internals. A content-type-only
   * mutation is not an observable presentation operation; public Core
   * mutations pair registration changes with an observed part add/update.
   */
  contentTypes: ContentTypes;
  private readonly orderedParts: Part[];
  private readonly changeListeners = new Set<() => void>();

  private constructor(parts: Part[], contentTypes: ContentTypes) {
    this.orderedParts = parts;
    this.contentTypes = contentTypes;
    for (const part of parts) this.observePart(part);
  }

  /** @internal — presentation observers use this without exposing package state publicly. */
  onChange(listener: () => void): () => void {
    this.changeListeners.add(listener);
    return () => this.changeListeners.delete(listener);
  }

  private notifyChange(): void {
    for (const listener of this.changeListeners) listener();
  }

  private observePart(part: Part): void {
    let contentType = part.contentType;
    let data = part.data;
    Object.defineProperties(part, {
      contentType: {
        configurable: true,
        enumerable: true,
        get: () => contentType,
        set: (next: string) => {
          if (next === contentType) return;
          contentType = next;
          this.contentTypes.overrides = this.contentTypes.overrides.filter(
            (override) => !partNamesEqual(override.partName, part.name),
          );
          if (lookupContentType(this.contentTypes, part.name) !== next) {
            this.contentTypes.overrides.push({ partName: part.name, contentType: next });
          }
          this.notifyChange();
        },
      },
      data: {
        configurable: true,
        enumerable: true,
        get: () => data,
        set: (next: Uint8Array) => {
          if (next === data) return;
          data = next;
          this.notifyChange();
        },
      },
    });
  }

  /**
   * Builds an empty package containing only the two default content-type
   * entries (xml, rels). No `.rels` or content parts yet — that's the caller's
   * job. PresentationML scaffolding will populate the typical PPTX skeleton.
   */
  static empty(): OpcPackage {
    return new OpcPackage([], emptyContentTypes());
  }

  /**
   * Loads a `.pptx` file from raw bytes. Parses `[Content_Types].xml`
   * eagerly; leaves part bodies as raw `Uint8Array`s for higher layers to
   * interpret.
   *
   * Throws if the archive is not a valid ZIP, or if `[Content_Types].xml` is
   * missing or malformed.
   */
  static load(bytes: Uint8Array): OpcPackage {
    const zip = readZip(bytes);
    let contentTypes: ContentTypes | null = null;
    const parts: Part[] = [];

    for (const entry of zip.entries) {
      if (entry.name === CONTENT_TYPES_PART) {
        contentTypes = parseContentTypes(decode(entry.data));
        continue;
      }
      // Some producers (PowerPoint on Windows, third-party tools) include
      // explicit directory entries in the ZIP central directory — zero-byte
      // names ending with "/". They are not OPC parts; skip them so part-name
      // validation does not reject them.
      if (entry.name.endsWith('/')) continue;
      const name = fromZipPath(entry.name);
      parts.push({
        name,
        contentType: '', // filled in after we've parsed Content_Types
        data: entry.data,
      });
    }

    if (contentTypes === null) {
      throw new Error('package is missing [Content_Types].xml');
    }

    for (const part of parts) {
      const ct = lookupContentType(contentTypes, part.name);
      if (ct === null) {
        throw new Error(
          `part "${part.name}" has no content type registered in [Content_Types].xml`,
        );
      }
      part.contentType = ct;
    }

    return new OpcPackage(parts, contentTypes);
  }

  /** Serialize the package to a `.pptx` byte stream. */
  save(): Uint8Array {
    const entries: ZipEntry[] = [];
    entries.push({
      name: CONTENT_TYPES_PART,
      data: encode(serializeContentTypes(this.contentTypes)),
    });
    for (const part of this.orderedParts) {
      entries.push({ name: toZipPath(part.name), data: part.data });
    }
    return writeZip(entries);
  }

  /**
   * Looks up a part by name (case-insensitive per OPC §9.1.1.7).
   */
  getPart(name: PartName): Part | null {
    for (const p of this.orderedParts) {
      if (partNamesEqual(p.name, name)) return p;
    }
    return null;
  }

  /**
   * Adds a new part. Throws if a part with the same name (case-insensitive)
   * already exists. The caller is responsible for ensuring the content type
   * is resolvable — either by extension default or by a registered override.
   */
  addPart(name: PartName, contentType: string, data: Uint8Array): Part {
    if (this.getPart(name) !== null) {
      throw new Error(`part "${name}" already exists`);
    }
    const part: Part = { name, contentType, data };
    this.observePart(part);
    this.orderedParts.push(part);
    // If neither an existing default nor an existing override covers this
    // content type, register an Override so the on-disk Content_Types stays
    // consistent. Callers can replace the override later if they prefer a
    // default.
    const existing = lookupContentType(this.contentTypes, name);
    if (existing !== contentType) {
      this.contentTypes.overrides.push({ partName: name, contentType });
    }
    this.notifyChange();
    return part;
  }

  /**
   * Removes a part by name. Returns true if a part was removed; false if no
   * part with that name was present. Also drops any matching `Override`
   * entry from Content_Types; defaults are left untouched.
   */
  removePart(name: PartName): boolean {
    let removed = false;
    for (let i = this.orderedParts.length - 1; i >= 0; i--) {
      const p = this.orderedParts[i];
      if (p && partNamesEqual(p.name, name)) {
        this.orderedParts.splice(i, 1);
        removed = true;
      }
    }
    if (removed) {
      this.contentTypes.overrides = this.contentTypes.overrides.filter(
        (o) => !partNamesEqual(o.partName, name),
      );
      this.notifyChange();
    }
    return removed;
  }

  /**
   * Reads (or parses, lazily) the relationships for the given part. Returns
   * `null` if the part has no sibling `.rels` part — this is normal for
   * content parts with no outgoing relationships.
   *
   * The returned `Relationships` is a fresh snapshot. Mutate it as needed,
   * then call `setRels(forPart, rels)` to persist the change back into the
   * package.
   */
  getRels(forPart: PartName): Relationships | null {
    const relsName = relsPartNameFor(forPart);
    const part = this.getPart(relsName);
    if (part === null) return null;
    return parseRels(decode(part.data));
  }

  /**
   * Writes relationships for the given part. Creates the sibling `.rels`
   * part if it didn't exist; replaces its bytes if it did. The relationships
   * part is given the standard relationships content type via the `rels`
   * extension default.
   */
  setRels(forPart: PartName, rels: Relationships): void {
    const relsName = relsPartNameFor(forPart);
    const bytes = encode(serializeRels(rels));
    const existing = this.getPart(relsName);
    if (existing !== null) {
      existing.data = bytes;
      return;
    }
    this.addPart(relsName, 'application/vnd.openxmlformats-package.relationships+xml', bytes);
  }

  /**
   * Convenience: read the package-root relationships (`/_rels/.rels`).
   * These point at the top-level parts — typically the presentation.xml,
   * core properties, and extended properties.
   */
  rootRels(): Relationships | null {
    const part = this.getPart(partName('/_rels/.rels'));
    if (part === null) return null;
    return parseRels(decode(part.data));
  }

  /** Convenience: write package-root relationships. */
  setRootRels(rels: Relationships): void {
    const name = partName('/_rels/.rels');
    const bytes = encode(serializeRels(rels));
    const existing = this.getPart(name);
    if (existing !== null) {
      existing.data = bytes;
      return;
    }
    this.addPart(name, 'application/vnd.openxmlformats-package.relationships+xml', bytes);
  }
}
