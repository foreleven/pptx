import {
  INTERNAL_PACKAGE,
  SLIDE_DOCUMENT,
  SLIDE_PART_NAME,
  type PresentationData,
  type SlideData,
} from './_internal-symbols.ts';
import type { OpcPackage } from '../internal/parts/index.ts';
import { partName } from '../internal/opc/index.ts';
import { parseXml } from '../internal/xml/index.ts';
import { decode, refreshSlideData } from './fn/_helpers.ts';

export type PresentationMutationRecord =
  | {
      readonly type: 'part-put';
      readonly partName: string;
      readonly contentType: string;
      readonly beforeHash: string | null;
      readonly afterHash: string;
      readonly data: Uint8Array;
    }
  | {
      readonly type: 'part-delete';
      readonly partName: string;
      readonly beforeHash: string;
    };

export type PresentationObserverCallback = (
  records: readonly PresentationMutationRecord[],
  observer: PresentationObserver,
) => void;

interface PartSnapshot {
  readonly contentType: string;
  readonly hash: string;
  readonly data: Uint8Array;
}

const observersByPackage = new WeakMap<OpcPackage, Set<PresentationObserver>>();

const hashBytes = (bytes: Uint8Array): string => {
  let hash = 0xcbf29ce484222325n;
  for (const byte of bytes) {
    hash ^= BigInt(byte);
    hash = BigInt.asUintN(64, hash * 0x100000001b3n);
  }
  return `fnv1a64:${hash.toString(16).padStart(16, '0')}`;
};

const snapshotPackage = (pkg: OpcPackage): Map<string, PartSnapshot> =>
  new Map(
    pkg.parts.map((part) => {
      const data = new Uint8Array(part.data);
      return [
        part.name,
        {
          contentType: part.contentType,
          hash: hashBytes(data),
          data,
        },
      ];
    }),
  );

const diffSnapshots = (
  before: ReadonlyMap<string, PartSnapshot>,
  after: ReadonlyMap<string, PartSnapshot>,
): PresentationMutationRecord[] => {
  const records: PresentationMutationRecord[] = [];
  for (const [name, next] of after) {
    const previous = before.get(name);
    if (previous?.hash === next.hash && previous.contentType === next.contentType) continue;
    records.push({
      type: 'part-put',
      partName: name,
      contentType: next.contentType,
      beforeHash: previous?.hash ?? null,
      afterHash: next.hash,
      data: new Uint8Array(next.data),
    });
  }
  for (const [name, previous] of before) {
    if (!after.has(name))
      records.push({ type: 'part-delete', partName: name, beforeHash: previous.hash });
  }
  return records;
};

export class PresentationObserver {
  readonly #callback: PresentationObserverCallback;
  #presentation: PresentationData | null = null;
  #baseline: Map<string, PartSnapshot> | null = null;
  #deliveryScheduled = false;
  #connectedGeneration = 0;
  #unsubscribePackage: (() => void) | null = null;

  constructor(callback: PresentationObserverCallback) {
    this.#callback = callback;
  }

  connect(presentation: PresentationData): void {
    if (this.#presentation === presentation) return;
    if (this.#presentation !== null) {
      throw new Error('PresentationObserver is already connected to another presentation.');
    }
    this.#presentation = presentation;
    this.#baseline = snapshotPackage(presentation[INTERNAL_PACKAGE]);
    const observers = observersByPackage.get(presentation[INTERNAL_PACKAGE]) ?? new Set();
    observers.add(this);
    observersByPackage.set(presentation[INTERNAL_PACKAGE], observers);
    this.#unsubscribePackage = presentation[INTERNAL_PACKAGE].onChange(() =>
      this.scheduleDelivery(),
    );
    this.#connectedGeneration += 1;
  }

  disconnect(): void {
    if (!this.#presentation) return;
    observersByPackage.get(this.#presentation[INTERNAL_PACKAGE])?.delete(this);
    this.#unsubscribePackage?.();
    this.#unsubscribePackage = null;
    this.#presentation = null;
    this.#baseline = null;
    this.#deliveryScheduled = false;
    this.#connectedGeneration += 1;
  }

  /** @internal */
  scheduleDelivery(): void {
    if (!this.#presentation || this.#deliveryScheduled) return;
    this.#deliveryScheduled = true;
    const generation = this.#connectedGeneration;
    queueMicrotask(() => {
      this.#deliveryScheduled = false;
      if (!this.#presentation || generation !== this.#connectedGeneration || !this.#baseline)
        return;
      const next = snapshotPackage(this.#presentation[INTERNAL_PACKAGE]);
      const records = diffSnapshots(this.#baseline, next);
      this.#baseline = next;
      if (records.length > 0) this.#callback(records, this);
    });
  }

  /** @internal */
  replaceBaseline(): void {
    if (this.#presentation) this.#baseline = snapshotPackage(this.#presentation[INTERNAL_PACKAGE]);
  }
}

/** @internal — mutation commits call this after package bytes change. */
export const notifyPresentationObservers = (pkg: OpcPackage): void => {
  for (const observer of observersByPackage.get(pkg) ?? []) observer.scheduleDelivery();
};

export const applyPresentationMutationRecords = (
  presentation: PresentationData,
  records: readonly PresentationMutationRecord[],
  options: { notifyObservers?: boolean } = {},
): void => {
  const pkg = presentation[INTERNAL_PACKAGE];
  const cachedSlides = presentation._slidesCache as SlideData[] | null;
  for (const record of records) {
    if (record.type === 'part-put' && hashBytes(record.data) !== record.afterHash) {
      throw new Error(
        `Presentation part ${record.partName} data does not match ${record.afterHash}.`,
      );
    }
    const existing = pkg.getPart(partName(record.partName));
    const actualHash = existing ? hashBytes(existing.data) : null;
    if (actualHash !== record.beforeHash) {
      throw new Error(
        `Presentation part ${record.partName} expected ${record.beforeHash ?? 'missing'} but found ${actualHash ?? 'missing'}.`,
      );
    }
  }

  for (const record of records) {
    const name = partName(record.partName);
    if (record.type === 'part-delete') {
      pkg.removePart(name);
      continue;
    }
    const existing = pkg.getPart(name);
    if (existing) {
      existing.contentType = record.contentType;
      existing.data = new Uint8Array(record.data);
    } else {
      pkg.addPart(name, record.contentType, new Uint8Array(record.data));
    }
  }

  let invalidateSlideCache = false;
  if (cachedSlides) {
    for (const record of records) {
      if (
        record.partName === '/ppt/presentation.xml' ||
        record.partName === '/ppt/_rels/presentation.xml.rels'
      ) {
        invalidateSlideCache = true;
      }
      if (!/^\/ppt\/slides\/slide[^/]+\.xml$/i.test(record.partName)) continue;
      if (record.type === 'part-delete' || record.beforeHash === null) {
        invalidateSlideCache = true;
        continue;
      }
      const slide = cachedSlides.find(
        (candidate) =>
          String(candidate[SLIDE_PART_NAME]).toLowerCase() === record.partName.toLowerCase(),
      );
      if (!slide) continue;
      slide[SLIDE_DOCUMENT] = parseXml(decode(record.data));
      refreshSlideData(slide);
    }
  } else {
    invalidateSlideCache = true;
  }
  if (invalidateSlideCache) presentation._slidesCache = null;
  if (options.notifyObservers === false) {
    for (const observer of observersByPackage.get(pkg) ?? []) observer.replaceBaseline();
  } else {
    notifyPresentationObservers(pkg);
  }
};
