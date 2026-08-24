import { normalizeGuid } from './bounds.ts';

const GUID_HASH_SEEDS = [0x811c9dc5, 0x9e3779b9, 0x85ebca6b, 0xc2b2ae35] as const;

/** Allocates stable, unique OOXML GUIDs within one caller-defined scope. */
export class GuidAllocator {
  readonly #used = new Set<string>();

  constructor(existing: Iterable<string> = []) {
    for (const value of existing)
      this.#used.add(normalizeGuid(value, 'GuidAllocator: existing GUID'));
  }

  reserve(value: string, field: string): string {
    const normalized = normalizeGuid(value, field);
    if (this.#used.has(normalized))
      throw new TypeError(`${field}: duplicate GUID ${JSON.stringify(value)}.`);
    this.#used.add(normalized);
    return normalized;
  }

  allocateStable(key: string): string {
    for (let collision = 0; ; collision += 1) {
      const source = collision === 0 ? key : `${key}\u0000${String(collision)}`;
      const candidate = guidFromStableHash(source);
      if (!this.#used.has(candidate)) {
        this.#used.add(candidate);
        return candidate;
      }
    }
  }
}

function guidFromStableHash(value: string): string {
  const hex = GUID_HASH_SEEDS.map((seed) => fnv1a32(value, seed).toString(16).padStart(8, '0'))
    .join('')
    .toUpperCase();
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-5${hex.slice(13, 16)}-8${hex.slice(17, 20)}-${hex.slice(20, 32)}}`;
}

function fnv1a32(value: string, seed: number): number {
  let hash = seed >>> 0;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return hash;
}
