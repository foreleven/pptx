import {
  _internalPackageOf,
  findSlideByPartName,
  getShapeImagePartName,
  getSlidePartName,
  getSlideShapes,
  getSlideSize,
  getSlides,
  type PresentationData,
  type PresentationMutationRecord,
  type SlideData,
} from '@office-kit/pptx';
import { renderSlideSvg } from './render-slide.ts';
import type { RenderSlideOptions } from './text-layout.ts';

export interface RenderedSlide {
  readonly slidePartName: string;
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly svg: string;
  readonly diagnostics: readonly RenderDiagnostic[];
}

export interface RenderDiagnostic {
  readonly kind: string;
  readonly message: string;
}

const slidePartFromRelationshipPart = (name: string): string | null => {
  const match = /^\/ppt\/slides\/_rels\/(slide[^/]+\.xml)\.rels$/i.exec(name);
  return match?.[1] ? `/ppt/slides/${match[1]}` : null;
};

const sourcePartFromRelationshipPart = (name: string): string | null => {
  const match = /^(.*)\/_rels\/([^/]+)\.rels$/i.exec(name);
  return match?.[1] && match[2] ? `${match[1]}/${match[2]}` : null;
};

const resolveRelationshipTarget = (source: string, target: string): string => {
  if (target.startsWith('/')) return target;
  const segments = `${source.slice(0, source.lastIndexOf('/') + 1)}${target}`.split('/');
  const normalized: string[] = [];
  for (const segment of segments) {
    if (!segment || segment === '.') continue;
    if (segment === '..') normalized.pop();
    else normalized.push(segment);
  }
  return `/${normalized.join('/')}`;
};

const slidesDependingOnPart = (
  presentation: PresentationData,
  target: string,
): readonly SlideData[] => {
  const pkg = _internalPackageOf(presentation);
  const reverseDependencies = new Map<string, Set<string>>();
  for (const part of pkg.parts) {
    if (String(part.name).endsWith('.rels')) continue;
    for (const relationship of pkg.getRels(part.name)?.items ?? []) {
      if (relationship.targetMode === 'External') continue;
      const resolved = resolveRelationshipTarget(String(part.name), relationship.target);
      const sources = reverseDependencies.get(resolved) ?? new Set<string>();
      sources.add(String(part.name));
      reverseDependencies.set(resolved, sources);
    }
  }
  const affectedParts = new Set<string>([target]);
  const pending = [target];
  while (pending.length > 0) {
    const current = pending.pop()!;
    for (const source of reverseDependencies.get(current) ?? []) {
      if (affectedParts.has(source)) continue;
      affectedParts.add(source);
      pending.push(source);
    }
  }
  return getSlides(presentation).filter((slide) => affectedParts.has(getSlidePartName(slide)));
};

export class PresentationRenderer {
  readonly #options: RenderSlideOptions;
  readonly #cache = new Map<string, RenderedSlide>();
  #presentation: PresentationData | null = null;

  constructor(options: RenderSlideOptions = {}) {
    this.#options = options;
  }

  connect(presentation: PresentationData): void {
    if (this.#presentation === presentation) return;
    if (this.#presentation !== null) {
      throw new Error('PresentationRenderer is already connected to another presentation.');
    }
    this.#presentation = presentation;
    this.#cache.clear();
  }

  disconnect(): void {
    this.#presentation = null;
    this.#cache.clear();
  }

  renderSlide(slidePartName: string): RenderedSlide {
    const presentation = this.requirePresentation();
    const slide = findSlideByPartName(presentation, slidePartName);
    if (!slide) throw new Error(`Presentation slide not found: ${slidePartName}.`);
    const rendered = this.renderSlideData(presentation, slide);
    this.#cache.set(slidePartName, rendered);
    return rendered;
  }

  renderAll(): readonly RenderedSlide[] {
    const presentation = this.requirePresentation();
    return getSlides(presentation).map((slide) => {
      const rendered = this.renderSlideData(presentation, slide);
      this.#cache.set(rendered.slidePartName, rendered);
      return rendered;
    });
  }

  renderChanges(records: readonly PresentationMutationRecord[]): readonly RenderedSlide[] {
    const presentation = this.requirePresentation();
    const dirty = new Set<string>();
    let renderAll = false;
    for (const record of records) {
      const name = record.partName;
      if (/^\/ppt\/slides\/slide[^/]+\.xml$/i.test(name)) {
        dirty.add(name);
        continue;
      }
      const relatedSlide = slidePartFromRelationshipPart(name);
      if (relatedSlide) {
        dirty.add(relatedSlide);
        continue;
      }
      const sourcePart = sourcePartFromRelationshipPart(name) ?? name;
      if (
        sourcePart === '/ppt/presentation.xml' ||
        sourcePart === '/ppt/_rels/presentation.xml.rels'
      ) {
        renderAll = true;
        continue;
      }
      for (const slide of slidesDependingOnPart(presentation, sourcePart)) {
        dirty.add(getSlidePartName(slide));
      }
      // Some malformed decks omit an image relationship but retain a shape
      // snapshot. Keep that recoverable case precise as well.
      if (/^\/ppt\/media\//i.test(name)) {
        for (const slide of getSlides(presentation)) {
          if (getSlideShapes(slide).some((shape) => getShapeImagePartName(shape) === name)) {
            dirty.add(getSlidePartName(slide));
          }
        }
      }
    }
    if (renderAll) return this.renderAll();
    const rendered: RenderedSlide[] = [];
    for (const name of dirty) {
      const slide = findSlideByPartName(presentation, name);
      if (!slide) {
        this.#cache.delete(name);
        continue;
      }
      const next = this.renderSlideData(presentation, slide);
      this.#cache.set(name, next);
      rendered.push(next);
    }
    return rendered;
  }

  private requirePresentation(): PresentationData {
    if (!this.#presentation) throw new Error('PresentationRenderer is not connected.');
    return this.#presentation;
  }

  private renderSlideData(presentation: PresentationData, slide: SlideData): RenderedSlide {
    const size = getSlideSize(presentation);
    if (!size) throw new Error('Presentation has no finite slide size.');
    return {
      slidePartName: getSlidePartName(slide),
      widthEmu: size.width as number,
      heightEmu: size.height as number,
      svg: renderSlideSvg(presentation, slide, this.#options),
      diagnostics: [],
    };
  }
}
