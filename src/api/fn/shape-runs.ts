// Per-run text accessors.

import { parseRPrLikeElement, parseTextRunState, resolveDrawingColor } from './shape-color.ts';
import {
  type BulletStyle,
  type ParagraphAlignment,
  type TextFormat,
  type TextRunState,
  applyBulletToParagraph,
  applyRunFormat,
  applyRunState,
} from '../../internal/drawingml/index.ts';
import {
  basename,
  emptyRels,
  type ImageFormat,
  nextRelId,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import { REL_TYPES } from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  insertChildByRank,
  qname,
  text,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  type PresentationData,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_PART_NAME,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { commitAndRefresh, ensureSlideImageRelationship, requireTxBody } from './_helpers.ts';
import { getPresentationTheme } from './theme.ts';
import { getSlides } from './slide-query.ts';
import { findCNvPr, NAME_HLINK_CLICK_FN, type ShapeClickAction } from './embedded.ts';
import {
  hyperlinkRelationshipIds,
  removeUnreferencedSlideRelationships,
} from './hyperlink-relationships.ts';
import {
  getRunDirectFillImageBytes,
  getRunUnderlineFillImageBytes,
  setRunDirectFillImage,
  setRunUnderlineFillImage,
} from './run-underline-image.ts';
import {
  emuCoordinate32,
  emuPositiveCoordinate32,
  normalizeGuid,
  oneOf,
} from '../../internal/bounds.ts';

const NAME_TX_BODY = qname('p', 'txBody', NS.pml);

// ---------------------------------------------------------------------------
// Per-run text accessors.
//
// Lets callers reach into a shape's text body to read or format a
// specific paragraph or run. `applyFormatToAllRuns` covers the bulk-edit
// case; these helpers cover "make this one word red."

const NAME_A_P = qname('a', 'p', NS.dml);
const NAME_A_R = qname('a', 'r', NS.dml);
const NAME_A_BR = qname('a', 'br', NS.dml);
const NAME_A_FLD = qname('a', 'fld', NS.dml);
export const NAME_A_RPR = qname('a', 'rPr', NS.dml);
const NAME_A_T = qname('a', 't', NS.dml);
const NAME_A_END_PARA_RPR = qname('a', 'endParaRPr', NS.dml);

/** Direct paragraph properties carried inside one DrawingML text field. */
export interface ShapeFieldParagraphProperties {
  readonly align?: ParagraphAlignment;
  readonly fontAlign?: ParagraphFontAlignment;
  readonly level?: number;
  readonly marL?: number;
  readonly marR?: number;
  readonly indent?: number;
  readonly lineSpacing?:
    | { readonly kind: 'pct'; readonly value: number }
    | { readonly kind: 'pts'; readonly value: number };
  readonly spcBefPts?: number;
  readonly spcAftPts?: number;
  readonly rtl?: boolean;
  readonly defaultTabSize?: number;
  readonly tabStops?: readonly ParagraphTabStop[];
}

/** Direct character properties carried by one paragraph's terminal mark. */
export interface ShapeEndParagraphProperties {
  readonly format: TextFormat;
  readonly state: TextRunState | null;
}

/** Editable input for one paragraph's terminal character properties. */
export interface ShapeEndParagraphPropertiesInput {
  readonly format?: TextFormat;
  readonly state?: TextRunState | null;
}

/** One authorable inline paragraph element. */
export type ShapeParagraphElementInput =
  | {
      readonly kind: 'r';
      readonly text: string;
      readonly format?: TextFormat;
      readonly state?: TextRunState;
    }
  | {
      readonly kind: 'fld';
      readonly id: string;
      readonly type?: string | null;
      readonly text: string;
      readonly format?: TextFormat;
      readonly state?: TextRunState;
      readonly paragraph?: ShapeFieldParagraphProperties;
    }
  | { readonly kind: 'br' };

const paragraphsOf = (txBody: XmlElement): XmlElement[] =>
  txBody.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' &&
      c.name.namespaceURI === NAME_A_P.namespaceURI &&
      c.name.localName === 'p',
  );

export const runsOf = (paragraph: XmlElement): XmlElement[] =>
  paragraph.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' &&
      c.name.namespaceURI === NAME_A_R.namespaceURI &&
      c.name.localName === 'r',
  );

export const requireParagraph = (shape: SlideShapeData, paragraphIndex: number): XmlElement => {
  const txBody = requireTxBody(shape);
  const paragraphs = paragraphsOf(txBody);
  const paragraph = paragraphs[paragraphIndex];
  if (!paragraph) {
    throw new RangeError(
      `paragraph index ${paragraphIndex} out of range (have ${paragraphs.length})`,
    );
  }
  return paragraph;
};

export const requireRun = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): XmlElement => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const runs = runsOf(paragraph);
  const run = runs[runIndex];
  if (!run) {
    throw new RangeError(
      `run index ${runIndex} out of range in paragraph ${paragraphIndex} (have ${runs.length})`,
    );
  }
  return run;
};

export const ensureRPr = (run: XmlElement): XmlElement => {
  const existing = firstChildElement(run, NAME_A_RPR);
  if (existing !== null) return existing;
  // `<a:rPr>` is the first child of `<a:r>` per the schema.
  const fresh = elem(NAME_A_RPR);
  run.children.unshift(fresh);
  return fresh;
};

const readRunText = (run: XmlElement): string => {
  const tEl = firstChildElement(run, NAME_A_T);
  if (tEl === null) return '';
  let out = '';
  for (const child of tEl.children) {
    if (child.kind === 'text' || child.kind === 'cdata') out += child.data;
  }
  return out;
};

const writeRunText = (run: XmlElement, value: string): void => {
  let tEl = firstChildElement(run, NAME_A_T);
  if (tEl === null) {
    tEl = elem(NAME_A_T);
    run.children.push(tEl);
  }
  tEl.children = [{ kind: 'text', data: value }];
};

/** Replace one paragraph's ordered runs and native line breaks. */
export const setShapeParagraphElements = (
  shape: SlideShapeData,
  paragraphIndex: number,
  elements: readonly ShapeParagraphElementInput[],
): void => {
  if (elements.length === 0) {
    throw new RangeError('setShapeParagraphElements requires at least one element.');
  }
  const paragraph = requireParagraph(shape, paragraphIndex);
  const oldRelationshipIds = hyperlinkRelationshipIds(paragraph);
  const paragraphProperties = firstChildElement(paragraph, NAME_A_PPR);
  const endProperties = firstChildElement(paragraph, NAME_A_END_PARA_RPR);
  const authoredElements = elements.map((value) => {
    if (value.kind === 'br') {
      return elem(NAME_A_BR);
    }
    const runProperties = value.format || value.state ? elem(NAME_A_RPR) : null;
    if (runProperties && value.format) applyRunFormat(runProperties, value.format);
    if (runProperties && value.state) applyRunState(runProperties, value.state);
    const textElement = elem(NAME_A_T, {
      children: [text(value.text)],
    });
    if (value.kind === 'fld') {
      if (value.type !== undefined && value.type !== null && typeof value.type !== 'string') {
        throw new TypeError('text field type must be a string, null, or undefined.');
      }
      const fieldParagraph =
        value.paragraph === undefined ? null : buildFieldParagraphProperties(value.paragraph);
      return elem(NAME_A_FLD, {
        attrs: [
          attr(qname('', 'id', ''), normalizeGuid(value.id, 'text field id')),
          ...(value.type === undefined || value.type === null
            ? []
            : [attr(qname('', 'type', ''), value.type)]),
        ],
        children: [
          ...(runProperties ? [runProperties] : []),
          ...(fieldParagraph ? [fieldParagraph] : []),
          textElement,
        ],
      });
    }
    return elem(NAME_A_R, {
      children: [...(runProperties ? [runProperties] : []), textElement],
    });
  });
  paragraph.children = [
    ...(paragraphProperties ? [paragraphProperties] : []),
    ...authoredElements,
    ...(endProperties ? [endProperties] : []),
  ];
  removeUnreferencedSlideRelationships(shape[SHAPE_SLIDE], oldRelationshipIds);
  commitAndRefresh(shape);
};

/** Number of paragraphs in the shape's text body. Throws for non-text shapes. */
export const getShapeParagraphCount = (shape: SlideShapeData): number =>
  paragraphsOf(requireTxBody(shape)).length;

/**
 * One inline element in a paragraph as ordered: a literal text run
 * (`<a:r>`), a field substitution (`<a:fld>` — slide number, date, etc.),
 * or a line break (`<a:br>`). Renderers walk this list instead of the
 * strict `<a:r>`-only `getShapeRunCount` / `getShapeRunText` pair when
 * they need to reproduce the paragraph's full visible content.
 *
 * `text` is the cached value (`<a:t>` content for `r` and `fld`; `''`
 * for `br`). `format` is the literal `<a:rPr>` on the element when
 * present; use `getShapeRunFormatEffective` to walk inheritance.
 *
 * Field kinds (`fld.type`): ECMA-376 deliberately types this attribute as
 * an arbitrary string. PowerPoint 16.111.2 inserts `slidenum`,
 * `datetime1` through `datetime13`, and `datetimeFigureOut`; its footer and
 * header UI uses PresentationML placeholders/settings rather than extra
 * DrawingML field tokens. Unrecognised strings come through unchanged so
 * renderers can decide whether to substitute live values.
 */
export type ShapeParagraphElement =
  | { readonly kind: 'r'; readonly text: string; readonly format: TextFormat | null }
  | {
      readonly kind: 'fld';
      readonly id: string | null;
      readonly text: string;
      readonly format: TextFormat | null;
      readonly state: TextRunState | null;
      readonly paragraph: ShapeFieldParagraphProperties | null;
      readonly type: string | null;
      /** Unsupported or malformed direct field semantics that an editable importer must diagnose. */
      readonly unsupported?: readonly string[];
    }
  | { readonly kind: 'br'; readonly format: TextFormat | null };

/**
 * Returns the inline children of a paragraph in document order — runs,
 * field placeholders, and line breaks. Used by renderers that need to
 * reproduce the paragraph faithfully (the `<a:r>`-only run accessors
 * silently drop fields and breaks).
 */
export const getShapeParagraphElements = (
  shape: SlideShapeData,
  paragraphIndex: number,
): ReadonlyArray<ShapeParagraphElement> =>
  readParagraphElements(requireParagraph(shape, paragraphIndex));

/** Read the direct `<a:endParaRPr>` properties, preserving authored-empty presence. */
export const getShapeEndParagraphProperties = (
  shape: SlideShapeData,
  paragraphIndex: number,
): ShapeEndParagraphProperties | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const endProperties = firstChildElement(paragraph, NAME_A_END_PARA_RPR);
  if (endProperties === null) return null;
  return {
    format: parseRPrLikeElement(endProperties) as TextFormat,
    state: parseTextRunState(endProperties),
  };
};

const stableJsonValue = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(stableJsonValue);
  if (value === null || typeof value !== 'object') return value;
  return Object.fromEntries(
    Object.entries(value as Record<string, unknown>)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([key, item]) => [key, stableJsonValue(item)]),
  );
};

const sameEndParagraphProperties = (
  current: ShapeEndParagraphProperties,
  requested: ShapeEndParagraphPropertiesInput,
): boolean =>
  JSON.stringify(stableJsonValue(current)) ===
  JSON.stringify(
    stableJsonValue({
      format: requested.format ?? {},
      state: requested.state ?? null,
    }),
  );

/** Replace or remove the direct `<a:endParaRPr>` terminal-mark properties. */
export const setShapeEndParagraphProperties = (
  shape: SlideShapeData,
  paragraphIndex: number,
  properties: ShapeEndParagraphPropertiesInput | null,
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const previous = firstChildElement(paragraph, NAME_A_END_PARA_RPR);
  if (previous === null && properties === null) return;
  if (previous !== null && properties !== null) {
    const current: ShapeEndParagraphProperties = {
      format: parseRPrLikeElement(previous) as TextFormat,
      state: parseTextRunState(previous),
    };
    if (sameEndParagraphProperties(current, properties)) return;
  }
  const oldRelationshipIds =
    previous === null ? new Set<string>() : hyperlinkRelationshipIds(previous);
  paragraph.children = paragraph.children.filter(
    (child) =>
      !(
        child.kind === 'element' &&
        child.name.namespaceURI === NAME_A_END_PARA_RPR.namespaceURI &&
        child.name.localName === NAME_A_END_PARA_RPR.localName
      ),
  );
  if (properties !== null) {
    const endProperties = elem(NAME_A_END_PARA_RPR);
    if (properties.format) applyRunFormat(endProperties, properties.format);
    if (properties.state) applyRunState(endProperties, properties.state);
    paragraph.children.push(endProperties);
  }
  removeUnreferencedSlideRelationships(shape[SHAPE_SLIDE], oldRelationshipIds);
  commitAndRefresh(shape);
};

/**
 * Walks a single `<a:p>` element and returns its inline children in
 * document order. Shared by the shape-text reader above and the table-cell
 * text reader: both use the identical DrawingML run/field/break grammar, so
 * only the way the paragraph element is located differs.
 *
 * @internal
 */
export const readParagraphElements = (
  paragraph: XmlElement,
  ctx?: Parameters<typeof parseRPrLikeElement>[1],
): ReadonlyArray<ShapeParagraphElement> => {
  const out: ShapeParagraphElement[] = [];
  const readT = (parent: XmlElement): string => {
    const tEl = firstChildElement(parent, NAME_A_T);
    if (!tEl) return '';
    let acc = '';
    for (const c of tEl.children) {
      if (c.kind === 'text' || c.kind === 'cdata') acc += c.data;
    }
    return acc;
  };
  const readFmt = (parent: XmlElement): TextFormat | null => {
    const rPr = firstChildElement(parent, NAME_A_RPR);
    if (!rPr) return null;
    return parseRPrLikeElement(rPr, ctx) as TextFormat;
  };
  for (const child of paragraph.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    if (child.name.localName === 'r') {
      out.push({ kind: 'r', text: readT(child), format: readFmt(child) });
    } else if (child.name.localName === 'fld') {
      const id = getAttrValue(child, qname('', 'id', ''));
      const type = getAttrValue(child, qname('', 'type', ''));
      const rPr = firstChildElement(child, NAME_A_RPR);
      const pPr = firstChildElement(child, NAME_A_PPR);
      const unsupported = unsupportedFieldSemantics(child, pPr);
      out.push({
        kind: 'fld',
        id,
        text: readT(child),
        format: readFmt(child),
        state: rPr === null ? null : parseTextRunState(rPr),
        paragraph: pPr === null ? null : parseFieldParagraphProperties(pPr),
        type,
        ...(unsupported.length === 0 ? {} : { unsupported }),
      });
    } else if (child.name.localName === 'br') {
      out.push({ kind: 'br', format: readFmt(child) });
    }
  }
  return out;
};

/**
 * Number of text runs in the given paragraph. Throws on out-of-range
 * paragraph index or non-text shapes.
 */
export const getShapeRunCount = (shape: SlideShapeData, paragraphIndex: number): number =>
  runsOf(requireParagraph(shape, paragraphIndex)).length;

/** Visible text of a single run. */
export const getShapeRunText = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): string => readRunText(requireRun(shape, paragraphIndex, runIndex));

/** Which native run interaction element carries the hyperlink payload. */
export type ShapeRunHyperlinkActivation = 'click' | 'hover';

/** Relationship-backed destination carried by a run hyperlink. */
export type ShapeRunHyperlinkTarget =
  | { readonly kind: 'url'; readonly url: string }
  | { readonly kind: 'slide'; readonly slide: SlideData };

/** Embedded WAV-compatible sound played by a native hyperlink interaction. */
export interface ShapeRunHyperlinkSound {
  readonly data: Uint8Array;
  readonly contentType: 'audio/mpeg' | 'audio/mp4' | 'audio/wav' | 'audio/x-wav';
  readonly name?: string;
}

const isShapeRunHyperlinkSoundContentType = (
  value: string,
): value is ShapeRunHyperlinkSound['contentType'] =>
  value === 'audio/mpeg' ||
  value === 'audio/mp4' ||
  value === 'audio/wav' ||
  value === 'audio/x-wav';

/** Editable DrawingML `CT_Hyperlink` attributes for one run interaction. */
export interface ShapeRunHyperlinkDescriptor {
  readonly target?: ShapeRunHyperlinkTarget;
  /** Exact native action URI. Unknown vendor actions are intentionally preserved. */
  readonly action?: string;
  readonly invalidUrl?: string;
  readonly targetFrame?: string;
  readonly tooltip?: string;
  readonly history?: boolean;
  readonly highlightClick?: boolean;
  readonly endSound?: boolean;
  readonly sound?: ShapeRunHyperlinkSound;
}

const hyperlinkElementName = (activation: ShapeRunHyperlinkActivation) =>
  qname('a', activation === 'click' ? 'hlinkClick' : 'hlinkMouseOver', NS.dml);

const hyperlinkChildRank = (element: XmlElement): number => {
  if (element.name.namespaceURI !== NS.dml) return 99;
  if (element.name.localName === 'hlinkClick') return 10;
  if (element.name.localName === 'hlinkMouseOver') return 11;
  if (element.name.localName === 'rtl') return 12;
  if (element.name.localName === 'extLst') return 13;
  return 9;
};

const readHyperlinkBoolean = (element: XmlElement, name: string): boolean | undefined => {
  const value = getAttrValue(element, qname('', name, ''));
  if (value === null) return undefined;
  return value === '1' || value === 'true';
};

const resolveHyperlinkTarget = (
  shape: SlideShapeData,
  relationshipId: string | null,
): ShapeRunHyperlinkTarget | undefined => {
  if (relationshipId === null || relationshipId === '') return undefined;
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const relationship = pkg
    .getRels(slide[SLIDE_PART_NAME])
    ?.items.find((candidate) => candidate.id === relationshipId);
  if (!relationship) return undefined;
  if (relationship.type === REL_TYPES.hyperlink && relationship.targetMode === 'External') {
    return { kind: 'url', url: relationship.target };
  }
  if (relationship.type !== REL_TYPES.slide || relationship.targetMode !== 'Internal') {
    return undefined;
  }
  const targetPartName = relationship.target.startsWith('/')
    ? partName(relationship.target)
    : resolveTarget(slide[SLIDE_PART_NAME], relationship.target);
  const presentation: PresentationData = { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
  const targetSlide = getSlides(presentation).find(
    (candidate) => candidate[SLIDE_PART_NAME] === targetPartName,
  );
  return targetSlide ? { kind: 'slide', slide: targetSlide } : undefined;
};

/** Read one complete click or hover `CT_Hyperlink` payload without flattening unknown actions. */
export const getShapeRunHyperlinkDescriptor = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
  activation: ShapeRunHyperlinkActivation,
): ShapeRunHyperlinkDescriptor | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, NAME_A_RPR);
  if (!rPr) return null;
  const hyperlink = firstChildElement(rPr, hyperlinkElementName(activation));
  if (!hyperlink) return null;
  const target = resolveHyperlinkTarget(
    shape,
    getAttrValue(hyperlink, qname('r', 'id', NS.officeDocRels)),
  );
  const read = (name: string): string | undefined =>
    getAttrValue(hyperlink, qname('', name, '')) ?? undefined;
  const action = read('action');
  const invalidUrl = read('invalidUrl');
  const targetFrame = read('tgtFrame');
  const tooltip = read('tooltip');
  const history = readHyperlinkBoolean(hyperlink, 'history');
  const highlightClick = readHyperlinkBoolean(hyperlink, 'highlightClick');
  const endSound = readHyperlinkBoolean(hyperlink, 'endSnd');
  const soundElement = firstChildElement(hyperlink, qname('a', 'snd', NS.dml));
  const soundRelationshipId = soundElement
    ? getAttrValue(soundElement, qname('r', 'embed', NS.officeDocRels))
    : null;
  const slide = shape[SHAPE_SLIDE];
  const soundRelationship = soundRelationshipId
    ? slide[INTERNAL_PACKAGE]
        .getRels(slide[SLIDE_PART_NAME])
        ?.items.find((candidate) => candidate.id === soundRelationshipId)
    : undefined;
  const soundPartName =
    soundRelationship?.type === REL_TYPES.audio && soundRelationship.targetMode === 'Internal'
      ? resolveTarget(slide[SLIDE_PART_NAME], soundRelationship.target)
      : null;
  const soundPart = soundPartName ? slide[INTERNAL_PACKAGE].getPart(partName(soundPartName)) : null;
  const sound =
    soundPart && isShapeRunHyperlinkSoundContentType(soundPart.contentType)
      ? {
          data: soundPart.data,
          contentType: soundPart.contentType,
          ...(soundElement && getAttrValue(soundElement, qname('', 'name', '')) !== null
            ? { name: getAttrValue(soundElement, qname('', 'name', '')) ?? '' }
            : {}),
        }
      : undefined;
  return {
    ...(target ? { target } : {}),
    ...(action === undefined ? {} : { action }),
    ...(invalidUrl === undefined ? {} : { invalidUrl }),
    ...(targetFrame === undefined ? {} : { targetFrame }),
    ...(tooltip === undefined ? {} : { tooltip }),
    ...(history === undefined ? {} : { history }),
    ...(highlightClick === undefined ? {} : { highlightClick }),
    ...(endSound === undefined ? {} : { endSound }),
    ...(sound === undefined ? {} : { sound }),
  };
};

/** Set or clear one complete run-level click/hover hyperlink payload. */
export const setShapeRunHyperlinkDescriptor = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
  activation: ShapeRunHyperlinkActivation,
  descriptor: ShapeRunHyperlinkDescriptor | null,
): void => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = ensureRPr(run);
  const name = hyperlinkElementName(activation);
  const oldRelationshipIds = new Set<string>();
  for (const child of rPr.children) {
    if (
      child.kind === 'element' &&
      child.name.namespaceURI === name.namespaceURI &&
      child.name.localName === name.localName
    ) {
      for (const id of hyperlinkRelationshipIds(child)) oldRelationshipIds.add(id);
    }
  }
  rPr.children = rPr.children.filter(
    (child) =>
      !(
        child.kind === 'element' &&
        child.name.namespaceURI === name.namespaceURI &&
        child.name.localName === name.localName
      ),
  );

  const slide = shape[SHAPE_SLIDE];
  if (descriptor !== null) {
    const attrs = [] as Array<ReturnType<typeof attr>>;
    const target = descriptor.target;
    if (target) {
      const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
      const existing = rels.items.find((relationship) =>
        target.kind === 'url'
          ? relationship.type === REL_TYPES.hyperlink &&
            relationship.targetMode === 'External' &&
            relationship.target === target.url
          : relationship.type === REL_TYPES.slide &&
            relationship.targetMode === 'Internal' &&
            relationship.target === `../slides/${basename(target.slide[SLIDE_PART_NAME])}`,
      );
      const relationshipId =
        existing?.id ?? nextRelId(rels.items.map((relationship) => relationship.id));
      if (!existing) {
        rels.items.push(
          target.kind === 'url'
            ? {
                id: relationshipId,
                type: REL_TYPES.hyperlink,
                target: target.url,
                targetMode: 'External',
              }
            : {
                id: relationshipId,
                type: REL_TYPES.slide,
                target: `../slides/${basename(target.slide[SLIDE_PART_NAME])}`,
                targetMode: 'Internal',
              },
        );
        slide[INTERNAL_PACKAGE].setRels(slide[SLIDE_PART_NAME], rels);
      }
      attrs.push(attr(qname('r', 'id', NS.officeDocRels), relationshipId));
    } else {
      // PowerPoint writes an explicit empty relationship ID for action-only
      // interactions and repairs run hyperlinks that omit the attribute.
      attrs.push(attr(qname('r', 'id', NS.officeDocRels), ''));
    }
    const addString = (name: string, value: string | undefined): void => {
      if (value !== undefined) attrs.push(attr(qname('', name, ''), value));
    };
    const addBoolean = (name: string, value: boolean | undefined): void => {
      if (value !== undefined) attrs.push(attr(qname('', name, ''), value ? '1' : '0'));
    };
    addString('invalidUrl', descriptor.invalidUrl);
    addString('action', descriptor.action);
    addString('tgtFrame', descriptor.targetFrame);
    addString('tooltip', descriptor.tooltip);
    addBoolean('history', descriptor.history);
    addBoolean('highlightClick', descriptor.highlightClick);
    addBoolean('endSnd', descriptor.endSound);
    const children: XmlElement[] = [];
    if (descriptor.sound) {
      const extension = {
        'audio/mpeg': 'mp3',
        'audio/mp4': 'm4a',
        'audio/wav': 'wav',
        'audio/x-wav': 'wav',
      }[descriptor.sound.contentType];
      let mediaNumber = 1;
      for (const part of slide[INTERNAL_PACKAGE].parts) {
        const value = /^\/ppt\/media\/media(\d+)\./u.exec(part.name)?.[1];
        if (value !== undefined) mediaNumber = Math.max(mediaNumber, Number(value) + 1);
      }
      const mediaPartName = partName(`/ppt/media/media${mediaNumber}.${extension}`);
      slide[INTERNAL_PACKAGE].addPart(
        mediaPartName,
        descriptor.sound.contentType,
        descriptor.sound.data,
      );
      const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
      const soundRelationshipId = nextRelId(rels.items.map((relationship) => relationship.id));
      rels.items.push({
        id: soundRelationshipId,
        type: REL_TYPES.audio,
        target: `../media/media${mediaNumber}.${extension}`,
        targetMode: 'Internal',
      });
      slide[INTERNAL_PACKAGE].setRels(slide[SLIDE_PART_NAME], rels);
      children.push(
        elem(qname('a', 'snd', NS.dml), {
          attrs: [
            attr(qname('r', 'embed', NS.officeDocRels), soundRelationshipId),
            ...(descriptor.sound.name === undefined
              ? []
              : [attr(qname('', 'name', ''), descriptor.sound.name)]),
          ],
        }),
      );
    }
    insertChildByRank(rPr, elem(name, { attrs, children }), hyperlinkChildRank);
  }
  removeUnreferencedSlideRelationships(slide, oldRelationshipIds);
  commitAndRefresh(shape);
};

/**
 * Sets `<a:hlinkClick>` on a single run. Per-run counterpart to
 * `setShapeHyperlink` (which targets every run in the shape). Pass
 * `null` to clear the link on that run alone — other runs are
 * untouched. Allocates or reuses a hyperlink rel on the slide
 * exactly like the shape-level setter.
 */
export const setShapeRunHyperlink = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
  url: string | null,
  tooltip?: string,
): void => {
  setShapeRunHyperlinkDescriptor(
    shape,
    paragraphIndex,
    runIndex,
    'click',
    url === null
      ? null
      : { target: { kind: 'url', url }, ...(tooltip === undefined ? {} : { tooltip }) },
  );
};

/**
 * Reads the external URL on a single run's `<a:hlinkClick>`. Per-run
 * counterpart to `getShapeHyperlink` (which only surfaces the first
 * link it finds). Returns `null` when this run has no link, or the
 * link's `r:id` resolves to a non-hyperlink rel.
 */
export const getShapeRunHyperlink = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): string | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, qname('a', 'rPr', NS.dml));
  if (!rPr) return null;
  const hlink = firstChildElement(rPr, qname('a', 'hlinkClick', NS.dml));
  if (!hlink) return null;
  const rId = getAttrValue(hlink, qname('r', 'id', NS.officeDocRels));
  if (!rId) return null;
  const slide = shape[SHAPE_SLIDE];
  const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const rel = rels.items.find((x) => x.id === rId);
  if (rel?.type === REL_TYPES.hyperlink && rel.targetMode === 'External') return rel.target;
  return null;
};

/**
 * Reads the tooltip text on the shape's `<a:hlinkClick tooltip="…"/>`.
 * Returns `null` when no hyperlink is set or the link doesn't author
 * a tooltip. Tooltips show up in PowerPoint when the user hovers over
 * a linked shape in slide-show mode.
 *
 * Scans run-level `<a:rPr><a:hlinkClick>` first (where
 * `setShapeHyperlink` writes) and falls back to the
 * `<p:nvSpPr><p:cNvPr><a:hlinkClick>` shape-click hyperlink. Mirrors
 * `getShapeHyperlink`'s read path so the writer / reader pair is
 * consistent.
 */
export const getShapeHyperlinkTooltip = (shape: SlideShapeData): string | null => {
  if (shape[SHAPE_SNAPSHOT].kind === 'shape') {
    const txBody = firstChildElement(shape[SHAPE_ELEMENT], NAME_TX_BODY);
    if (txBody) {
      for (const p of txBody.children) {
        if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p')
          continue;
        for (const r of p.children) {
          if (r.kind !== 'element' || r.name.namespaceURI !== NS.dml || r.name.localName !== 'r')
            continue;
          const rPr = firstChildElement(r, qname('a', 'rPr', NS.dml));
          if (!rPr) continue;
          const hlink = firstChildElement(rPr, qname('a', 'hlinkClick', NS.dml));
          if (!hlink) continue;
          const tt = getAttrValue(hlink, qname('', 'tooltip', ''));
          if (tt !== null) return tt;
        }
      }
    }
  }
  const cNvPr = findCNvPr(shape);
  if (!cNvPr) return null;
  const hlink = firstChildElement(cNvPr, NAME_HLINK_CLICK_FN);
  if (!hlink) return null;
  const tt = getAttrValue(hlink, qname('', 'tooltip', ''));
  return tt ?? null;
};

/**
 * Reads the tooltip on a per-run `<a:rPr><a:hlinkClick tooltip="…"/>`.
 * Same semantics as `getShapeHyperlinkTooltip` but scoped to a single
 * run.
 */
export const getShapeRunHyperlinkTooltip = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): string | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, qname('a', 'rPr', NS.dml));
  if (!rPr) return null;
  const hlink = firstChildElement(rPr, qname('a', 'hlinkClick', NS.dml));
  if (!hlink) return null;
  const tt = getAttrValue(hlink, qname('', 'tooltip', ''));
  return tt ?? null;
};

/**
 * Same as `getShapeClickAction` but reads the per-run
 * `<a:rPr><a:hlinkClick action=… r:id=…/>`. Recognises:
 *
 *   - `{ kind: 'url', url }` — external hyperlink rel
 *   - `{ kind: 'slide', slide }` — slide-jump action + slide rel
 *   - `{ kind: 'nextSlide' | 'prevSlide' | 'firstSlide' | 'lastSlide' }`
 *
 * Returns `null` for runs without an action or unknown action tokens.
 */
export const getShapeRunClickAction = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): ShapeClickAction | null => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  const rPr = firstChildElement(run, qname('a', 'rPr', NS.dml));
  if (!rPr) return null;
  const hlink = firstChildElement(rPr, qname('a', 'hlinkClick', NS.dml));
  if (!hlink) return null;
  const action = getAttrValue(hlink, qname('', 'action', ''));
  const rId = getAttrValue(hlink, qname('r', 'id', NS.officeDocRels));

  if (action === 'ppaction://hlinkshowjump?jump=nextslide') return { kind: 'nextSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=previousslide') return { kind: 'prevSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=firstslide') return { kind: 'firstSlide' };
  if (action === 'ppaction://hlinkshowjump?jump=lastslide') return { kind: 'lastSlide' };

  if (rId === null || rId === '') return null;
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const rel = rels.items.find((r) => r.id === rId);
  if (!rel) return null;
  if (action === 'ppaction://hlinksldjump' && rel.type === REL_TYPES.slide) {
    const targetPartName = rel.target.startsWith('/')
      ? partName(rel.target)
      : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
    const pres: PresentationData = { [INTERNAL_PACKAGE]: pkg, _slidesCache: null };
    for (const candidate of getSlides(pres)) {
      if (candidate[SLIDE_PART_NAME] === targetPartName) return { kind: 'slide', slide: candidate };
    }
    return null;
  }
  if (rel.type === REL_TYPES.hyperlink && rel.targetMode === 'External') {
    return { kind: 'url', url: rel.target };
  }
  return null;
};

export const NAME_A_PPR = qname('a', 'pPr', NS.dml);
export const ATTR_LVL = qname('', 'lvl', '');
const ATTR_ALGN_FN = qname('', 'algn', '');
const ATTR_RTL_FN = qname('', 'rtl', '');
const ATTR_FONT_ALGN_FN = qname('', 'fontAlgn', '');

const ensurePPr = (paragraph: XmlElement): XmlElement => {
  const existing = firstChildElement(paragraph, NAME_A_PPR);
  if (existing !== null) return existing;
  const fresh = elem(NAME_A_PPR);
  // <a:pPr> must be the first child of <a:p>.
  paragraph.children.unshift(fresh);
  return fresh;
};

const alignTokenForFn = (a: ParagraphAlignment): string => {
  switch (a) {
    case 'left':
    case 'l':
      return 'l';
    case 'center':
    case 'ctr':
      return 'ctr';
    case 'right':
    case 'r':
      return 'r';
    case 'justify':
    case 'just':
      return 'just';
    case 'distribute':
    case 'dist':
      return 'dist';
    default:
      return a;
  }
};

/**
 * Sets the horizontal alignment of a single paragraph. Same token set
 * as `setShapeAlignment`. Other paragraphs are untouched.
 */
export const setParagraphAlignment = (
  shape: SlideShapeData,
  paragraphIndex: number,
  align: ParagraphAlignment,
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);
  pPr.attrs = pPr.attrs.filter((a) => a.name.localName !== 'algn');
  pPr.attrs.push(attr(ATTR_ALGN_FN, alignTokenForFn(align)));
  commitAndRefresh(shape);
};

/**
 * Sets the paragraph's nesting level (`<a:pPr lvl="N"/>`). Levels are
 * 0-indexed; PowerPoint accepts 0 through 8. Pass `0` to clear an
 * existing level — `<a:pPr lvl="0"/>` is the same as omitting the attr.
 *
 * Used in tandem with bullets to author nested lists:
 *
 *   setShapeText(shape, 'Item 1\nNested\nItem 2');
 *   setShapeBullets(shape, 'bullet');
 *   setParagraphLevel(shape, 1, 1);  // indent the second line
 */
export const setParagraphLevel = (
  shape: SlideShapeData,
  paragraphIndex: number,
  level: number,
): void => {
  if (!Number.isInteger(level) || level < 0 || level > 8) {
    throw new RangeError(`paragraph level must be an integer in [0, 8], got ${level}`);
  }
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);
  pPr.attrs = pPr.attrs.filter((a) => a.name.localName !== 'lvl');
  if (level > 0) pPr.attrs.push(attr(ATTR_LVL, String(level)));
  commitAndRefresh(shape);
};

/**
 * Reads the paragraph's horizontal alignment. Returns `null` when no
 * `algn` attribute is present (inherits from layout / master).
 */
export const getParagraphAlignment = (
  shape: SlideShapeData,
  paragraphIndex: number,
): ParagraphAlignment | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (pPr === null) return null;
  const v = getAttrValue(pPr, ATTR_ALGN_FN);
  return (v as ParagraphAlignment | null) ?? null;
};

/**
 * Sets or clears one paragraph's right-to-left flag (`<a:pPr rtl="…"/>`).
 * `true` and `false` write explicit OOXML booleans; `null` restores inheritance.
 */
export const setParagraphRtl = (
  shape: SlideShapeData,
  paragraphIndex: number,
  value: boolean | null,
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const existing = firstChildElement(paragraph, NAME_A_PPR);
  // Clearing an inherited value is a no-op when the paragraph has no direct
  // properties; avoid materialising an empty <a:pPr> solely for that no-op.
  if (value === null && existing === null) return;
  const pPr = existing ?? ensurePPr(paragraph);
  pPr.attrs = pPr.attrs.filter(
    (candidate) =>
      !(
        candidate.name.namespaceURI === ATTR_RTL_FN.namespaceURI &&
        candidate.name.localName === 'rtl'
      ),
  );
  if (value !== null) pPr.attrs.push(attr(ATTR_RTL_FN, value ? '1' : '0'));
  commitAndRefresh(shape);
};

/** Reads a paragraph's direct RTL flag, or `null` when it inherits the value. */
export const getParagraphRtl = (shape: SlideShapeData, paragraphIndex: number): boolean | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (pPr === null) return null;
  const value = getAttrValue(pPr, ATTR_RTL_FN);
  if (value === '1' || value === 'true') return true;
  if (value === '0' || value === 'false') return false;
  return null;
};

/** DrawingML alignment for differently sized fonts on one paragraph line. */
export type ParagraphFontAlignment = 'auto' | 'top' | 'center' | 'baseline' | 'bottom';

const PARAGRAPH_FONT_ALIGNMENT_TO_TOKEN: Record<ParagraphFontAlignment, string> = {
  auto: 'auto',
  top: 't',
  center: 'ctr',
  baseline: 'base',
  bottom: 'b',
};

const PARAGRAPH_FONT_TOKEN_TO_ALIGNMENT: Readonly<Record<string, ParagraphFontAlignment>> = {
  auto: 'auto',
  t: 'top',
  ctr: 'center',
  base: 'baseline',
  b: 'bottom',
};

const PARAGRAPH_FONT_ALIGNMENTS = ['auto', 'top', 'center', 'baseline', 'bottom'] as const;

/** @internal Parses a direct or inherited pPr font-alignment token. */
export const parseParagraphFontAlignment = (pPr: XmlElement): ParagraphFontAlignment | null => {
  const token = getAttrValue(pPr, ATTR_FONT_ALGN_FN);
  return token === null ? null : (PARAGRAPH_FONT_TOKEN_TO_ALIGNMENT[token] ?? null);
};

/** Reads a paragraph's direct font alignment, or `null` when it inherits. */
export const getParagraphFontAlignment = (
  shape: SlideShapeData,
  paragraphIndex: number,
): ParagraphFontAlignment | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  return pPr === null ? null : parseParagraphFontAlignment(pPr);
};

/** Authors one paragraph's font alignment; `null` restores inheritance. */
export const setParagraphFontAlignment = (
  shape: SlideShapeData,
  paragraphIndex: number,
  alignment: ParagraphFontAlignment | null,
): void => {
  const normalized =
    alignment === null
      ? null
      : oneOf(alignment, PARAGRAPH_FONT_ALIGNMENTS, 'paragraph font alignment');
  const paragraph = requireParagraph(shape, paragraphIndex);
  const existing = firstChildElement(paragraph, NAME_A_PPR);
  if (normalized === null && existing === null) return;
  const pPr = existing ?? ensurePPr(paragraph);
  pPr.attrs = pPr.attrs.filter(
    (candidate) => !(candidate.name.namespaceURI === '' && candidate.name.localName === 'fontAlgn'),
  );
  if (normalized !== null) {
    pPr.attrs.push(attr(ATTR_FONT_ALGN_FN, PARAGRAPH_FONT_ALIGNMENT_TO_TOKEN[normalized]));
  }
  commitAndRefresh(shape);
};

/**
 * Reads the paragraph's nesting level (`lvl` attribute), or `0` when
 * absent — PowerPoint's default. Returns `null` for non-existent
 * paragraphs.
 */
export const getParagraphLevel = (shape: SlideShapeData, paragraphIndex: number): number => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (pPr === null) return 0;
  const v = getAttrValue(pPr, ATTR_LVL);
  if (v === null) return 0;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n : 0;
};

// CT_TextParagraphProperties (a:pPr) is an xsd:sequence: line spacing, then
// before/after spacing, then the bullet-related groups, then tabLst/defRPr.
// Setters strip their element then re-insert at the mandated slot.
const PPR_CHILD_RANK: Record<string, number> = {
  lnSpc: 0,
  spcBef: 1,
  spcAft: 2,
  buClrTx: 3,
  buClr: 3,
  buSzTx: 4,
  buSzPct: 4,
  buSzPts: 4,
  buFontTx: 5,
  buFont: 5,
  buNone: 6,
  buAutoNum: 6,
  buChar: 6,
  buBlip: 6,
  tabLst: 7,
  defRPr: 8,
  extLst: 9,
};
const pPrChildRank = (el: XmlElement): number =>
  el.name.namespaceURI === NS.dml ? (PPR_CHILD_RANK[el.name.localName] ?? 99) : 99;

/** Alignment of one DrawingML paragraph tab stop. */
export type ParagraphTabAlignment = 'left' | 'center' | 'right' | 'decimal';

/** One paragraph tab stop. Positions use DrawingML EMU coordinates. */
export interface ParagraphTabStop {
  readonly positionEmu: number;
  readonly alignment: ParagraphTabAlignment;
}

const TAB_ALIGNMENT_TO_TOKEN: Record<ParagraphTabAlignment, string> = {
  left: 'l',
  center: 'ctr',
  right: 'r',
  decimal: 'dec',
};

const TAB_TOKEN_TO_ALIGNMENT: Readonly<Record<string, ParagraphTabAlignment>> = {
  l: 'left',
  ctr: 'center',
  r: 'right',
  dec: 'decimal',
};

const PARAGRAPH_TAB_ALIGNMENTS = ['left', 'center', 'right', 'decimal'] as const;

/** @internal Parses one pPr layer while preserving authored-empty versus inherited tab lists. */
export const parseParagraphTabStops = (pPr: XmlElement): readonly ParagraphTabStop[] | null => {
  const tabList = firstChildElement(pPr, qname('a', 'tabLst', NS.dml));
  if (!tabList) return null;
  const stops: ParagraphTabStop[] = [];
  for (const child of tabList.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    if (child.name.localName !== 'tab') continue;
    const position = getAttrValue(child, qname('', 'pos', ''));
    const token = getAttrValue(child, qname('', 'algn', ''));
    const alignment = token === null ? undefined : TAB_TOKEN_TO_ALIGNMENT[token];
    if (position === null || alignment === undefined) continue;
    const positionEmu = Number.parseInt(position, 10);
    if (Number.isFinite(positionEmu)) stops.push({ positionEmu, alignment });
  }
  return stops;
};

const FIELD_PARAGRAPH_ALIGNMENT_FROM_TOKEN: Readonly<Record<string, ParagraphAlignment>> = {
  l: 'left',
  ctr: 'center',
  r: 'right',
  just: 'justify',
  justLow: 'justLow',
  dist: 'dist',
  thaiDist: 'thaiDist',
};

const FIELD_PARAGRAPH_SUPPORTED_ATTRIBUTES = new Set([
  'algn',
  'fontAlgn',
  'lvl',
  'marL',
  'marR',
  'indent',
  'rtl',
  'defTabSz',
]);
const FIELD_PARAGRAPH_SUPPORTED_CHILDREN = new Set(['lnSpc', 'spcBef', 'spcAft', 'tabLst']);

const integerLexeme = (value: string): number | null =>
  /^-?\d+$/u.test(value) && Number.isSafeInteger(Number(value)) ? Number(value) : null;

const unsupportedFieldParagraphSemantics = (pPr: XmlElement): string[] => {
  const unsupported: string[] = [];
  for (const attribute of pPr.attrs) {
    if (
      attribute.name.namespaceURI === '' &&
      !FIELD_PARAGRAPH_SUPPORTED_ATTRIBUTES.has(attribute.name.localName)
    ) {
      unsupported.push(`text field paragraph attribute ${attribute.name.localName}`);
    }
  }
  const validateIntegerAttribute = (
    name: 'lvl' | 'marL' | 'marR' | 'indent' | 'defTabSz',
    validate: (value: number) => boolean,
  ): void => {
    const raw = getAttrValue(pPr, qname('', name, ''));
    if (raw === null) return;
    const value = integerLexeme(raw);
    if (value === null || !validate(value)) {
      unsupported.push(`text field paragraph ${name} ${raw}`);
    }
  };
  const alignment = getAttrValue(pPr, ATTR_ALGN_FN);
  if (alignment !== null && FIELD_PARAGRAPH_ALIGNMENT_FROM_TOKEN[alignment] === undefined) {
    unsupported.push(`text field paragraph alignment ${alignment}`);
  }
  const fontAlignment = getAttrValue(pPr, ATTR_FONT_ALGN_FN);
  if (fontAlignment !== null && PARAGRAPH_FONT_TOKEN_TO_ALIGNMENT[fontAlignment] === undefined) {
    unsupported.push(`text field paragraph font alignment ${fontAlignment}`);
  }
  validateIntegerAttribute('lvl', (value) => value >= 0 && value <= 8);
  validateIntegerAttribute('marL', (value) => value >= 0 && value <= 2_147_483_647);
  validateIntegerAttribute('marR', (value) => value >= 0 && value <= 2_147_483_647);
  validateIntegerAttribute('indent', (value) => value >= -2_147_483_648 && value <= 2_147_483_647);
  validateIntegerAttribute('defTabSz', (value) => value >= 0 && value <= 2_147_483_647);
  const rtl = getAttrValue(pPr, ATTR_RTL_FN);
  if (rtl !== null && !['0', '1', 'false', 'true'].includes(rtl)) {
    unsupported.push(`text field paragraph rtl ${rtl}`);
  }

  const spacingChildren = (outer: XmlElement): XmlElement[] =>
    outer.children.filter((child): child is XmlElement => child.kind === 'element');
  const validateSpacing = (outer: XmlElement, allowPercent: boolean): void => {
    const children = spacingChildren(outer);
    if (children.length !== 1) {
      unsupported.push(`text field paragraph ${outer.name.localName} structure`);
      return;
    }
    const inner = children[0]!;
    if (
      inner.name.namespaceURI !== NS.dml ||
      (inner.name.localName !== 'spcPts' && !(allowPercent && inner.name.localName === 'spcPct'))
    ) {
      unsupported.push(`text field paragraph ${outer.name.localName} ${inner.name.localName}`);
      return;
    }
    if (
      inner.attrs.some(
        (attribute) => attribute.name.namespaceURI !== '' || attribute.name.localName !== 'val',
      )
    ) {
      unsupported.push(`text field paragraph ${outer.name.localName} spacing attributes`);
    }
    const raw = getAttrValue(inner, qname('', 'val', ''));
    const value = raw === null ? null : integerLexeme(raw);
    if (value === null || value < 0) {
      unsupported.push(`text field paragraph ${outer.name.localName} spacing ${String(raw)}`);
    }
  };

  for (const child of pPr.children) {
    if (child.kind !== 'element') continue;
    if (child.name.namespaceURI !== NS.dml || child.name.localName === 'extLst') continue;
    if (!FIELD_PARAGRAPH_SUPPORTED_CHILDREN.has(child.name.localName)) {
      unsupported.push(`text field paragraph child ${child.name.localName}`);
      continue;
    }
    if (child.name.localName === 'lnSpc') validateSpacing(child, true);
    else if (child.name.localName === 'spcBef' || child.name.localName === 'spcAft') {
      validateSpacing(child, false);
    } else if (child.name.localName === 'tabLst') {
      const tabs = child.children.filter(
        (candidate): candidate is XmlElement => candidate.kind === 'element',
      );
      if (tabs.length > 32) unsupported.push(`text field paragraph tab count ${tabs.length}`);
      for (const tab of tabs) {
        if (tab.name.namespaceURI !== NS.dml || tab.name.localName !== 'tab') {
          unsupported.push(`text field paragraph tab child ${tab.name.localName}`);
          continue;
        }
        if (
          tab.attrs.some(
            (attribute) =>
              attribute.name.namespaceURI !== '' ||
              !['pos', 'algn'].includes(attribute.name.localName),
          ) ||
          tab.children.some((candidate) => candidate.kind === 'element')
        ) {
          unsupported.push('text field paragraph tab structure');
        }
        const positionRaw = getAttrValue(tab, qname('', 'pos', ''));
        const position = positionRaw === null ? null : integerLexeme(positionRaw);
        if (position === null || position < -2_147_483_648 || position > 2_147_483_647) {
          unsupported.push(`text field paragraph tab position ${String(positionRaw)}`);
        }
        const token = getAttrValue(tab, qname('', 'algn', ''));
        if (token === null || TAB_TOKEN_TO_ALIGNMENT[token] === undefined) {
          unsupported.push(`text field paragraph tab alignment ${String(token)}`);
        }
      }
    }
  }
  return unsupported;
};

const unsupportedFieldSemantics = (field: XmlElement, pPr: XmlElement | null): string[] => {
  const unsupported = field.attrs
    .filter(
      (attribute) =>
        attribute.name.namespaceURI === '' && !['id', 'type'].includes(attribute.name.localName),
    )
    .map((attribute) => `text field attribute ${attribute.name.localName}`);
  const fieldId = getAttrValue(field, qname('', 'id', ''));
  if (fieldId !== null) {
    try {
      normalizeGuid(fieldId, 'text field id');
    } catch {
      unsupported.push(`text field identity ${fieldId}`);
    }
  }
  const childCounts = new Map<string, number>();
  for (const child of field.children) {
    if (child.kind !== 'element') continue;
    childCounts.set(child.name.localName, (childCounts.get(child.name.localName) ?? 0) + 1);
    if (child.name.namespaceURI !== NS.dml || child.name.localName === 'extLst') continue;
    if (!['rPr', 'pPr', 't'].includes(child.name.localName)) {
      unsupported.push(`text field child ${child.name.localName}`);
    }
  }
  for (const name of ['rPr', 'pPr', 't']) {
    const count = childCounts.get(name) ?? 0;
    if (count > 1) unsupported.push(`text field duplicate ${name}`);
  }
  if (pPr !== null) unsupported.push(...unsupportedFieldParagraphSemantics(pPr));
  return unsupported;
};

/** Parse one field-local pPr without applying the containing paragraph's inheritance cascade. */
const parseFieldParagraphProperties = (pPr: XmlElement): ShapeFieldParagraphProperties => {
  const properties: {
    -readonly [K in keyof ShapeFieldParagraphProperties]?: ShapeFieldParagraphProperties[K];
  } = {};
  const readIntegerAttr = (name: string): number | undefined => {
    const value = getAttrValue(pPr, qname('', name, ''));
    if (value === null) return undefined;
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) ? parsed : undefined;
  };
  const alignment = getAttrValue(pPr, ATTR_ALGN_FN);
  if (alignment !== null && FIELD_PARAGRAPH_ALIGNMENT_FROM_TOKEN[alignment] !== undefined) {
    properties.align = FIELD_PARAGRAPH_ALIGNMENT_FROM_TOKEN[alignment];
  }
  const fontAlign = parseParagraphFontAlignment(pPr);
  if (fontAlign !== null) properties.fontAlign = fontAlign;
  const level = readIntegerAttr('lvl');
  if (level !== undefined) properties.level = level;
  const marL = readIntegerAttr('marL');
  if (marL !== undefined) properties.marL = marL;
  const marR = readIntegerAttr('marR');
  if (marR !== undefined) properties.marR = marR;
  const indent = readIntegerAttr('indent');
  if (indent !== undefined) properties.indent = indent;
  const defaultTabSize = readIntegerAttr('defTabSz');
  if (defaultTabSize !== undefined) properties.defaultTabSize = defaultTabSize;
  const rtl = getAttrValue(pPr, ATTR_RTL_FN);
  if (rtl !== null) properties.rtl = rtl === '1' || rtl === 'true';
  const tabStops = parseParagraphTabStops(pPr);
  if (tabStops !== null) properties.tabStops = tabStops;

  const readPointSpacing = (name: 'spcBef' | 'spcAft'): number | undefined => {
    const outer = firstChildElement(pPr, qname('a', name, NS.dml));
    const inner = outer && firstChildElement(outer, qname('a', 'spcPts', NS.dml));
    const value = inner && getAttrValue(inner, qname('', 'val', ''));
    if (value === null || value === undefined) return undefined;
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) ? parsed / 100 : undefined;
  };
  const before = readPointSpacing('spcBef');
  if (before !== undefined) properties.spcBefPts = before;
  const after = readPointSpacing('spcAft');
  if (after !== undefined) properties.spcAftPts = after;
  const lineSpacing = firstChildElement(pPr, qname('a', 'lnSpc', NS.dml));
  if (lineSpacing) {
    const pct = firstChildElement(lineSpacing, qname('a', 'spcPct', NS.dml));
    const pts = firstChildElement(lineSpacing, qname('a', 'spcPts', NS.dml));
    const pctValue = pct && getAttrValue(pct, qname('', 'val', ''));
    const ptsValue = pts && getAttrValue(pts, qname('', 'val', ''));
    if (pctValue !== null && pctValue !== undefined) {
      const parsed = Number.parseFloat(pctValue);
      if (Number.isFinite(parsed)) {
        properties.lineSpacing = {
          kind: 'pct',
          value: Math.abs(parsed) > 1 ? parsed / 100000 : parsed,
        };
      }
    } else if (ptsValue !== null && ptsValue !== undefined) {
      const parsed = Number.parseInt(ptsValue, 10);
      if (Number.isFinite(parsed)) properties.lineSpacing = { kind: 'pts', value: parsed / 100 };
    }
  }
  return properties;
};

const FIELD_PARAGRAPH_ALIGNMENTS = [
  'left',
  'center',
  'right',
  'justify',
  'distribute',
  'l',
  'ctr',
  'r',
  'just',
  'dist',
  'justLow',
  'thaiDist',
] as const;

/** Build one schema-ordered field-local pPr from the public direct-property model. */
const buildFieldParagraphProperties = (value: ShapeFieldParagraphProperties): XmlElement => {
  const pPr = elem(NAME_A_PPR);
  if (value.align !== undefined) {
    const alignment = oneOf(
      value.align,
      FIELD_PARAGRAPH_ALIGNMENTS,
      'text field paragraph alignment',
    );
    pPr.attrs.push(attr(ATTR_ALGN_FN, alignTokenForFn(alignment)));
  }
  if (value.fontAlign !== undefined) {
    const alignment = oneOf(
      value.fontAlign,
      PARAGRAPH_FONT_ALIGNMENTS,
      'text field paragraph font alignment',
    );
    pPr.attrs.push(attr(ATTR_FONT_ALGN_FN, PARAGRAPH_FONT_ALIGNMENT_TO_TOKEN[alignment]));
  }
  if (value.level !== undefined) {
    if (!Number.isInteger(value.level) || value.level < 0 || value.level > 8) {
      throw new RangeError(
        `text field paragraph level must be an integer in [0, 8], got ${value.level}`,
      );
    }
    if (value.level > 0) pPr.attrs.push(attr(ATTR_LVL, String(value.level)));
  }
  if (value.marL !== undefined) {
    pPr.attrs.push(
      attr(
        qname('', 'marL', ''),
        String(emuPositiveCoordinate32(value.marL, 'text field paragraph marL')),
      ),
    );
  }
  if (value.marR !== undefined) {
    pPr.attrs.push(
      attr(
        qname('', 'marR', ''),
        String(emuPositiveCoordinate32(value.marR, 'text field paragraph marR')),
      ),
    );
  }
  if (value.indent !== undefined) {
    pPr.attrs.push(
      attr(
        qname('', 'indent', ''),
        String(emuCoordinate32(value.indent, 'text field paragraph indent')),
      ),
    );
  }
  if (value.rtl !== undefined) {
    if (typeof value.rtl !== 'boolean')
      throw new TypeError('text field paragraph rtl must be a boolean.');
    pPr.attrs.push(attr(ATTR_RTL_FN, value.rtl ? '1' : '0'));
  }
  if (value.defaultTabSize !== undefined) {
    pPr.attrs.push(
      attr(
        qname('', 'defTabSz', ''),
        String(
          emuPositiveCoordinate32(value.defaultTabSize, 'text field paragraph default tab size'),
        ),
      ),
    );
  }

  const addPointSpacing = (name: 'spcBef' | 'spcAft', points: number | undefined): void => {
    if (points === undefined) return;
    if (!Number.isFinite(points) || points < 0) {
      throw new RangeError(
        `text field paragraph ${name} must be a non-negative number, got ${points}`,
      );
    }
    insertChildByRank(
      pPr,
      elem(qname('a', name, NS.dml), {
        children: [
          elem(qname('a', 'spcPts', NS.dml), {
            attrs: [attr(qname('', 'val', ''), String(Math.round(points * 100)))],
          }),
        ],
      }),
      pPrChildRank,
    );
  };
  if (value.lineSpacing !== undefined) {
    if (!Number.isFinite(value.lineSpacing.value) || value.lineSpacing.value < 0) {
      throw new RangeError(
        `text field paragraph line spacing must be a non-negative number, got ${value.lineSpacing.value}`,
      );
    }
    const inner =
      value.lineSpacing.kind === 'pct'
        ? elem(qname('a', 'spcPct', NS.dml), {
            attrs: [
              attr(qname('', 'val', ''), String(Math.round(value.lineSpacing.value * 100000))),
            ],
          })
        : value.lineSpacing.kind === 'pts'
          ? elem(qname('a', 'spcPts', NS.dml), {
              attrs: [
                attr(qname('', 'val', ''), String(Math.round(value.lineSpacing.value * 100))),
              ],
            })
          : null;
    if (inner === null)
      throw new RangeError('text field paragraph line spacing kind must be pct or pts.');
    insertChildByRank(pPr, elem(qname('a', 'lnSpc', NS.dml), { children: [inner] }), pPrChildRank);
  }
  addPointSpacing('spcBef', value.spcBefPts);
  addPointSpacing('spcAft', value.spcAftPts);
  if (value.tabStops !== undefined) {
    if (value.tabStops.length > 32) {
      throw new RangeError(
        `text field paragraph tab stops must contain at most 32 entries, got ${value.tabStops.length}`,
      );
    }
    const stops = value.tabStops.map((stop, index) => ({
      positionEmu: emuCoordinate32(
        stop.positionEmu,
        `text field paragraph tab stop ${index} position`,
      ),
      alignment: oneOf(
        stop.alignment,
        PARAGRAPH_TAB_ALIGNMENTS,
        `text field paragraph tab stop ${index} alignment`,
      ),
    }));
    insertChildByRank(
      pPr,
      elem(qname('a', 'tabLst', NS.dml), {
        children: stops.map((stop) =>
          elem(qname('a', 'tab', NS.dml), {
            attrs: [
              attr(qname('', 'pos', ''), String(stop.positionEmu)),
              attr(qname('', 'algn', ''), TAB_ALIGNMENT_TO_TOKEN[stop.alignment]),
            ],
          }),
        ),
      }),
      pPrChildRank,
    );
  }
  return pPr;
};

/** Reads the literal paragraph default-tab interval in EMU, or `null` when it inherits. */
export const getParagraphDefaultTabSize = (
  shape: SlideShapeData,
  paragraphIndex: number,
): number | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return null;
  const raw = getAttrValue(pPr, qname('', 'defTabSz', ''));
  if (raw === null) return null;
  const parsed = Number.parseInt(raw, 10);
  return Number.isFinite(parsed) ? parsed : null;
};

/** Authors a paragraph default-tab interval in EMU; `null` restores inheritance. */
export const setParagraphDefaultTabSize = (
  shape: SlideShapeData,
  paragraphIndex: number,
  sizeEmu: number | null,
): void => {
  const normalized =
    sizeEmu === null ? null : emuPositiveCoordinate32(sizeEmu, 'paragraph default tab size');
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);
  pPr.attrs = pPr.attrs.filter(
    (candidate) => !(candidate.name.namespaceURI === '' && candidate.name.localName === 'defTabSz'),
  );
  if (normalized !== null) {
    pPr.attrs.push(attr(qname('', 'defTabSz', ''), String(normalized)));
  }
  commitAndRefresh(shape);
};

/** Reads the literal custom paragraph tab-stop list, or `null` when no list is authored. */
export const getParagraphTabStops = (
  shape: SlideShapeData,
  paragraphIndex: number,
): readonly ParagraphTabStop[] | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return null;
  return parseParagraphTabStops(pPr);
};

/** Authors up to 32 custom paragraph tab stops; `null` removes the authored list. */
export const setParagraphTabStops = (
  shape: SlideShapeData,
  paragraphIndex: number,
  stops: readonly ParagraphTabStop[] | null,
): void => {
  if (stops !== null && stops.length > 32) {
    throw new RangeError(
      `paragraph tab stops must contain at most 32 entries, got ${stops.length}`,
    );
  }
  const normalized = stops?.map((stop, index) => ({
    positionEmu: emuCoordinate32(stop.positionEmu, `paragraph tab stop ${index} position`),
    alignment: oneOf(
      stop.alignment,
      PARAGRAPH_TAB_ALIGNMENTS,
      `paragraph tab stop ${index} alignment`,
    ),
  }));
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);
  pPr.children = pPr.children.filter(
    (candidate) =>
      !(
        candidate.kind === 'element' &&
        candidate.name.namespaceURI === NS.dml &&
        candidate.name.localName === 'tabLst'
      ),
  );
  if (normalized !== undefined) {
    const tabList = elem(qname('a', 'tabLst', NS.dml), {
      children: normalized.map((stop) =>
        elem(qname('a', 'tab', NS.dml), {
          attrs: [
            attr(qname('', 'pos', ''), String(stop.positionEmu)),
            attr(qname('', 'algn', ''), TAB_ALIGNMENT_TO_TOKEN[stop.alignment]),
          ],
        }),
      ),
    });
    insertChildByRank(pPr, tabList, pPrChildRank);
  }
  commitAndRefresh(shape);
};

/**
 * Sets the spacing before and/or after a paragraph, in points (where
 * a "point" is 1/72 inch). PowerPoint stores these as hundredths of a
 * point inside `<a:pPr><a:spcBef>/<a:spcAft><a:spcPts val="…"/>` —
 * the helper converts.
 *
 *   setParagraphSpacing(shape, 0, { beforePts: 6, afterPts: 3 });
 *
 * Omitting a side keeps the existing value (or layout default).
 * Passing a side as `null` removes that spacing element.
 */
export const setParagraphSpacing = (
  shape: SlideShapeData,
  paragraphIndex: number,
  opts: { beforePts?: number | null; afterPts?: number | null },
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);

  const writeSide = (localName: 'spcBef' | 'spcAft', value: number | null | undefined): void => {
    if (value === undefined) return;
    pPr.children = pPr.children.filter(
      (c) =>
        !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === localName),
    );
    if (value === null) return;
    if (!Number.isFinite(value) || value < 0) {
      throw new RangeError(`paragraph ${localName} must be a non-negative number, got ${value}`);
    }
    const spcEl = elem(qname('a', localName, NS.dml), {
      children: [
        elem(qname('a', 'spcPts', NS.dml), {
          attrs: [attr(qname('', 'val', ''), String(Math.round(value * 100)))],
        }),
      ],
    });
    // spcBef/spcAft must precede any bullet (buChar/buAutoNum/buFont/...) and
    // tabLst/defRPr in CT_TextParagraphProperties; insert at the schema slot
    // rather than pushing to the end of a paragraph that already has bullets.
    insertChildByRank(pPr, spcEl, pPrChildRank);
  };

  writeSide('spcBef', opts.beforePts);
  writeSide('spcAft', opts.afterPts);
  commitAndRefresh(shape);
};

/**
 * Reads back paragraph spacing in points. Returns `{ beforePts,
 * afterPts }`; each side is `null` when no `<a:spcBef>` / `<a:spcAft>`
 * is present or when the inner element isn't `<a:spcPts>` (percentage
 * spacing is reported as `null` for now).
 */
export const getParagraphSpacing = (
  shape: SlideShapeData,
  paragraphIndex: number,
): { readonly beforePts: number | null; readonly afterPts: number | null } => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return { beforePts: null, afterPts: null };
  const readSide = (localName: 'spcBef' | 'spcAft'): number | null => {
    const side = firstChildElement(pPr, qname('a', localName, NS.dml));
    if (!side) return null;
    const spcPts = firstChildElement(side, qname('a', 'spcPts', NS.dml));
    if (!spcPts) return null;
    const v = getAttrValue(spcPts, qname('', 'val', ''));
    if (v === null) return null;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n / 100 : null;
  };
  return { beforePts: readSide('spcBef'), afterPts: readSide('spcAft') };
};

/**
 * Reads the paragraph's left / right / first-line indents from
 * `<a:pPr marL="…" marR="…" indent="…"/>`. Each is in EMU (matching
 * PowerPoint's internal storage); positive means a positive indent,
 * negative `indent` is a hanging indent (typical for bullets).
 *
 * Returns `null` for sides the paragraph doesn't set (those inherit
 * from the layout / master).
 */
export const getParagraphIndent = (
  shape: SlideShapeData,
  paragraphIndex: number,
): { leftEmu: number | null; rightEmu: number | null; firstLineEmu: number | null } => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return { leftEmu: null, rightEmu: null, firstLineEmu: null };
  const read = (name: string): number | null => {
    const raw = getAttrValue(pPr, qname('', name, ''));
    if (raw === null) return null;
    const n = Number.parseInt(raw, 10);
    return Number.isFinite(n) ? n : null;
  };
  return {
    leftEmu: read('marL'),
    rightEmu: read('marR'),
    firstLineEmu: read('indent'),
  };
};

/**
 * Sets paragraph left/right margins and first-line indentation in EMU.
 * Omitted sides are preserved; `null` removes one authored attribute so it
 * inherits again. Left/right margins are non-negative while first-line
 * indentation may be negative for hanging-indent layouts.
 */
export const setParagraphIndent = (
  shape: SlideShapeData,
  paragraphIndex: number,
  indent: {
    readonly leftEmu?: number | null;
    readonly rightEmu?: number | null;
    readonly firstLineEmu?: number | null;
  },
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);
  const write = (
    name: 'marL' | 'marR' | 'indent',
    value: number | null | undefined,
    normalize: (value: number, field: string) => number,
  ): void => {
    if (value === undefined) return;
    pPr.attrs = pPr.attrs.filter(
      (candidate) => !(candidate.name.namespaceURI === '' && candidate.name.localName === name),
    );
    if (value !== null)
      pPr.attrs.push(attr(qname('', name, ''), String(normalize(value, `paragraph ${name}`))));
  };
  write('marL', indent.leftEmu, emuPositiveCoordinate32);
  write('marR', indent.rightEmu, emuPositiveCoordinate32);
  write('indent', indent.firstLineEmu, emuCoordinate32);
  commitAndRefresh(shape);
};

/**
 * Reads the paragraph's `<a:lnSpc>` line spacing. PowerPoint stores
 * line spacing two ways:
 *
 *   - Multiple of the natural line height — `<a:spcPct val="150000"/>`
 *     (= 1.5×). Returns `{ kind: 'pct', value }` with value as the unit
 *     fraction (1.5).
 *   - Fixed points — `<a:spcPts val="2400"/>` (= 24pt). Returns
 *     `{ kind: 'pts', value }` with value in points.
 *
 * Returns `null` when no `<a:lnSpc>` is present (the paragraph
 * inherits line spacing from the layout / master).
 */
export const getParagraphLineSpacing = (
  shape: SlideShapeData,
  paragraphIndex: number,
):
  | { readonly kind: 'pct'; readonly value: number }
  | { readonly kind: 'pts'; readonly value: number }
  | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return null;
  const lnSpc = firstChildElement(pPr, qname('a', 'lnSpc', NS.dml));
  if (!lnSpc) return null;
  const pct = firstChildElement(lnSpc, qname('a', 'spcPct', NS.dml));
  if (pct) {
    const v = getAttrValue(pct, qname('', 'val', ''));
    if (v !== null) {
      let n = Number.parseFloat(v);
      if (Number.isFinite(n)) {
        if (Math.abs(n) > 1) n = n / 100000;
        return { kind: 'pct', value: n };
      }
    }
  }
  const pts = firstChildElement(lnSpc, qname('a', 'spcPts', NS.dml));
  if (pts) {
    const v = getAttrValue(pts, qname('', 'val', ''));
    if (v !== null) {
      const n = Number.parseInt(v, 10);
      if (Number.isFinite(n)) return { kind: 'pts', value: n / 100 };
    }
  }
  return null;
};

/**
 * Sets a paragraph's line spacing — the writer counterpart to
 * `getParagraphLineSpacing`. Two modes (mirroring `<a:lnSpc>`):
 *
 *   - `{ kind: 'pct', value }` — a multiple of single spacing
 *     (`1` = single, `1.5` = 150%, `2` = double) → `<a:spcPct>`.
 *   - `{ kind: 'pts', value }` — a fixed leading in points → `<a:spcPts>`.
 *
 * Pass `null` to clear the override (the paragraph then inherits line
 * spacing from the layout / master).
 */
export const setParagraphLineSpacing = (
  shape: SlideShapeData,
  paragraphIndex: number,
  spacing:
    | { readonly kind: 'pct'; readonly value: number }
    | { readonly kind: 'pts'; readonly value: number }
    | null,
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);
  pPr.children = pPr.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'lnSpc'),
  );
  if (spacing !== null) {
    if (!Number.isFinite(spacing.value) || spacing.value < 0) {
      throw new RangeError(
        `line spacing value must be a non-negative number, got ${spacing.value}`,
      );
    }
    const inner =
      spacing.kind === 'pct'
        ? elem(qname('a', 'spcPct', NS.dml), {
            attrs: [attr(qname('', 'val', ''), String(Math.round(spacing.value * 100000)))],
          })
        : elem(qname('a', 'spcPts', NS.dml), {
            attrs: [attr(qname('', 'val', ''), String(Math.round(spacing.value * 100)))],
          });
    // <a:lnSpc> is the first child of CT_TextParagraphProperties.
    insertChildByRank(pPr, elem(qname('a', 'lnSpc', NS.dml), { children: [inner] }), pPrChildRank);
  }
  commitAndRefresh(shape);
};

/**
 * Reads back the bullet style on a single paragraph, or `null` when
 * no `<a:buChar>` / `<a:buAutoNum>` / `<a:buNone>` is present (the
 * paragraph inherits its bullet from the layout / master).
 */
export const getParagraphBullet = (
  shape: SlideShapeData,
  paragraphIndex: number,
): BulletStyle | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (pPr === null) return null;
  for (const c of pPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'buNone') return 'none';
    if (c.name.localName === 'buChar') {
      const char = getAttrValue(c, qname('', 'char', ''));
      if (char === '•') return 'bullet';
      if (char !== null) return { char };
    }
    if (c.name.localName === 'buAutoNum') {
      const t = getAttrValue(c, qname('', 'type', ''));
      const startAt = Number(getAttrValue(c, qname('', 'startAt', '')) ?? '1');
      if (t === 'arabicPeriod' && startAt === 1) return 'number';
      if (t !== null) return { autoNum: t, ...(startAt === 1 ? {} : { startAt }) };
    }
  }
  return null;
};

/**
 * Returns `true` when the paragraph uses an image as its bullet
 * (`<a:pPr><a:buBlip r:embed="…"/>`). Renderers without image
 * support should fall back to a generic bullet glyph.
 *
 * The underlying rId / image bytes aren't surfaced here — resolving
 * that would need the rels of the layout / master the paragraph
 * inherits from, which can be cumbersome. Knowing that the bullet
 * *is* an image is usually enough for the UI to pick a fallback.
 */
export const isParagraphBulletPicture = (
  shape: SlideShapeData,
  paragraphIndex: number,
): boolean => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return false;
  return firstChildElement(pPr, qname('a', 'buBlip', NS.dml)) !== null;
};

/**
 * Returns the embedded image bytes backing a paragraph's picture bullet
 * (`<a:pPr><a:buBlip><a:blip r:embed="…"/>`), or `null` when the
 * paragraph has no picture bullet or the relationship can't be resolved
 * (external `r:link`, missing media part, etc.).
 *
 * The `r:embed` is resolved against the *slide's* relationships, matching
 * `getShapeImageBytes` — a bullet whose blip lives only on an inherited
 * layout/master paragraph isn't followed. The returned `Uint8Array` is a
 * live view into the package media part; treat it as read-only.
 */
export const getParagraphBulletImageBytes = (
  shape: SlideShapeData,
  paragraphIndex: number,
): Uint8Array | null => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return null;
  const buBlip = firstChildElement(pPr, qname('a', 'buBlip', NS.dml));
  if (!buBlip) return null;
  const blip = firstChildElement(buBlip, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const rEmbed = getAttrValue(blip, qname('r', 'embed', NS.officeDocRels));
  if (rEmbed === null) return null;
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const rel = rels.items.find((r) => r.id === rEmbed);
  if (!rel || rel.targetMode === 'External') return null;
  const mediaName = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  const part = pkg.getPart(mediaName);
  return part?.data ?? null;
};

/** Returns the embedded bytes backing one run's direct picture underline fill. */
export const getShapeRunUnderlineFillImageBytes = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): Uint8Array | null =>
  getRunUnderlineFillImageBytes(shape[SHAPE_SLIDE], requireRun(shape, paragraphIndex, runIndex));

/** Returns the embedded bytes backing one run's direct picture text fill. */
export const getShapeRunFillImageBytes = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
): Uint8Array | null =>
  getRunDirectFillImageBytes(shape[SHAPE_SLIDE], requireRun(shape, paragraphIndex, runIndex));

/** Embeds image bytes and assigns them as one run's canonical editable `<a:blipFill>`. */
export const setShapeRunFillImage = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
  bytes: Uint8Array,
  options: { readonly format?: ImageFormat } = {},
): void => {
  const slide = shape[SHAPE_SLIDE];
  const run = requireRun(shape, paragraphIndex, runIndex);
  const oldRelationshipId = setRunDirectFillImage(slide, run, bytes, {
    ...options,
    operation: 'setShapeRunFillImage',
  });
  commitAndRefresh(shape);
  if (oldRelationshipId !== null) {
    removeUnreferencedSlideRelationships(slide, new Set([oldRelationshipId]));
  }
};

/** Embeds image bytes and assigns them as one run's editable `<a:uFill><a:blipFill>`. */
export const setShapeRunUnderlineFillImage = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
  bytes: Uint8Array,
  options: { readonly format?: ImageFormat } = {},
): void => {
  const slide = shape[SHAPE_SLIDE];
  const run = requireRun(shape, paragraphIndex, runIndex);
  const oldRelationshipId = setRunUnderlineFillImage(slide, run, bytes, {
    ...options,
    operation: 'setShapeRunUnderlineFillImage',
  });
  commitAndRefresh(shape);
  if (oldRelationshipId !== null) {
    removeUnreferencedSlideRelationships(slide, new Set([oldRelationshipId]));
  }
};

/**
 * Reads the bullet's per-paragraph color, size, and font overrides —
 * `<a:buClr>` (theme-resolved hex), `<a:buSzPct>` / `<a:buSzPts>`
 * (size relative to run or fixed pt), and `<a:buFont typeface="…"/>`.
 *
 * Returns `{ color: null, sizePct: null, sizePts: null, font: null }`
 * when the paragraph doesn't override any of them (the bullet inherits
 * from the run / layout).
 */
export const getParagraphBulletStyle = (
  pres: PresentationData,
  shape: SlideShapeData,
  paragraphIndex: number,
): {
  color: string | null;
  sizePct: number | null;
  sizePts: number | null;
  font: string | null;
} => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  if (!pPr) return { color: null, sizePct: null, sizePts: null, font: null };
  const theme = getPresentationTheme(pres);
  let color: string | null = null;
  let sizePct: number | null = null;
  let sizePts: number | null = null;
  let font: string | null = null;
  const buClr = firstChildElement(pPr, qname('a', 'buClr', NS.dml));
  if (buClr) {
    for (const c of buClr.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      color = resolveDrawingColor(c, theme);
      break;
    }
  }
  const buSzPct = firstChildElement(pPr, qname('a', 'buSzPct', NS.dml));
  if (buSzPct) {
    const v = getAttrValue(buSzPct, qname('', 'val', ''));
    if (v !== null) {
      let n = Number.parseFloat(v);
      if (Number.isFinite(n)) {
        if (Math.abs(n) > 1) n = n / 100000;
        sizePct = n;
      }
    }
  }
  const buSzPts = firstChildElement(pPr, qname('a', 'buSzPts', NS.dml));
  if (buSzPts) {
    const v = getAttrValue(buSzPts, qname('', 'val', ''));
    if (v !== null) {
      const n = Number.parseInt(v, 10);
      if (Number.isFinite(n)) sizePts = n / 100;
    }
  }
  const buFont = firstChildElement(pPr, qname('a', 'buFont', NS.dml));
  if (buFont) {
    const t = getAttrValue(buFont, qname('', 'typeface', ''));
    if (t !== null) font = t;
  }
  return { color, sizePct, sizePts, font };
};

/**
 * Sets the bullet style on a single paragraph. Same `BulletStyle` shape
 * as `setShapeBullets` — pass `'bullet'` / `'number'` / `'none'` or an
 * object like `{ char: '◆' }` / `{ autoNum: 'romanLcPeriod' }`.
 * Replacing a picture bullet also removes its old relationship and media part,
 * but only after the complete slide no longer references them.
 */
export const setParagraphBullet = (
  shape: SlideShapeData,
  paragraphIndex: number,
  style: BulletStyle,
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = firstChildElement(paragraph, NAME_A_PPR);
  const oldRelationshipIds = pPr ? paragraphBulletRelationshipIds(pPr) : new Set<string>();
  applyBulletToParagraph(paragraph, style);
  commitAndRefresh(shape);
  removeUnreferencedSlideRelationships(shape[SHAPE_SLIDE], oldRelationshipIds);
};

/**
 * Sets one paragraph's direct picture bullet from image bytes.
 *
 * Identical package media and slide relationships are reused so repeated
 * markers remain deterministic. Replacing a picture bullet removes its stale
 * relationship and media part only after the complete slide no longer
 * references them. The image format is detected from the bytes unless an
 * explicit format is supplied.
 */
export const setParagraphBulletImage = (
  shape: SlideShapeData,
  paragraphIndex: number,
  bytes: Uint8Array,
  options: { readonly format?: ImageFormat } = {},
): void => {
  const paragraph = requireParagraph(shape, paragraphIndex);
  const pPr = ensurePPr(paragraph);
  const oldRelationshipIds = paragraphBulletRelationshipIds(pPr);
  const slide = shape[SHAPE_SLIDE];
  const relationshipId = ensureSlideImageRelationship(slide, bytes, {
    ...options,
    operation: 'setParagraphBulletImage',
  });

  // Reuse the ordinary bullet path for PowerPoint's level-aware hanging indent,
  // then replace only the marker identity with the picture relationship.
  applyBulletToParagraph(paragraph, 'bullet');
  pPr.children = pPr.children.filter(
    (child) =>
      !(
        child.kind === 'element' &&
        child.name.namespaceURI === NS.dml &&
        BULLET_CHOICE_NAMES.has(child.name.localName)
      ),
  );
  const blip = elem(qname('a', 'blip', NS.dml), {
    attrs: [attr(qname('r', 'embed', NS.officeDocRels), relationshipId)],
  });
  insertChildByRank(pPr, elem(qname('a', 'buBlip', NS.dml), { children: [blip] }), pPrChildRank);
  commitAndRefresh(shape);
  removeUnreferencedSlideRelationships(slide, oldRelationshipIds);
};

const BULLET_CHOICE_NAMES = new Set([
  'buClrTx',
  'buClr',
  'buSzTx',
  'buSzPct',
  'buSzPts',
  'buFontTx',
  'buFont',
  'buNone',
  'buAutoNum',
  'buChar',
  'buBlip',
]);

/** Collect relationships owned by the direct picture-bullet choice on one paragraph. */
const paragraphBulletRelationshipIds = (pPr: XmlElement): ReadonlySet<string> => {
  const picture = firstChildElement(pPr, qname('a', 'buBlip', NS.dml));
  return picture ? hyperlinkRelationshipIds(picture) : new Set<string>();
};

/**
 * Sets the text of a single run. Existing rPr (font, size, color, ...)
 * is preserved — only the visible characters change.
 */
export const setShapeRunText = (
  shape: SlideShapeData,
  paragraphIndex: number,
  runIndex: number,
  text: string,
): void => {
  const run = requireRun(shape, paragraphIndex, runIndex);
  writeRunText(run, text);
  commitAndRefresh(shape);
};
