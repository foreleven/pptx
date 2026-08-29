import { type ImageFormat, partName, resolveTarget } from '../../internal/opc/index.ts';
import { applyRunFormat } from '../../internal/drawingml/index.ts';
import {
  NS,
  type XmlElement,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import { INTERNAL_PACKAGE, SLIDE_PART_NAME, type SlideData } from '../_internal-symbols.ts';
import { ensureSlideImageRelationship } from './_helpers.ts';

const getRunFillImageBytes = (
  slide: SlideData,
  run: XmlElement,
  container: 'direct' | 'underline',
): Uint8Array | null => {
  const rPr = firstChildElement(run, qname('a', 'rPr', NS.dml));
  const fillHost =
    container === 'underline' && rPr ? firstChildElement(rPr, qname('a', 'uFill', NS.dml)) : rPr;
  const blipFill = fillHost ? firstChildElement(fillHost, qname('a', 'blipFill', NS.dml)) : null;
  const blip = blipFill ? firstChildElement(blipFill, qname('a', 'blip', NS.dml)) : null;
  const relationshipId = blip ? getAttrValue(blip, qname('r', 'embed', NS.officeDocRels)) : null;
  if (relationshipId === null) return null;
  const pkg = slide[INTERNAL_PACKAGE];
  const relationship = pkg
    .getRels(slide[SLIDE_PART_NAME])
    ?.items.find((candidate) => candidate.id === relationshipId);
  if (!relationship || relationship.targetMode === 'External') return null;
  const mediaName = relationship.target.startsWith('/')
    ? partName(relationship.target)
    : resolveTarget(slide[SLIDE_PART_NAME], relationship.target);
  return pkg.getPart(mediaName)?.data ?? null;
};

/** Resolve one run's direct picture text fill against its owning slide. */
export const getRunDirectFillImageBytes = (slide: SlideData, run: XmlElement): Uint8Array | null =>
  getRunFillImageBytes(slide, run, 'direct');

/** Resolve one run's direct picture underline fill against its owning slide. */
export const getRunUnderlineFillImageBytes = (
  slide: SlideData,
  run: XmlElement,
): Uint8Array | null => getRunFillImageBytes(slide, run, 'underline');

/**
 * Embed bytes and assign a canonical direct picture fill to one run.
 * The caller owns committing its shape/cell and pruning the returned old relationship.
 */
export const setRunDirectFillImage = (
  slide: SlideData,
  run: XmlElement,
  bytes: Uint8Array,
  options: { readonly format?: ImageFormat; readonly operation: string },
): string | null => {
  const relationshipId = ensureSlideImageRelationship(slide, bytes, options);
  const existingRPr = firstChildElement(run, qname('a', 'rPr', NS.dml));
  const rPr = existingRPr ?? elem(qname('a', 'rPr', NS.dml));
  if (existingRPr === null) run.children.unshift(rPr);
  const oldBlipFill = firstChildElement(rPr, qname('a', 'blipFill', NS.dml));
  const oldBlip = oldBlipFill ? firstChildElement(oldBlipFill, qname('a', 'blip', NS.dml)) : null;
  const oldRelationshipId = oldBlip
    ? getAttrValue(oldBlip, qname('r', 'embed', NS.officeDocRels))
    : null;
  applyRunFormat(rPr, { pictureFill: { relationshipId } });
  return oldRelationshipId;
};

/**
 * Embed bytes and assign a direct picture underline fill to one run.
 * The caller owns committing its shape/cell and pruning the returned old relationship.
 */
export const setRunUnderlineFillImage = (
  slide: SlideData,
  run: XmlElement,
  bytes: Uint8Array,
  options: { readonly format?: ImageFormat; readonly operation: string },
): string | null => {
  const relationshipId = ensureSlideImageRelationship(slide, bytes, options);
  const existingRPr = firstChildElement(run, qname('a', 'rPr', NS.dml));
  const rPr = existingRPr ?? elem(qname('a', 'rPr', NS.dml));
  if (existingRPr === null) {
    run.children.unshift(rPr);
  }
  const oldRelationshipId = (() => {
    const uFill = firstChildElement(rPr, qname('a', 'uFill', NS.dml));
    const blipFill = uFill ? firstChildElement(uFill, qname('a', 'blipFill', NS.dml)) : null;
    const blip = blipFill ? firstChildElement(blipFill, qname('a', 'blip', NS.dml)) : null;
    return blip ? getAttrValue(blip, qname('r', 'embed', NS.officeDocRels)) : null;
  })();
  applyRunFormat(rPr, {
    underlineFill: { kind: 'picture', relationshipId },
  });
  return oldRelationshipId;
};
