import { resolveTarget } from '../../internal/opc/index.ts';
import {
  NS,
  firstChildElement,
  getAttrValue,
  qname,
  walkElements,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_PART_NAME,
  type SlideShapeData,
} from '../_internal-symbols.ts';

export interface ShapeMedia {
  readonly kind: 'video' | 'audio';
  readonly partName: string;
  readonly contentType: string;
  readonly bytes: Uint8Array;
  readonly posterPartName: string | null;
  readonly posterContentType: string | null;
  readonly posterBytes: Uint8Array | null;
  readonly clickToPlay: boolean;
}

const NAME_NV_PIC_PR = qname('p', 'nvPicPr', NS.pml);
const NAME_C_NV_PR = qname('p', 'cNvPr', NS.pml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_VIDEO_FILE = qname('a', 'videoFile', NS.dml);
const NAME_AUDIO_FILE = qname('a', 'audioFile', NS.dml);
const NAME_HLINK_CLICK = qname('a', 'hlinkClick', NS.dml);
const NAME_P14_MEDIA = qname('p14', 'media', NS.p14);
const NAME_BLIP_FILL = qname('p', 'blipFill', NS.pml);
const NAME_BLIP = qname('a', 'blip', NS.dml);
const ATTR_R_LINK = qname('r', 'link', NS.officeDocRels);
const ATTR_R_EMBED = qname('r', 'embed', NS.officeDocRels);
const ATTR_ACTION = qname('', 'action', '');

/** Detects media markup even when its relationship is external or broken. */
export const getShapeMediaKind = (shape: SlideShapeData): 'video' | 'audio' | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const nvPicPr = firstChildElement(shape[SHAPE_ELEMENT], NAME_NV_PIC_PR);
  const nvPr = nvPicPr ? firstChildElement(nvPicPr, NAME_NV_PR) : null;
  if (!nvPr) return null;
  if (firstChildElement(nvPr, NAME_VIDEO_FILE)) return 'video';
  if (firstChildElement(nvPr, NAME_AUDIO_FILE)) return 'audio';
  return null;
};

/** Returns embedded audio/video bytes and its poster for a media-backed picture. */
export const getShapeMedia = (shape: SlideShapeData): ShapeMedia | null => {
  const kind = getShapeMediaKind(shape);
  if (kind === null) return null;
  const element = shape[SHAPE_ELEMENT];
  const nvPicPr = firstChildElement(element, NAME_NV_PIC_PR);
  const nvPr = nvPicPr ? firstChildElement(nvPicPr, NAME_NV_PR) : null;
  if (!nvPr) return null;
  const videoFile = firstChildElement(nvPr, NAME_VIDEO_FILE);
  const audioFile = firstChildElement(nvPr, NAME_AUDIO_FILE);
  const mediaFile = videoFile ?? audioFile;
  if (!mediaFile) return null;
  const fileRelId = getAttrValue(mediaFile, ATTR_R_LINK);
  let embeddedRelId: string | null = null;
  walkElements(nvPr, (candidate) => {
    if (
      candidate.name.namespaceURI === NAME_P14_MEDIA.namespaceURI &&
      candidate.name.localName === 'media'
    ) {
      embeddedRelId = getAttrValue(candidate, ATTR_R_EMBED);
      return false;
    }
    return true;
  });
  const media = relatedPart(shape, embeddedRelId ?? fileRelId);
  if (!media) return null;

  const blipFill = firstChildElement(element, NAME_BLIP_FILL);
  const blip = blipFill ? firstChildElement(blipFill, NAME_BLIP) : null;
  const poster = relatedPart(shape, blip ? getAttrValue(blip, ATTR_R_EMBED) : null);
  const cNvPr = nvPicPr ? firstChildElement(nvPicPr, NAME_C_NV_PR) : null;
  const click = cNvPr ? firstChildElement(cNvPr, NAME_HLINK_CLICK) : null;
  return {
    kind,
    partName: media.name,
    contentType: media.contentType,
    bytes: media.data,
    posterPartName: poster?.name ?? null,
    posterContentType: poster?.contentType ?? null,
    posterBytes: poster?.data ?? null,
    clickToPlay: click ? getAttrValue(click, ATTR_ACTION) === 'ppaction://media' : false,
  };
};

function relatedPart(
  shape: SlideShapeData,
  relationshipId: string | null,
): { name: string; contentType: string; data: Uint8Array } | null {
  if (!relationshipId) return null;
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const relationship = pkg
    .getRels(slide[SLIDE_PART_NAME])
    ?.items.find((item) => item.id === relationshipId);
  if (!relationship || relationship.targetMode === 'External') return null;
  const name = resolveTarget(slide[SLIDE_PART_NAME], relationship.target);
  const part = pkg.getPart(name);
  return part ? { name: part.name, contentType: part.contentType, data: part.data } : null;
}
