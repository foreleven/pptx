import { type XmlElement, NS, walkElements } from '../../internal/xml/index.ts';
import { partName, resolveTarget } from '../../internal/opc/index.ts';
import {
  INTERNAL_PACKAGE,
  SLIDE_DOCUMENT,
  SLIDE_PART_NAME,
  type SlideData,
} from '../_internal-symbols.ts';

/** Collect every slide relationship referenced by one hyperlink subtree, including an embedded sound. */
export const hyperlinkRelationshipIds = (hyperlink: XmlElement): ReadonlySet<string> => {
  const ids = new Set<string>();
  walkElements(hyperlink, (element) => {
    for (const attribute of element.attrs) {
      if (attribute.name.namespaceURI === NS.officeDocRels && attribute.value !== '') {
        ids.add(attribute.value);
      }
    }
  });
  return ids;
};

/** Remove only candidate relationships that no longer have any reference in the complete slide XML. */
export const removeUnreferencedSlideRelationships = (
  slide: SlideData,
  candidates: ReadonlySet<string>,
): void => {
  if (candidates.size === 0) return;
  const referenced = new Set<string>();
  walkElements(slide[SLIDE_DOCUMENT].root, (element) => {
    for (const attribute of element.attrs) {
      if (
        attribute.name.namespaceURI === NS.officeDocRels &&
        attribute.value !== '' &&
        candidates.has(attribute.value)
      ) {
        referenced.add(attribute.value);
      }
    }
  });
  const stale = new Set([...candidates].filter((id) => !referenced.has(id)));
  if (stale.size === 0) return;
  const pkg = slide[INTERNAL_PACKAGE];
  const slidePartName = slide[SLIDE_PART_NAME];
  const relationships = pkg.getRels(slidePartName);
  if (relationships === null) return;
  const staleMediaParts = relationships.items
    .filter((relationship) => stale.has(relationship.id) && relationship.targetMode === 'Internal')
    .map((relationship) => resolveTarget(slidePartName, relationship.target))
    .filter((target) => target.startsWith('/ppt/media/'));
  const retained = relationships.items.filter((relationship) => !stale.has(relationship.id));
  if (retained.length === relationships.items.length) return;
  relationships.items = retained;
  pkg.setRels(slidePartName, relationships);
  for (const mediaPartName of staleMediaParts) {
    const referenced = pkg.parts.some((part) => {
      if (!part.name.endsWith('.rels')) return false;
      const sourceName = part.name.replace('/_rels/', '/').replace(/\.rels$/u, '');
      if (sourceName === '' || sourceName === '/') return false;
      return (
        pkg
          .getRels(partName(sourceName))
          ?.items.some(
            (relationship) =>
              relationship.targetMode === 'Internal' &&
              resolveTarget(partName(sourceName), relationship.target) === mediaPartName,
          ) ?? false
      );
    });
    if (!referenced) pkg.removePart(partName(mediaPartName));
  }
};
