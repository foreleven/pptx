import { type XmlElement, NS, walkElements } from '../../internal/xml/index.ts';
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
  const partName = slide[SLIDE_PART_NAME];
  const relationships = pkg.getRels(partName);
  if (relationships === null) return;
  const retained = relationships.items.filter((relationship) => !stale.has(relationship.id));
  if (retained.length === relationships.items.length) return;
  relationships.items = retained;
  pkg.setRels(partName, relationships);
};
