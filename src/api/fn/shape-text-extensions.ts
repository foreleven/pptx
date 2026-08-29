import {
  CELL_ELEMENT,
  CELL_TABLE,
  type SlideShapeData,
  type TableCellData,
} from '../_internal-symbols.ts';
import { commitAndRefresh, requireTxBody } from './_helpers.ts';
import {
  NS,
  type QName,
  type XmlAttr,
  type XmlElement,
  type XmlNode,
  attr,
  cdata,
  comment,
  elem,
  firstChildElement,
  pi,
  qname,
  text,
} from '../../internal/xml/index.ts';

const NAME_A_BODY_PR = qname('a', 'bodyPr', NS.dml);
const NAME_A_TX_BODY = qname('a', 'txBody', NS.dml);
const NAME_A_P = qname('a', 'p', NS.dml);
const NAME_A_PPR = qname('a', 'pPr', NS.dml);
const NAME_A_R = qname('a', 'r', NS.dml);
const NAME_A_BR = qname('a', 'br', NS.dml);
const NAME_A_FLD = qname('a', 'fld', NS.dml);
const NAME_A_RPR = qname('a', 'rPr', NS.dml);
const NAME_A_U_LN = qname('a', 'uLn', NS.dml);
const NAME_A_LN = qname('a', 'ln', NS.dml);

/** Namespace declaration retained on one known DrawingML text element. */
export interface ShapeTextExtensionNamespace {
  readonly prefix: string;
  readonly uri: string;
}

/** Namespace-aware XML name used by the restricted text-extension carrier. */
export interface ShapeTextExtensionName {
  readonly prefix: string;
  readonly localName: string;
  readonly namespaceUri: string;
}

/** Foreign attribute retained on one known DrawingML text element. */
export interface ShapeTextExtensionAttribute {
  readonly name: ShapeTextExtensionName;
  readonly value: string;
}

/** Structured XML payload accepted inside an extension child. */
export type ShapeTextExtensionNode =
  | {
      readonly kind: 'element';
      readonly name: ShapeTextExtensionName;
      readonly namespaces?: readonly ShapeTextExtensionNamespace[];
      readonly attributes?: readonly ShapeTextExtensionAttribute[];
      readonly children?: readonly ShapeTextExtensionNode[];
    }
  | { readonly kind: 'text'; readonly value: string }
  | { readonly kind: 'cdata'; readonly value: string }
  | { readonly kind: 'comment'; readonly value: string }
  | { readonly kind: 'processingInstruction'; readonly target: string; readonly value: string };

/** Unknown attributes and extension children retained at one exact text node. */
export interface ShapeTextExtensionContent {
  readonly namespaces?: readonly ShapeTextExtensionNamespace[];
  readonly attributes?: readonly ShapeTextExtensionAttribute[];
  readonly children?: readonly ShapeTextExtensionNode[];
  /** Zero-based XML child slots retained relative to known children; omitted legacy payloads append. */
  readonly childPositions?: readonly number[];
}

/** Exact known text node that owns one extension payload. */
export type ShapeTextExtensionTarget =
  | { readonly kind: 'textBody' }
  | { readonly kind: 'bodyProperties' }
  | { readonly kind: 'paragraph'; readonly paragraphIndex: number }
  | { readonly kind: 'paragraphProperties'; readonly paragraphIndex: number }
  | {
      readonly kind: 'run';
      readonly paragraphIndex: number;
      readonly elementIndex: number;
    }
  | {
      readonly kind: 'runProperties';
      readonly paragraphIndex: number;
      readonly elementIndex: number;
    }
  | {
      readonly kind: 'runUnderlineLineProperties';
      readonly paragraphIndex: number;
      readonly elementIndex: number;
    }
  | {
      readonly kind: 'runOutlineProperties';
      readonly paragraphIndex: number;
      readonly elementIndex: number;
    }
  | {
      readonly kind: 'field';
      readonly paragraphIndex: number;
      readonly elementIndex: number;
    }
  | {
      readonly kind: 'fieldParagraphProperties';
      readonly paragraphIndex: number;
      readonly elementIndex: number;
    };

/** One deterministic, index-addressed extension payload for a shape text body. */
export interface ShapeTextExtensionPayload {
  readonly target: ShapeTextExtensionTarget;
  readonly content: ShapeTextExtensionContent;
}

const paragraphsOf = (textBody: XmlElement): XmlElement[] =>
  textBody.children.filter(
    (child): child is XmlElement =>
      child.kind === 'element' &&
      child.name.namespaceURI === NAME_A_P.namespaceURI &&
      child.name.localName === NAME_A_P.localName,
  );

const inlineElementsOf = (paragraph: XmlElement): XmlElement[] =>
  paragraph.children.filter(
    (child): child is XmlElement =>
      child.kind === 'element' &&
      child.name.namespaceURI === NS.dml &&
      [NAME_A_R.localName, NAME_A_BR.localName, NAME_A_FLD.localName].includes(
        child.name.localName,
      ),
  );

const isExtensionChild = (node: XmlNode): node is XmlElement =>
  node.kind === 'element' &&
  (node.name.namespaceURI !== NS.dml || node.name.localName === 'extLst');

const extensionAttributes = (element: XmlElement): XmlAttr[] =>
  element.attrs.filter((attribute) => attribute.name.namespaceURI !== '');

const ROOT_TEXT_NAMESPACE_SCOPE = new Map([
  ['a', NS.dml],
  ['p', NS.pml],
  ['r', NS.officeDocRels],
  ['xml', NS.xml],
]);

/** Hoist only bindings inherited from outside the retained text node. */
const extensionNamespaces = (
  element: XmlElement,
  attributes: readonly XmlAttr[],
  children: readonly XmlElement[],
): readonly ShapeTextExtensionNamespace[] => {
  const declarations = new Map(element.prefixDecls);
  const rootScope = new Map([...ROOT_TEXT_NAMESPACE_SCOPE, ...declarations]);
  const requireName = (name: QName, scope: Map<string, string>): void => {
    if (name.namespaceURI === '' || scope.get(name.prefix) === name.namespaceURI) return;
    declarations.set(name.prefix, name.namespaceURI);
    rootScope.set(name.prefix, name.namespaceURI);
    scope.set(name.prefix, name.namespaceURI);
  };
  for (const attribute of attributes) requireName(attribute.name, rootScope);
  const visit = (node: XmlNode, parentScope: Map<string, string>): void => {
    if (node.kind !== 'element') return;
    const scope = new Map([...parentScope, ...node.prefixDecls]);
    requireName(node.name, scope);
    for (const attribute of node.attrs) requireName(attribute.name, scope);
    for (const child of node.children) visit(child, scope);
  };
  for (const child of children) visit(child, rootScope);
  return [...declarations].map(([prefix, uri]) => ({ prefix, uri }));
};

const publicName = (name: QName): ShapeTextExtensionName => ({
  prefix: name.prefix,
  localName: name.localName,
  namespaceUri: name.namespaceURI,
});

const publicNode = (node: XmlNode): ShapeTextExtensionNode => {
  switch (node.kind) {
    case 'element':
      return {
        kind: 'element',
        name: publicName(node.name),
        ...(node.prefixDecls.size === 0
          ? {}
          : {
              namespaces: [...node.prefixDecls].map(([prefix, uri]) => ({ prefix, uri })),
            }),
        ...(node.attrs.length === 0
          ? {}
          : {
              attributes: node.attrs.map((attribute) => ({
                name: publicName(attribute.name),
                value: attribute.value,
              })),
            }),
        ...(node.children.length === 0
          ? {}
          : { children: node.children.map((child) => publicNode(child)) }),
      };
    case 'text':
    case 'cdata':
    case 'comment':
      return { kind: node.kind, value: node.data };
    case 'pi':
      return { kind: 'processingInstruction', target: node.target, value: node.data };
  }
};

const extensionContent = (
  element: XmlElement,
  retainChildPositions = false,
): ShapeTextExtensionContent | null => {
  const attributes = extensionAttributes(element);
  const childEntries = element.children
    .map((child, position) => ({ child, position }))
    .filter((entry): entry is { child: XmlElement; position: number } =>
      isExtensionChild(entry.child),
    );
  const children = childEntries.map((entry) => entry.child);
  const namespaces = extensionNamespaces(element, attributes, children);
  if (namespaces.length === 0 && attributes.length === 0 && children.length === 0) return null;
  return {
    ...(namespaces.length === 0 ? {} : { namespaces }),
    ...(attributes.length === 0
      ? {}
      : {
          attributes: attributes.map((attribute) => ({
            name: publicName(attribute.name),
            value: attribute.value,
          })),
        }),
    ...(children.length === 0
      ? {}
      : {
          children: children.map((child) => publicNode(child)),
          ...(retainChildPositions
            ? { childPositions: childEntries.map((entry) => entry.position) }
            : {}),
        }),
  };
};

const appendPayload = (
  payloads: ShapeTextExtensionPayload[],
  target: ShapeTextExtensionTarget,
  element: XmlElement | null,
): void => {
  if (element === null) return;
  const content = extensionContent(
    element,
    target.kind === 'textBody' || target.kind === 'paragraph' || target.kind === 'run',
  );
  if (content !== null) payloads.push({ target, content });
};

/**
 * Reads foreign attributes, local namespace declarations, and extension children from known
 * DrawingML text nodes. The carrier is deliberately index-addressed and excludes ordinary
 * DrawingML attributes/children so it cannot become a parallel raw-XML authoring API.
 */
const textExtensionPayloads = (textBody: XmlElement): readonly ShapeTextExtensionPayload[] => {
  const payloads: ShapeTextExtensionPayload[] = [];
  appendPayload(payloads, { kind: 'textBody' }, textBody);
  appendPayload(payloads, { kind: 'bodyProperties' }, firstChildElement(textBody, NAME_A_BODY_PR));
  paragraphsOf(textBody).forEach((paragraph, paragraphIndex) => {
    appendPayload(payloads, { kind: 'paragraph', paragraphIndex }, paragraph);
    appendPayload(
      payloads,
      { kind: 'paragraphProperties', paragraphIndex },
      firstChildElement(paragraph, NAME_A_PPR),
    );
    inlineElementsOf(paragraph).forEach((inline, elementIndex) => {
      if (inline.name.localName === NAME_A_R.localName) {
        appendPayload(payloads, { kind: 'run', paragraphIndex, elementIndex }, inline);
        const runProperties = firstChildElement(inline, NAME_A_RPR);
        appendPayload(
          payloads,
          { kind: 'runProperties', paragraphIndex, elementIndex },
          runProperties,
        );
        appendPayload(
          payloads,
          { kind: 'runUnderlineLineProperties', paragraphIndex, elementIndex },
          runProperties === null ? null : firstChildElement(runProperties, NAME_A_U_LN),
        );
        appendPayload(
          payloads,
          { kind: 'runOutlineProperties', paragraphIndex, elementIndex },
          runProperties === null ? null : firstChildElement(runProperties, NAME_A_LN),
        );
      } else if (inline.name.localName === NAME_A_FLD.localName) {
        appendPayload(payloads, { kind: 'field', paragraphIndex, elementIndex }, inline);
        appendPayload(
          payloads,
          { kind: 'runProperties', paragraphIndex, elementIndex },
          firstChildElement(inline, NAME_A_RPR),
        );
        appendPayload(
          payloads,
          { kind: 'fieldParagraphProperties', paragraphIndex, elementIndex },
          firstChildElement(inline, NAME_A_PPR),
        );
      }
    });
  });
  return payloads;
};

export const getShapeTextExtensionPayloads = (
  shape: SlideShapeData,
): readonly ShapeTextExtensionPayload[] => textExtensionPayloads(requireTxBody(shape));

/** Reads the same restricted text-extension carrier from one table cell. */
export const getTableCellTextExtensionPayloads = (
  cell: TableCellData,
): readonly ShapeTextExtensionPayload[] => {
  const textBody = firstChildElement(cell[CELL_ELEMENT], NAME_A_TX_BODY);
  return textBody === null ? [] : textExtensionPayloads(textBody);
};

const integerIndex = (value: number, label: string): void => {
  if (!Number.isInteger(value) || value < 0) {
    throw new RangeError(`${label} must be a non-negative integer.`);
  }
};

const validateName = (name: ShapeTextExtensionName, label: string, attribute: boolean): QName => {
  if (
    typeof name.prefix !== 'string' ||
    typeof name.localName !== 'string' ||
    typeof name.namespaceUri !== 'string' ||
    !/^[A-Za-z_][A-Za-z0-9._-]*$/u.test(name.localName) ||
    (name.prefix !== '' && !/^[A-Za-z_][A-Za-z0-9._-]*$/u.test(name.prefix))
  ) {
    throw new TypeError(`${label} must contain a valid XML prefix, localName, and namespaceUri.`);
  }
  if (attribute && name.namespaceUri === '') {
    throw new TypeError(
      `${label} must be namespace-qualified; ordinary DrawingML attributes are not extensions.`,
    );
  }
  if (attribute && name.prefix === '') {
    throw new TypeError(
      `${label} must use a prefix because default namespaces do not apply to XML attributes.`,
    );
  }
  if (name.prefix === 'xmlns' || name.namespaceUri === NS.xmlns) {
    throw new TypeError(`${label} cannot author xmlns as an ordinary attribute.`);
  }
  if (name.prefix !== '' && name.namespaceUri === '') {
    throw new TypeError(`${label} cannot use an unbound XML prefix.`);
  }
  if ((name.prefix === 'xml') !== (name.namespaceUri === NS.xml)) {
    throw new TypeError(`${label} must pair the xml prefix with its reserved namespace URI.`);
  }
  return qname(name.prefix, name.localName, name.namespaceUri);
};

const namespaceMap = (
  namespaces: readonly ShapeTextExtensionNamespace[] | undefined,
  label: string,
): Map<string, string> => {
  const result = new Map<string, string>();
  for (const namespace of namespaces ?? []) {
    if (
      typeof namespace.prefix !== 'string' ||
      typeof namespace.uri !== 'string' ||
      (namespace.prefix !== '' && !/^[A-Za-z_][A-Za-z0-9._-]*$/u.test(namespace.prefix)) ||
      namespace.prefix === 'xmlns' ||
      namespace.uri.length === 0 ||
      namespace.uri === NS.xmlns ||
      (namespace.prefix === 'xml') !== (namespace.uri === NS.xml)
    ) {
      throw new TypeError(`${label} contains an invalid namespace declaration.`);
    }
    if (result.has(namespace.prefix)) {
      throw new TypeError(`${label} contains duplicate namespace prefix ${namespace.prefix}.`);
    }
    result.set(namespace.prefix, namespace.uri);
  }
  return result;
};

const assertBoundName = (name: QName, scope: ReadonlyMap<string, string>, label: string): void => {
  if (name.namespaceURI === '' || name.namespaceURI === NS.xml) return;
  if (scope.get(name.prefix) !== name.namespaceURI) {
    throw new TypeError(
      `${label} uses prefix ${name.prefix || '(default)'} without a matching namespace declaration.`,
    );
  }
};

const privateNode = (
  node: ShapeTextExtensionNode,
  label: string,
  parentScope: ReadonlyMap<string, string>,
  depth = 0,
): XmlNode => {
  if (depth > 64) throw new RangeError(`${label} exceeds the maximum XML nesting depth of 64.`);
  if (node.kind !== 'element' && typeof node.value !== 'string') {
    throw new TypeError(`${label}.value must be a string.`);
  }
  if (node.kind === 'text') return text(node.value);
  if (node.kind === 'cdata') return cdata(node.value);
  if (node.kind === 'comment') {
    if (node.value.includes('--') || node.value.endsWith('-')) {
      throw new TypeError(`${label}.value is not valid XML comment content.`);
    }
    return comment(node.value);
  }
  if (node.kind === 'processingInstruction') {
    if (
      !/^[A-Za-z_][A-Za-z0-9._-]*$/u.test(node.target) ||
      node.target.toLowerCase() === 'xml' ||
      node.value.includes('?>')
    ) {
      throw new TypeError(`${label} is not a valid XML processing instruction.`);
    }
    return pi(node.target, node.value);
  }
  const prefixDecls = namespaceMap(node.namespaces, `${label}.namespaces`);
  const scope = new Map([...parentScope, ...prefixDecls]);
  const name = validateName(node.name, `${label}.name`, false);
  assertBoundName(name, scope, `${label}.name`);
  const attrs = (node.attributes ?? []).map((attribute, index) => {
    if (typeof attribute.value !== 'string') {
      throw new TypeError(`${label}.attributes[${index}].value must be a string.`);
    }
    const attributeName = validateName(attribute.name, `${label}.attributes[${index}].name`, false);
    assertBoundName(attributeName, scope, `${label}.attributes[${index}].name`);
    return attr(attributeName, attribute.value);
  });
  return elem(name, {
    prefixDecls,
    attrs,
    children: (node.children ?? []).map((child, index) =>
      privateNode(child, `${label}.children[${index}]`, scope, depth + 1),
    ),
  });
};

const privateContent = (
  content: ShapeTextExtensionContent,
  label: string,
): ShapeTextExtensionContent & {
  readonly prefixDecls: Map<string, string>;
  readonly privateAttributes: XmlAttr[];
  readonly privateChildren: XmlNode[];
  readonly privateChildPositions: readonly number[] | null;
} => {
  const prefixDecls = namespaceMap(content.namespaces, `${label}.namespaces`);
  const scope = new Map([...ROOT_TEXT_NAMESPACE_SCOPE, ...prefixDecls]);
  const privateAttributes = (content.attributes ?? []).map((attribute, index) => {
    if (typeof attribute.value !== 'string') {
      throw new TypeError(`${label}.attributes[${index}].value must be a string.`);
    }
    const attributeName = validateName(attribute.name, `${label}.attributes[${index}].name`, true);
    assertBoundName(attributeName, scope, `${label}.attributes[${index}].name`);
    return attr(attributeName, attribute.value);
  });
  const privateChildren = (content.children ?? []).map((child, index) =>
    privateNode(child, `${label}.children[${index}]`, scope),
  );
  for (const [index, child] of privateChildren.entries()) {
    if (!isExtensionChild(child)) {
      throw new TypeError(
        `${label}.children[${index}] must be a foreign-namespace element or DrawingML extLst.`,
      );
    }
  }
  const privateChildPositions = content.childPositions ?? null;
  if (privateChildPositions !== null) {
    if (
      !Array.isArray(privateChildPositions) ||
      privateChildPositions.length !== privateChildren.length
    ) {
      throw new TypeError(`${label}.childPositions must match the number of extension children.`);
    }
    let previous = -1;
    for (const [index, position] of privateChildPositions.entries()) {
      if (!Number.isInteger(position) || position < 0 || position <= previous) {
        throw new RangeError(
          `${label}.childPositions[${index}] must be a strictly increasing non-negative integer.`,
        );
      }
      previous = position;
    }
  }
  return {
    ...content,
    prefixDecls,
    privateAttributes,
    privateChildren,
    privateChildPositions,
  };
};

const targetKey = (target: ShapeTextExtensionTarget): string =>
  'paragraphIndex' in target
    ? `${target.kind}:${target.paragraphIndex}:${'elementIndex' in target ? target.elementIndex : ''}`
    : target.kind;

const requireParagraph = (textBody: XmlElement, paragraphIndex: number): XmlElement => {
  integerIndex(paragraphIndex, 'text extension paragraphIndex');
  const paragraph = paragraphsOf(textBody)[paragraphIndex];
  if (!paragraph)
    throw new RangeError(`text extension paragraphIndex ${paragraphIndex} is out of range.`);
  return paragraph;
};

const ensureFirstChild = (parent: XmlElement, name: QName): XmlElement => {
  const existing = firstChildElement(parent, name);
  if (existing !== null) return existing;
  const created = elem(name);
  parent.children.unshift(created);
  return created;
};

const resolveTarget = (textBody: XmlElement, target: ShapeTextExtensionTarget): XmlElement => {
  if (target.kind === 'textBody') return textBody;
  if (target.kind === 'bodyProperties') return ensureFirstChild(textBody, NAME_A_BODY_PR);
  const paragraph = requireParagraph(textBody, target.paragraphIndex);
  if (target.kind === 'paragraph') return paragraph;
  if (target.kind === 'paragraphProperties') return ensureFirstChild(paragraph, NAME_A_PPR);
  integerIndex(target.elementIndex, 'text extension elementIndex');
  const inline = inlineElementsOf(paragraph)[target.elementIndex];
  if (!inline)
    throw new RangeError(`text extension elementIndex ${target.elementIndex} is out of range.`);
  if (target.kind === 'run') {
    if (inline.name.localName !== NAME_A_R.localName) {
      throw new TypeError(`text extension target run points to ${inline.name.localName}.`);
    }
    return inline;
  }
  if (target.kind === 'runProperties') return ensureFirstChild(inline, NAME_A_RPR);
  if (target.kind === 'runUnderlineLineProperties' || target.kind === 'runOutlineProperties') {
    const runProperties = ensureFirstChild(inline, NAME_A_RPR);
    return ensureFirstChild(
      runProperties,
      target.kind === 'runUnderlineLineProperties' ? NAME_A_U_LN : NAME_A_LN,
    );
  }
  if (inline.name.localName !== NAME_A_FLD.localName) {
    throw new TypeError(`text extension target ${target.kind} points to ${inline.name.localName}.`);
  }
  if (target.kind === 'field') return inline;
  return ensureFirstChild(inline, NAME_A_PPR);
};

const replaceExtensionContent = (
  element: XmlElement,
  content: ReturnType<typeof privateContent>,
  replacementChildren: readonly XmlNode[],
): void => {
  const current = extensionContent(element);
  for (const namespace of current?.namespaces ?? []) element.prefixDecls.delete(namespace.prefix);
  for (const [prefix, uri] of content.prefixDecls) element.prefixDecls.set(prefix, uri);
  element.attrs = [
    ...element.attrs.filter((attribute) => attribute.name.namespaceURI === ''),
    ...content.privateAttributes,
  ];
  element.children = [...replacementChildren];
};

const replacementChildren = (
  element: XmlElement,
  content: ReturnType<typeof privateContent>,
  label: string,
): readonly XmlNode[] => {
  const knownChildren: XmlNode[] = element.children.filter((child) => !isExtensionChild(child));
  if (content.privateChildPositions === null) {
    return [...knownChildren, ...content.privateChildren];
  }
  const result = [...knownChildren];
  content.privateChildren.forEach((child, index) => {
    const position = content.privateChildPositions?.[index];
    if (position === undefined || position > result.length) {
      throw new RangeError(
        `${label}.childPositions[${index}] is out of range for the target's known children.`,
      );
    }
    result.splice(position, 0, child);
  });
  return result;
};

const cloneXmlNode = (node: XmlNode): XmlNode => {
  if (node.kind === 'element') {
    return elem(node.name, {
      prefixDecls: new Map(node.prefixDecls),
      attrs: node.attrs.map((attribute) => attr(attribute.name, attribute.value)),
      children: node.children.map(cloneXmlNode),
    });
  }
  if (node.kind === 'text') return text(node.data);
  if (node.kind === 'cdata') return cdata(node.data);
  if (node.kind === 'comment') return comment(node.data);
  return pi(node.target, node.data);
};

const setTextExtensionPayloads = (
  textBody: XmlElement,
  payloads: readonly ShapeTextExtensionPayload[],
): void => {
  const draft = cloneXmlNode(textBody);
  if (draft.kind !== 'element') throw new Error('text body clone must remain an XML element');
  const keys = new Set<string>();
  const prepared = payloads.map((payload, index) => {
    const key = targetKey(payload.target);
    if (keys.has(key)) throw new TypeError(`duplicate text extension target ${key}.`);
    keys.add(key);
    const target = resolveTarget(draft, payload.target);
    const label = `text extension payloads[${index}].content`;
    const content = privateContent(payload.content, label);
    return {
      target,
      content,
      replacementChildren: replacementChildren(target, content, label),
    };
  });
  for (const entry of prepared) {
    replaceExtensionContent(entry.target, entry.content, entry.replacementChildren);
  }
  textBody.prefixDecls = draft.prefixDecls;
  textBody.attrs = draft.attrs;
  textBody.children = draft.children;
};

/**
 * Replaces the restricted extension payload at exact text nodes. All descriptors are validated
 * and resolved before mutation, so a malformed path or XML descriptor cannot partially edit a shape.
 */
export const setShapeTextExtensionPayloads = (
  shape: SlideShapeData,
  payloads: readonly ShapeTextExtensionPayload[],
): void => {
  setTextExtensionPayloads(requireTxBody(shape), payloads);
  commitAndRefresh(shape);
};

/** Replaces restricted extension payloads at exact nodes inside one table-cell text body. */
export const setTableCellTextExtensionPayloads = (
  cell: TableCellData,
  payloads: readonly ShapeTextExtensionPayload[],
): void => {
  const textBody = firstChildElement(cell[CELL_ELEMENT], NAME_A_TX_BODY);
  if (textBody === null) throw new Error('table cell has no <a:txBody>');
  setTextExtensionPayloads(textBody, payloads);
  commitAndRefresh(cell[CELL_TABLE]);
};
