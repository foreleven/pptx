import { type SlideShapeData } from '../_internal-symbols.ts';
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
const NAME_A_P = qname('a', 'p', NS.dml);
const NAME_A_PPR = qname('a', 'pPr', NS.dml);
const NAME_A_R = qname('a', 'r', NS.dml);
const NAME_A_BR = qname('a', 'br', NS.dml);
const NAME_A_FLD = qname('a', 'fld', NS.dml);
const NAME_A_RPR = qname('a', 'rPr', NS.dml);

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
}

/** Exact known text node that owns one extension payload. */
export type ShapeTextExtensionTarget =
  | { readonly kind: 'bodyProperties' }
  | { readonly kind: 'paragraphProperties'; readonly paragraphIndex: number }
  | {
      readonly kind: 'runProperties';
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

const extensionContent = (element: XmlElement): ShapeTextExtensionContent | null => {
  const attributes = extensionAttributes(element);
  const children = element.children.filter(isExtensionChild);
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
    ...(children.length === 0 ? {} : { children: children.map((child) => publicNode(child)) }),
  };
};

const appendPayload = (
  payloads: ShapeTextExtensionPayload[],
  target: ShapeTextExtensionTarget,
  element: XmlElement | null,
): void => {
  if (element === null) return;
  const content = extensionContent(element);
  if (content !== null) payloads.push({ target, content });
};

/**
 * Reads foreign attributes, local namespace declarations, and extension children from known
 * DrawingML text nodes. The carrier is deliberately index-addressed and excludes ordinary
 * DrawingML attributes/children so it cannot become a parallel raw-XML authoring API.
 */
export const getShapeTextExtensionPayloads = (
  shape: SlideShapeData,
): readonly ShapeTextExtensionPayload[] => {
  const textBody = requireTxBody(shape);
  const payloads: ShapeTextExtensionPayload[] = [];
  appendPayload(payloads, { kind: 'bodyProperties' }, firstChildElement(textBody, NAME_A_BODY_PR));
  paragraphsOf(textBody).forEach((paragraph, paragraphIndex) => {
    appendPayload(
      payloads,
      { kind: 'paragraphProperties', paragraphIndex },
      firstChildElement(paragraph, NAME_A_PPR),
    );
    inlineElementsOf(paragraph).forEach((inline, elementIndex) => {
      if (inline.name.localName === NAME_A_R.localName) {
        appendPayload(
          payloads,
          { kind: 'runProperties', paragraphIndex, elementIndex },
          firstChildElement(inline, NAME_A_RPR),
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
  return {
    ...content,
    prefixDecls,
    privateAttributes,
    privateChildren,
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
  if (target.kind === 'bodyProperties') return ensureFirstChild(textBody, NAME_A_BODY_PR);
  const paragraph = requireParagraph(textBody, target.paragraphIndex);
  if (target.kind === 'paragraphProperties') return ensureFirstChild(paragraph, NAME_A_PPR);
  integerIndex(target.elementIndex, 'text extension elementIndex');
  const inline = inlineElementsOf(paragraph)[target.elementIndex];
  if (!inline)
    throw new RangeError(`text extension elementIndex ${target.elementIndex} is out of range.`);
  if (target.kind === 'runProperties') return ensureFirstChild(inline, NAME_A_RPR);
  if (inline.name.localName !== NAME_A_FLD.localName) {
    throw new TypeError(`text extension target ${target.kind} points to ${inline.name.localName}.`);
  }
  if (target.kind === 'field') return inline;
  return ensureFirstChild(inline, NAME_A_PPR);
};

const replaceExtensionContent = (
  element: XmlElement,
  content: ReturnType<typeof privateContent>,
): void => {
  const current = extensionContent(element);
  for (const namespace of current?.namespaces ?? []) element.prefixDecls.delete(namespace.prefix);
  for (const [prefix, uri] of content.prefixDecls) element.prefixDecls.set(prefix, uri);
  element.attrs = [
    ...element.attrs.filter((attribute) => attribute.name.namespaceURI === ''),
    ...content.privateAttributes,
  ];
  element.children = [
    ...element.children.filter((child) => !isExtensionChild(child)),
    ...content.privateChildren,
  ];
};

/**
 * Replaces the restricted extension payload at exact text nodes. All descriptors are validated
 * and resolved before mutation, so a malformed path or XML descriptor cannot partially edit a shape.
 */
export const setShapeTextExtensionPayloads = (
  shape: SlideShapeData,
  payloads: readonly ShapeTextExtensionPayload[],
): void => {
  const textBody = requireTxBody(shape);
  const keys = new Set<string>();
  const prepared = payloads.map((payload, index) => {
    const key = targetKey(payload.target);
    if (keys.has(key)) throw new TypeError(`duplicate text extension target ${key}.`);
    keys.add(key);
    return {
      target: resolveTarget(textBody, payload.target),
      content: privateContent(payload.content, `text extension payloads[${index}].content`),
    };
  });
  for (const entry of prepared) replaceExtensionContent(entry.target, entry.content);
  commitAndRefresh(shape);
};
