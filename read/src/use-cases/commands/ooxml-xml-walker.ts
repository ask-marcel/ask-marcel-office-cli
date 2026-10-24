import { XMLParser } from 'fast-xml-parser';

/**
 * Tiny generic walkers over a fast-xml-parser tree. OOXML is namespace-
 * heavy (`w:`, `cp:`, `dc:`, `dcterms:`, `vt:`, `w15:`, `x:`, `p:`, ...) and
 * the only traversal primitives we need are: find-every-element-by-tag-name,
 * read a single attribute, read the leaf text content. Keeping these here
 * lets the per-format metadata modules focus on what to extract, not how to
 * traverse.
 *
 * fast-xml-parser tree shape:
 *   - element names are object keys (with `w:` etc. prefix preserved)
 *   - attributes are keys prefixed with `@_` (e.g. `@_w:val`, `@_TargetMode`)
 *   - text content sits under `#text` when the element also has attributes,
 *     OR directly as a string when the element has only text and no attrs
 *   - repeated same-named children come as an array; single occurrences as
 *     a bare object — every walker must handle both shapes
 */
type XmlObject = Record<string, unknown>;

const isObject = (node: unknown): node is XmlObject => node !== null && typeof node === 'object' && !Array.isArray(node);

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

/**
 * The order-preserving twin of `parser`.
 *
 * The default parse groups same-named siblings into one array per tag, which
 * destroys the interleaving BETWEEN tags: a paragraph holding del, ins, del,
 * ins parses to `{'w:del': [d1, d3], 'w:ins': [i2, i4]}`, indistinguishable
 * from del, del, ins, ins. Anything that has to know which element followed
 * which — pairing a tracked deletion with the insertion that replaced it —
 * cannot be answered from that shape at any cost.
 *
 * `preserveOrder` answers it, at the price of a different shape: every element
 * becomes a single-key object inside an ordered sibling array, with attributes
 * under `:@` instead of inline. That is why it is a SECOND parser rather than a
 * flag on the first: every existing walker here reads the grouped shape, and
 * switching them all over to buy one feature would be a rewrite.
 */
const orderedParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  preserveOrder: true,
});

const parseXml = (xml: string | undefined): unknown => {
  if (xml === undefined || xml === '') return undefined;
  return parser.parse(xml) as unknown;
};

const walkVisit = (node: unknown, visit: (key: string, value: unknown) => void): void => {
  if (Array.isArray(node)) {
    for (const item of node) walkVisit(item, visit);
    return;
  }
  if (!isObject(node)) return;
  for (const [key, value] of Object.entries(node)) {
    if (key.startsWith('@_') || key === '#text') continue;
    visit(key, value);
    walkVisit(value, visit);
  }
};

/**
 * An XML namespace prefix (`p:`, `p188:`, `a:`) is a per-document alias for a
 * namespace URI, not part of the schema: another writer may bind the same URI to
 * a different prefix. The `*Local` walkers match on the local name alone, so an
 * element is found whatever prefix its author chose. Use them only where the local
 * name is unambiguous inside the part being read: `w:t` and `a:t` share the local
 * name `t` in a .docx, so paragraph text stays on the qualified match.
 */
const localNameOf = (key: string): string => key.slice(key.indexOf(':') + 1);

const findAllBy = (root: unknown, matches: (key: string) => boolean): ReadonlyArray<XmlObject> => {
  const out: Array<XmlObject> = [];
  walkVisit(root, (key, value) => {
    if (!matches(key)) return;
    if (Array.isArray(value)) {
      for (const item of value) if (isObject(item)) out.push(item);
      return;
    }
    if (isObject(value)) out.push(value);
  });
  return out;
};

const findAll = (root: unknown, tagName: string): ReadonlyArray<XmlObject> => findAllBy(root, (key) => key === tagName);

const findAllLocal = (root: unknown, localName: string): ReadonlyArray<XmlObject> => findAllBy(root, (key) => localNameOf(key) === localName);

const textOf = (node: unknown): string => {
  if (typeof node === 'string') return node;
  if (!isObject(node)) return '';
  const t = node['#text'];
  return typeof t === 'string' ? t : '';
};

const attrOf = (node: XmlObject, name: string): string => {
  const v = node[`@_${name}`];
  return typeof v === 'string' ? v : '';
};

/**
 * Yield each occurrence of `tagName` inside `node` as its own text value.
 * Used for leaf-text elements like `<w:instrText>` where each instance is a
 * distinct entry (one MERGEFIELD per occurrence) — unlike collectText, which
 * flattens every match into a single string for "the visible text of this run".
 */
const findAllTexts = (root: unknown, tagName: string): ReadonlyArray<string> => {
  const out: Array<string> = [];
  walkVisit(root, (key, value) => {
    if (key !== tagName) return;
    if (Array.isArray(value)) {
      for (const item of value) out.push(textOf(item));
      return;
    }
    out.push(textOf(value));
  });
  return out;
};

/**
 * Concatenate the text content of every element matching `tagName` inside `node`,
 * regardless of nesting depth. Used to flatten a `<w:p>` (or `<w:ins>` / `<w:comment>`)
 * down to its visible text by gathering every `<w:t>` (or `<w:delText>`) descendant.
 */
const collectTextBy = (node: unknown, matches: (key: string) => boolean): string => {
  let result = '';
  walkVisit(node, (key, value) => {
    if (!matches(key)) return;
    if (Array.isArray(value)) {
      for (const item of value) result += textOf(item);
      return;
    }
    result += textOf(value);
  });
  return result;
};

const collectText = (node: unknown, tagName: string): string => collectTextBy(node, (key) => key === tagName);

const collectTextLocal = (node: unknown, localName: string): string => collectTextBy(node, (key) => localNameOf(key) === localName);

/**
 * One element in the `preserveOrder` shape: its tag name, and the node itself
 * (whose `[tag]` holds the ordered children and whose `:@` holds attributes).
 */
type OrderedNode = { readonly tag: string; readonly node: XmlObject };

const parseXmlOrdered = (xml: string | undefined): unknown => {
  if (xml === undefined || xml === '') return undefined;
  return orderedParser.parse(xml) as unknown;
};

const orderedTagOf = (node: XmlObject): string | undefined => Object.keys(node).find((k) => k !== ':@');

const orderedChildrenOf = (node: XmlObject, tag: string): ReadonlyArray<unknown> => {
  const value = node[tag];
  return Array.isArray(value) ? value : [];
};

/**
 * Every sibling list in the tree, each in document order, elements only.
 *
 * Sibling lists rather than one flat stream: adjacency only means anything
 * WITHIN a parent. A deletion ending one paragraph and an insertion opening the
 * next are consecutive in a flat walk and unrelated in the document.
 */
const orderedSiblingGroups = (root: unknown): ReadonlyArray<ReadonlyArray<OrderedNode>> => {
  const groups: Array<ReadonlyArray<OrderedNode>> = [];
  const visit = (siblings: unknown): void => {
    if (!Array.isArray(siblings)) return;
    const group: Array<OrderedNode> = [];
    for (const item of siblings) {
      if (!isObject(item)) continue;
      const tag = orderedTagOf(item);
      if (tag === undefined || tag === '#text') continue;
      group.push({ tag, node: item });
    }
    if (group.length > 0) groups.push(group);
    for (const entry of group) visit(entry.node[entry.tag]);
  };
  visit(root);
  return groups;
};

/**
 * Every element in true document order, pre-order (an element, then what it
 * contains). Unlike `orderedSiblingGroups`, which reports a whole sibling level
 * before descending, this is the order the start tags appear in the file, which
 * is what reading flat range markers requires: `<w:moveFromRangeStart>` opens a
 * span that a later element sits inside without being its child.
 */
const orderedElements = (root: unknown): ReadonlyArray<OrderedNode> => {
  const out: Array<OrderedNode> = [];
  const visit = (siblings: unknown): void => {
    if (!Array.isArray(siblings)) return;
    for (const item of siblings) {
      if (!isObject(item)) continue;
      const tag = orderedTagOf(item);
      if (tag === undefined || tag === '#text') continue;
      out.push({ tag, node: item });
      visit(item[tag]);
    }
  };
  visit(root);
  return out;
};

const orderedAttrOf = (node: XmlObject, name: string): string => {
  const attrs = node[':@'];
  if (!isObject(attrs)) return '';
  const value = attrs[`@_${name}`];
  return typeof value === 'string' ? value : '';
};

/** The `collectText` of the ordered shape: flatten every `tagName` descendant to one string. */
const collectOrderedText = (node: XmlObject, tagName: string): string => {
  let out = '';
  const visit = (current: XmlObject): void => {
    const tag = orderedTagOf(current);
    if (tag === undefined) return;
    const children = orderedChildrenOf(current, tag);
    if (tag === tagName) {
      for (const child of children) {
        if (isObject(child) && typeof child['#text'] === 'string') out += child['#text'];
      }
      return;
    }
    for (const child of children) if (isObject(child)) visit(child);
  };
  visit(node);
  return out;
};

export {
  attrOf,
  collectOrderedText,
  collectText,
  collectTextLocal,
  findAll,
  findAllLocal,
  findAllTexts,
  orderedAttrOf,
  orderedElements,
  orderedSiblingGroups,
  parseXml,
  parseXmlOrdered,
  textOf,
};
export type { OrderedNode, XmlObject };
