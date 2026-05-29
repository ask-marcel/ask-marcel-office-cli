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

const findAll = (root: unknown, tagName: string): ReadonlyArray<XmlObject> => {
  const out: Array<XmlObject> = [];
  walkVisit(root, (key, value) => {
    if (key !== tagName) return;
    if (Array.isArray(value)) {
      for (const item of value) if (isObject(item)) out.push(item);
      return;
    }
    if (isObject(value)) out.push(value);
  });
  return out;
};

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
const collectText = (node: unknown, tagName: string): string => {
  let result = '';
  walkVisit(node, (key, value) => {
    if (key !== tagName) return;
    if (Array.isArray(value)) {
      for (const item of value) result += textOf(item);
      return;
    }
    result += textOf(value);
  });
  return result;
};

export { attrOf, collectText, findAll, findAllTexts, parseXml, textOf };
export type { XmlObject };
