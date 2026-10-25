import { XMLParser } from 'fast-xml-parser';

/**
 * Correlates each docx comment to the body text it annotates. In
 * `word/document.xml` a comment's anchored span is bracketed by empty marker
 * elements `<w:commentRangeStart w:id="N"/> …runs… <w:commentRangeEnd w:id="N"/>`,
 * keyed by the same `w:id` the comment in `word/comments.xml` carries. mammoth
 * strips those markers from its HTML, so to fold the comment's context into the
 * markdown we read document.xml directly.
 *
 * The shared walker (ooxml-xml-walker) groups same-named siblings and loses
 * cross-element order, which would scramble which runs fall inside a range. So,
 * like odf-content-to-markdown, this parses with `preserveOrder` and walks the
 * ordered tree, tracking the set of currently-open comment ids and appending
 * each `<w:t>` run's text to every open range. Handles overlapping / nested
 * ranges (more than one open at once). Pure (string → Map); no IO.
 */

type Node = Record<string, unknown>;

const TEXT = '#text';
const ATTRS = ':@';
// Clamp the quoted span so a comment over a whole page doesn't bloat the output.
const MAX_ANCHOR_LEN = 200;

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', parseAttributeValue: false, parseTagValue: false, trimValues: false, preserveOrder: true });

const tagOf = (node: Node): string | undefined => {
  for (const key of Object.keys(node)) if (key !== ATTRS && key !== TEXT) return key;
  return undefined;
};

const attr = (node: Node, name: string): string => {
  const bag = node[ATTRS];
  if (bag === null || typeof bag !== 'object') return '';
  const value = (bag as Record<string, unknown>)[`@_${name}`];
  return typeof value === 'string' ? value : '';
};

const leafText = (children: ReadonlyArray<Node>): string => {
  let out = '';
  for (const child of children) {
    const leaf = child[TEXT];
    if (typeof leaf === 'string') out += leaf;
  }
  return out;
};

const clamp = (span: string): string => {
  const trimmed = span.trim();
  return trimmed.length > MAX_ANCHOR_LEN ? `${trimmed.slice(0, MAX_ANCHOR_LEN)}…` : trimmed;
};

const extractCommentAnchors = (documentXml: string | undefined): ReadonlyMap<string, string> => {
  const spans = new Map<string, string>();
  if (documentXml === undefined || documentXml === '') return spans;
  const open = new Set<string>();
  const visit = (nodes: ReadonlyArray<Node>): void => {
    for (const node of nodes) {
      const tag = tagOf(node);
      if (tag === undefined) continue;
      if (tag === 'w:commentRangeStart') {
        open.add(attr(node, 'w:id'));
        continue;
      }
      if (tag === 'w:commentRangeEnd') {
        open.delete(attr(node, 'w:id'));
        continue;
      }
      const children = node[tag];
      if (!Array.isArray(children)) continue;
      if (tag === 'w:t') {
        if (open.size > 0) {
          const text = leafText(children as ReadonlyArray<Node>);
          for (const id of open) spans.set(id, (spans.get(id) ?? '') + text);
        }
        continue;
      }
      visit(children as ReadonlyArray<Node>);
    }
  };
  visit(parser.parse(documentXml) as ReadonlyArray<Node>);

  const out = new Map<string, string>();
  for (const [id, span] of spans) {
    const clamped = clamp(span);
    if (clamped !== '') out.set(id, clamped);
  }
  return out;
};

// Pure preserveOrder helpers, exported for direct unit testing (their
// edge branches — `:@`/`#text` keys, a null/absent attr bag, non-string
// attr values — never arise from real parser output, so they're only
// reachable via hand-constructed nodes).
export { attr, clamp, extractCommentAnchors, leafText, tagOf };
