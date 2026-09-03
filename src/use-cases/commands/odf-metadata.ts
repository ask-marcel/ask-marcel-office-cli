import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import { openOoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { OrderedNode, XmlObject } from './ooxml-xml-walker.ts';
import { attrOf, collectOrderedText, findAll, findAllTexts, orderedAttrOf, orderedElements, parseXml, parseXmlOrdered, textOf } from './ooxml-xml-walker.ts';

/**
 * Pulls the side-channel content out of an OpenDocument package (.odt / .ods /
 * .odp and their .ot* template variants). OpenDocument is also a ZIP, so the
 * shared `openOoxmlZip` adapter + XML walker apply directly; only the metadata
 * shape differs — it lives in a single `meta.xml` (`office:document-meta >
 * office:meta`) rather than the OOXML `docProps/*` parts.
 *
 * High-value subset: document properties (Dublin Core + ODF meta fields like
 * generator / editing-cycles / creation-date), keyword list, and user-defined
 * custom properties (`meta:user-defined`) — the ODF analog of OOXML custom
 * document properties.
 */

type UserDefined = { readonly name: string; readonly value: string };
type OdfTrackedChange = { readonly id: string; readonly author: string; readonly date: string; readonly text: string };
type OdfReplacement = {
  readonly deletionId: string;
  readonly insertionId: string;
  readonly author: string;
  readonly date: string;
  readonly before: string;
  readonly after: string;
};
/** ODF records who and when for a format change, never which properties moved. */
type OdfFormatChange = { readonly author: string; readonly date: string; readonly text: string };
type OdfMetadata = {
  readonly insertions: ReadonlyArray<OdfTrackedChange>;
  readonly deletions: ReadonlyArray<OdfTrackedChange>;
  /** A deletion mark and the insertion beside it, by one author: one edit, reported once. */
  readonly replacements: ReadonlyArray<OdfReplacement>;
  readonly formatChanges: ReadonlyArray<OdfFormatChange>;
  readonly properties: Readonly<Record<string, string>>;
  readonly keywords: ReadonlyArray<string>;
  readonly userDefined: ReadonlyArray<UserDefined>;
};

// Tags handled by their own section, so they are not folded into the flat
// property record.
const SPECIAL_TAGS: ReadonlySet<string> = new Set(['meta:keyword', 'meta:user-defined']);

const stripNs = (tag: string): string => {
  const colon = tag.indexOf(':');
  return colon === -1 ? tag : tag.slice(colon + 1);
};

const extractProperties = (meta: unknown): Readonly<Record<string, string>> => {
  const container = findAll(meta, 'office:meta')[0];
  if (container === undefined) return {};
  const out: Record<string, string> = {};
  for (const [key, value] of Object.entries(container)) {
    if (key.startsWith('@_') || key === '#text' || SPECIAL_TAGS.has(key)) continue;
    const text = textOf(value);
    if (text !== '') out[stripNs(key)] = text;
  }
  return out;
};

const extractUserDefined = (meta: unknown): ReadonlyArray<UserDefined> =>
  findAll(meta, 'meta:user-defined')
    .map((p) => ({ name: attrOf(p, 'meta:name'), value: textOf(p) }))
    .filter((p) => p.name !== '');

// ── Tracked changes ────────────────────────────────────────────────────────
// ODF declares every change once under `text:tracked-changes` (kind, author,
// date, and for a deletion the deleted paragraphs) and anchors it in the body
// with MARKS: `text:change` for a deletion, `text:change-start` / `-end` around
// inserted or reformatted text. So an insertion's text is read off the body
// between its marks, a deletion's off its region, and a replacement is
// adjacency of marks in reading order, not of sibling elements as in docx.
type RegionKind = 'insertion' | 'deletion' | 'format-change';
type Region = { readonly id: string; readonly kind: RegionKind; readonly author: string; readonly date: string; readonly deleted: string };
type Event =
  | { readonly kind: 'del'; readonly id: string }
  | { readonly kind: 'start'; readonly id: string }
  | { readonly kind: 'end'; readonly id: string }
  | { readonly kind: 'text' }
  | { readonly kind: 'break' };

const tagOfOrdered = (node: XmlObject): string | undefined => Object.keys(node).find((k) => k !== ':@');
const childrenOf = (node: XmlObject, tag: string): ReadonlyArray<unknown> => {
  const value = node[tag];
  return Array.isArray(value) ? value : [];
};
const isNode = (value: unknown): value is XmlObject => value !== null && typeof value === 'object' && !Array.isArray(value);

// The visible text under an ODF node: `text:s` is a run of spaces, tab and
// line-break read as one space, annotations are not body text.
const odfText = (children: ReadonlyArray<unknown>): string => {
  let out = '';
  for (const child of children) {
    if (!isNode(child)) continue;
    const leaf = child['#text'];
    if (typeof leaf === 'string') {
      out += leaf;
      continue;
    }
    const tag = tagOfOrdered(child);
    if (tag === undefined || tag === 'office:annotation' || tag === 'office:annotation-end') continue;
    if (tag === 'text:s') {
      const count = Number.parseInt(orderedAttrOf(child, 'text:c'), 10);
      out += ' '.repeat(Number.isFinite(count) && count > 0 ? count : 1);
    } else if (tag === 'text:tab' || tag === 'text:line-break') out += ' ';
    else out += odfText(childrenOf(child, tag));
  }
  return out;
};

const REGION_KINDS: ReadonlyArray<RegionKind> = ['insertion', 'deletion', 'format-change'];

const regionOf = (entry: OrderedNode): Region | undefined => {
  const kinds = childrenOf(entry.node, entry.tag).filter(isNode);
  const body = kinds.find((k) => REGION_KINDS.some((kind) => tagOfOrdered(k) === `text:${kind}`));
  if (body === undefined) return undefined;
  const tag = tagOfOrdered(body) ?? '';
  const kind = tag.slice('text:'.length) as RegionKind;
  const paragraphs = childrenOf(body, tag).filter((c) => isNode(c) && (tagOfOrdered(c) === 'text:p' || tagOfOrdered(c) === 'text:h'));
  // Who and when come from the region's own change-info, never from the body
  // subtree: a deleted paragraph can carry an annotation with a creator of its
  // own, and reading the subtree would fuse the two names.
  const info = childrenOf(body, tag)
    .filter(isNode)
    .find((c) => tagOfOrdered(c) === 'office:change-info');
  return {
    id: orderedAttrOf(entry.node, 'text:id'),
    kind,
    author: info === undefined ? '' : collectOrderedText(info, 'dc:creator'),
    date: info === undefined ? '' : collectOrderedText(info, 'dc:date'),
    deleted: paragraphs
      .map((p) => (isNode(p) ? odfText(childrenOf(p, tagOfOrdered(p) ?? '')) : ''))
      .join(' ')
      .trim(),
  };
};

const BLOCK_TAGS: ReadonlySet<string> = new Set(['text:p', 'text:h', 'text:list-item', 'table:table-cell']);

// Walk the body in reading order, skipping the tracked-changes declarations,
// and emit what pairing needs: marks, "some text passed", and block boundaries.
// Spans collect the text between their start and end marks along the way.
const walkBody = (children: ReadonlyArray<unknown>, events: Array<Event>, spans: Map<string, string>, open: Set<string>): void => {
  for (const child of children) {
    if (!isNode(child)) continue;
    const leaf = child['#text'];
    if (typeof leaf === 'string') {
      if (leaf !== '') events.push({ kind: 'text' });
      for (const id of open) spans.set(id, (spans.get(id) ?? '') + leaf);
      continue;
    }
    const tag = tagOfOrdered(child);
    if (tag === undefined || tag === 'text:tracked-changes' || tag === 'office:annotation' || tag === 'office:annotation-end') continue;
    if (tag === 'text:change') {
      events.push({ kind: 'del', id: orderedAttrOf(child, 'text:change-id') });
      continue;
    }
    if (tag === 'text:change-start') {
      const id = orderedAttrOf(child, 'text:change-id');
      events.push({ kind: 'start', id });
      open.add(id);
      spans.set(id, spans.get(id) ?? '');
      continue;
    }
    if (tag === 'text:change-end') {
      const id = orderedAttrOf(child, 'text:change-id');
      events.push({ kind: 'end', id });
      open.delete(id);
      continue;
    }
    if (tag === 'text:s' || tag === 'text:tab' || tag === 'text:line-break') {
      const piece = tag === 'text:s' ? ' '.repeat(Math.max(1, Number.parseInt(orderedAttrOf(child, 'text:c'), 10) || 1)) : ' ';
      events.push({ kind: 'text' });
      for (const id of open) spans.set(id, (spans.get(id) ?? '') + piece);
      continue;
    }
    walkBody(childrenOf(child, tag), events, spans, open);
    if (BLOCK_TAGS.has(tag)) events.push({ kind: 'break' });
  }
};

// A replacement is a deletion mark and an insertion span touching each other in
// reading order, in either order, by the same author, both carrying text. Any
// text or block boundary between them is two edits, not one; a mark of another
// change is transparent only if it is the partner itself.
const pairReplacements = (events: ReadonlyArray<Event>, byId: ReadonlyMap<string, Region>, spans: ReadonlyMap<string, string>): ReadonlyArray<OdfReplacement> => {
  const out: Array<OdfReplacement> = [];
  const pair = (delId: string, insId: string): void => {
    const del = byId.get(delId);
    const ins = byId.get(insId);
    const after = (spans.get(insId) ?? '').trim();
    if (del === undefined || ins === undefined || del.kind !== 'deletion' || ins.kind !== 'insertion') return;
    if (del.author !== ins.author || del.deleted === '' || after === '') return;
    out.push({ deletionId: delId, insertionId: insId, author: del.author, date: del.date, before: del.deleted, after });
  };
  for (let i = 0; i + 1 < events.length; i += 1) {
    const here = events[i];
    const next = events[i + 1];
    if (here === undefined || next === undefined) continue;
    if (here.kind === 'del' && next.kind === 'start') pair(here.id, next.id);
    else if (here.kind === 'end' && next.kind === 'del') pair(next.id, here.id);
  }
  return out;
};

const extractTrackedChanges = (contentXml: string | undefined): Pick<OdfMetadata, 'insertions' | 'deletions' | 'replacements' | 'formatChanges'> => {
  const root = parseXmlOrdered(contentXml);
  const regions = orderedElements(root)
    .filter((e) => e.tag === 'text:changed-region')
    .map(regionOf)
    .filter((r): r is Region => r !== undefined);
  const byId = new Map(regions.map((r) => [r.id, r] as const));
  const events: Array<Event> = [];
  const spans = new Map<string, string>();
  const textBodies = orderedElements(root).filter((e) => e.tag === 'office:text');
  for (const body of textBodies) walkBody(childrenOf(body.node, body.tag), events, spans, new Set());
  const replacements = pairReplacements(events, byId, spans);
  const paired = new Set(replacements.flatMap((r) => [r.deletionId, r.insertionId]));
  const change = (r: Region, text: string): OdfTrackedChange => ({ id: r.id, author: r.author, date: r.date, text });
  return {
    insertions: regions
      .filter((r) => r.kind === 'insertion' && !paired.has(r.id))
      .map((r) => change(r, (spans.get(r.id) ?? '').trim()))
      .filter((c) => c.text !== ''),
    deletions: regions
      .filter((r) => r.kind === 'deletion' && !paired.has(r.id))
      .map((r) => change(r, r.deleted))
      .filter((c) => c.text !== ''),
    replacements,
    formatChanges: regions
      .filter((r) => r.kind === 'format-change')
      .map((r) => ({ author: r.author, date: r.date, text: (spans.get(r.id) ?? '').trim() }))
      .filter((f) => f.text !== ''),
  };
};

const extractOdfMetadata = async (bytes: Uint8Array): Promise<Result<OdfMetadata, GraphError>> => {
  const zipR = await openOoxmlZip(bytes);
  if (!zipR.ok) return zipR;
  const meta = parseXml(zipR.value.read('meta.xml'));
  return ok({
    properties: extractProperties(meta),
    keywords: findAllTexts(meta, 'meta:keyword').filter((k) => k.trim() !== ''),
    ...extractTrackedChanges(zipR.value.read('content.xml')),
    userDefined: extractUserDefined(meta),
  });
};

export { extractOdfMetadata };
export type { OdfFormatChange, OdfMetadata, OdfReplacement, OdfTrackedChange, UserDefined };
