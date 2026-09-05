import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { OoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import { openOoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { extractAppProps, extractCoreProps, extractCustomProps, extractExternalRels, extractMacros } from './ooxml-metadata.ts';
import type { CustomProp, ExternalRel } from './ooxml-metadata.ts';
import { extractCommentAnchors } from './docx-comment-anchors.ts';
import type { OrderedNode, XmlObject } from './ooxml-xml-walker.ts';
import {
  attrOf,
  collectOrderedText,
  collectText,
  findAll,
  findAllTexts,
  orderedAttrOf,
  orderedElements,
  orderedSiblingGroups,
  parseXml,
  parseXmlOrdered,
} from './ooxml-xml-walker.ts';

/**
 * Pulls the side-channel content out of a .docx zip — every text-bearing
 * surface mammoth drops on the floor: core / app / custom doc properties,
 * people registry, external hyperlinks, comments, tracked changes (replacements,
 * the insertions and deletions that pair with nothing, moves, and run,
 * paragraph, table, row and cell formatting changes),
 * hidden text (w:vanish), text-box / shape text (w:txbxContent), header/footer
 * body prose, field instructions (MERGEFIELD / HYPERLINK / DOCVARIABLE), bookmarks.
 *
 * The package-level parts (docProps/*, every *.rels) come from the shared
 * ooxml-metadata module; this file owns only the docx-body-specific parts.
 *
 * Pure use-case logic — no IO. The zip is opened upstream via the infra
 * adapter; this module just walks parsed XML trees. Try/catch lives in
 * the infra adapter, not here.
 */

type CoreProps = Readonly<Record<string, string>>;
type AppProps = Readonly<Record<string, string>>;
type Person = { readonly author: string; readonly providerId: string; readonly userId: string };
type Comment = { readonly id: string; readonly author: string; readonly initials: string; readonly date: string; readonly text: string; readonly anchor?: string };
type TrackedChange = { readonly id: string; readonly author: string; readonly date: string; readonly text: string };
type Replacement = { readonly deletionId: string; readonly insertionId: string; readonly author: string; readonly date: string; readonly before: string; readonly after: string };
type Move = { readonly name: string; readonly author: string; readonly date: string; readonly text: string; readonly halves: 'both' | 'from-only' | 'to-only' };
type FormatChangeScope = 'run' | 'paragraph' | 'table' | 'row' | 'cell';
type FormatChange = { readonly scope: FormatChangeScope; readonly author: string; readonly date: string; readonly text: string; readonly properties: ReadonlyArray<string> };
type Field = { readonly source: string; readonly instruction: string };
type Bookmark = { readonly id: string; readonly name: string };
type HeaderFooter = { readonly part: string; readonly text: string };

type DocxMetadata = {
  readonly core: CoreProps;
  readonly app: AppProps;
  readonly custom: ReadonlyArray<CustomProp>;
  readonly people: ReadonlyArray<Person>;
  readonly externalRels: ReadonlyArray<ExternalRel>;
  readonly comments: ReadonlyArray<Comment>;
  readonly insertions: ReadonlyArray<TrackedChange>;
  readonly deletions: ReadonlyArray<TrackedChange>;
  /** A deletion and the insertion beside it, reported as the one edit they are. */
  readonly replacements: ReadonlyArray<Replacement>;
  /** Text moved elsewhere, joined by the range name that brackets both halves. */
  readonly moves: ReadonlyArray<Move>;
  /** Run, paragraph, table, row or cell properties changed under revision marking. */
  readonly formatChanges: ReadonlyArray<FormatChange>;
  readonly hiddenText: ReadonlyArray<string>;
  readonly textBoxes: ReadonlyArray<string>;
  readonly headersFooters: ReadonlyArray<HeaderFooter>;
  readonly fields: ReadonlyArray<Field>;
  readonly bookmarks: ReadonlyArray<Bookmark>;
  readonly macros: ReadonlyArray<string>;
};

const extractPeople = (root: unknown): ReadonlyArray<Person> => {
  const persons = findAll(root, 'w15:person');
  return persons.map((p) => {
    const presence = findAll(p, 'w15:presenceInfo')[0] ?? {};
    return {
      author: attrOf(p, 'w15:author'),
      providerId: attrOf(presence, 'w15:providerId'),
      userId: attrOf(presence, 'w15:userId'),
    };
  });
};

const extractComments = (root: unknown, anchors: ReadonlyMap<string, string>): ReadonlyArray<Comment> => {
  const comments = findAll(root, 'w:comment');
  return comments.map((c) => {
    const id = attrOf(c, 'w:id');
    const base: Comment = { id, author: attrOf(c, 'w:author'), initials: attrOf(c, 'w:initials'), date: attrOf(c, 'w:date'), text: collectText(c, 'w:t') };
    const anchor = anchors.get(id);
    return anchor === undefined ? base : { ...base, anchor };
  });
};

// OOXML has no "replace" revision. Word records replacing a span as a deletion
// sitting next to an insertion, so reported as two loose halves one edit reads
// as an unrelated cut plus an unrelated addition, and nothing links them.
//
// Pairing needs document order, which the default parse cannot express (see
// `orderedParser`), so this walks the order-preserving tree instead. Scoped to
// siblings inside one parent: a deletion closing a paragraph and an insertion
// opening the next are adjacent in a flat walk and unrelated on the page.
const REVISION_TEXT_TAG: Readonly<Record<string, string>> = { 'w:ins': 'w:t', 'w:del': 'w:delText' };

// Between the two halves of one edit Word may write markers that render no
// glyph: its spell-check hints, and bookmark / comment range anchors. None of
// them separates the halves. An untouched run of prose does.
const TRANSPARENT_TAGS: ReadonlySet<string> = new Set(['w:proofErr', 'w:bookmarkStart', 'w:bookmarkEnd', 'w:commentRangeStart', 'w:commentRangeEnd']);

type Revision = { readonly kind: string; readonly id: string; readonly author: string; readonly date: string; readonly text: string };

const revisionOf = (entry: OrderedNode): Revision | undefined => {
  const textTag = REVISION_TEXT_TAG[entry.tag];
  if (textTag === undefined) return undefined;
  const text = collectOrderedText(entry.node, textTag);
  if (text === '') return undefined;
  return { kind: entry.tag, id: orderedAttrOf(entry.node, 'w:id'), author: orderedAttrOf(entry.node, 'w:author'), date: orderedAttrOf(entry.node, 'w:date'), text };
};

// Same author on both halves is required: one person deleting and another
// inserting beside it is two people disagreeing, not one person replacing a
// span, and reporting it as a replacement would misattribute both edits.
const pairInGroup = (group: ReadonlyArray<OrderedNode>): ReadonlyArray<Replacement> => {
  const out: Array<Replacement> = [];
  let pending: Revision | undefined;
  for (const entry of group) {
    if (TRANSPARENT_TAGS.has(entry.tag)) continue;
    const revision = revisionOf(entry);
    if (revision === undefined) {
      pending = undefined;
      continue;
    }
    if (pending !== undefined && pending.kind !== revision.kind && pending.author === revision.author) {
      const deletion = pending.kind === 'w:del' ? pending : revision;
      const insertion = pending.kind === 'w:ins' ? pending : revision;
      out.push({ deletionId: deletion.id, insertionId: insertion.id, author: deletion.author, date: deletion.date, before: deletion.text, after: insertion.text });
      pending = undefined;
      continue;
    }
    pending = revision;
  }
  return out;
};

const extractReplacements = (documentXml: string | undefined): ReadonlyArray<Replacement> =>
  orderedSiblingGroups(parseXmlOrdered(documentXml)).flatMap((group) => pairInGroup(group));

const extractTracked = (root: unknown, kind: 'w:ins' | 'w:del'): ReadonlyArray<TrackedChange> => {
  const nodes = findAll(root, kind);
  const textTag = kind === 'w:ins' ? 'w:t' : 'w:delText';
  return nodes.map((n) => ({ id: attrOf(n, 'w:id'), author: attrOf(n, 'w:author'), date: attrOf(n, 'w:date'), text: collectText(n, textTag) })).filter((t) => t.text !== '');
};

// A `<w:r>` is hidden when its `<w:rPr>` carries a `<w:vanish/>` child.
// Walk all runs, check each one's rPr for the vanish flag.
const extractHidden = (root: unknown): ReadonlyArray<string> => {
  const runs = findAll(root, 'w:r');
  const out: Array<string> = [];
  for (const r of runs) {
    const rPr = r['w:rPr'];
    if (!rPr || typeof rPr !== 'object') continue;
    if (!Object.hasOwn(rPr, 'w:vanish')) continue;
    const text = collectText(r, 'w:t');
    if (text !== '') out.push(text);
  }
  return out;
};

const extractFieldsFromOne = (root: unknown, source: string): ReadonlyArray<Field> => {
  const out: Array<Field> = [];
  for (const text of findAllTexts(root, 'w:instrText')) {
    const instr = text.trim();
    if (instr !== '') out.push({ source, instruction: instr });
  }
  for (const fs of findAll(root, 'w:fldSimple')) {
    const instr = attrOf(fs, 'w:instr').trim();
    if (instr !== '') out.push({ source, instruction: instr });
  }
  return out;
};

const extractBookmarks = (root: unknown): ReadonlyArray<Bookmark> => {
  const nodes = findAll(root, 'w:bookmarkStart');
  return nodes.map((b) => ({ id: attrOf(b, 'w:id'), name: attrOf(b, 'w:name') })).filter((b) => b.name !== '');
};

const headerFooterPaths = (zip: OoxmlZip): ReadonlyArray<string> => zip.list().filter((p) => /^word\/(header|footer)\d+\.xml$/.test(p));

// Field codes live in the body and in every header/footer part. Discover the
// header/footer parts dynamically (not a fixed header1..3/footer1..3 list, which
// silently misses header4+/footer4+ and is brittle to renumbering).
const collectFields = (zip: OoxmlZip): ReadonlyArray<Field> => {
  const out: Array<Field> = [];
  for (const path of ['word/document.xml', ...headerFooterPaths(zip)]) {
    const parsed = parseXml(zip.read(path));
    if (parsed === undefined) continue;
    for (const f of extractFieldsFromOne(parsed, path)) out.push(f);
  }
  return out;
};

// Header/footer body prose — mammoth drops headers/footers entirely, and `collectFields`
// only pulls their field codes, so the regular paragraph text is captured here.
const extractHeadersFooters = (zip: OoxmlZip): ReadonlyArray<HeaderFooter> => {
  const out: Array<HeaderFooter> = [];
  for (const part of headerFooterPaths(zip)) {
    const text = collectText(parseXml(zip.read(part)), 'w:t').trim();
    if (text !== '') out.push({ part, text });
  }
  return out;
};

// Text-box / shape prose (`w:txbxContent`) anywhere in the body or headers/footers —
// neither mammoth nor any other side-channel extractor surfaces it.
const extractTextBoxes = (zip: OoxmlZip): ReadonlyArray<string> => {
  const out: Array<string> = [];
  for (const part of ['word/document.xml', ...headerFooterPaths(zip)]) {
    const parsed = parseXml(zip.read(part));
    if (parsed === undefined) continue;
    for (const box of findAll(parsed, 'w:txbxContent')) {
      const text = collectText(box, 'w:t').trim();
      if (text !== '') out.push(text);
    }
  }
  return out;
};

// A move is two halves in two places, and neither `w:moveFrom` nor `w:moveTo`
// carries anything linking them. The link is the `w:name` on the range-start
// markers that BRACKET them: flat elements opening a span the halves sit inside
// without being their children, which is why this reads document order rather
// than structure.
//
// Not paired by matching text, though Word does guarantee the halves match: a
// document that moves the same sentence twice then cross-pairs four halves into
// two wrong moves, and nothing in the output would show it happened.
type MoveHalf = { readonly author: string; readonly date: string; readonly text: string };

const moveHalfOf = (node: XmlObject, textTag: string): MoveHalf => ({
  author: orderedAttrOf(node, 'w:author'),
  date: orderedAttrOf(node, 'w:date'),
  text: collectOrderedText(node, textTag),
});

const halvesOf = (from: MoveHalf | undefined, to: MoveHalf | undefined): Move['halves'] => {
  if (from !== undefined && to !== undefined) return 'both';
  return from === undefined ? 'to-only' : 'from-only';
};

const extractMoves = (documentXml: string | undefined): ReadonlyArray<Move> => {
  const froms = new Map<string, MoveHalf>();
  const tos = new Map<string, MoveHalf>();
  let openFrom: string | undefined;
  let openTo: string | undefined;
  for (const element of orderedElements(parseXmlOrdered(documentXml))) {
    if (element.tag === 'w:moveFromRangeStart') openFrom = orderedAttrOf(element.node, 'w:name');
    else if (element.tag === 'w:moveFromRangeEnd') openFrom = undefined;
    else if (element.tag === 'w:moveToRangeStart') openTo = orderedAttrOf(element.node, 'w:name');
    else if (element.tag === 'w:moveToRangeEnd') openTo = undefined;
    else if (element.tag === 'w:moveFrom' && openFrom !== undefined) froms.set(openFrom, moveHalfOf(element.node, 'w:delText'));
    else if (element.tag === 'w:moveTo' && openTo !== undefined) tos.set(openTo, moveHalfOf(element.node, 'w:t'));
  }
  const names = [...new Set([...froms.keys(), ...tos.keys()])];
  const moves: Array<Move> = [];
  for (const name of names) {
    const from = froms.get(name);
    const to = tos.get(name);
    // A half whose partner never arrives still gets a row. Dropping it would
    // put that text back where this whole change started: reported as nothing.
    const half = from ?? to;
    if (half === undefined || half.text === '') continue;
    moves.push({ name, author: half.author, date: half.date, text: half.text, halves: halvesOf(from, to) });
  }
  return moves;
};

// `w:rPrChange` / `w:pPrChange` hold the properties as they were BEFORE the
// edit; the properties as they are now are its siblings in the enclosing
// `w:rPr` / `w:pPr`. So a real before/after is available, and what makes it
// useful is naming which properties moved.
//
// Compared by attributes as well as tag name: `w:color` is present on both
// sides of a recolour and only its `w:val` moves, so a comparison of tag names
// alone reports a recoloured run as unchanged.
const attrSignature = (value: unknown): string => {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return '';
  return Object.entries(value as Record<string, unknown>)
    .filter(([key]) => key.startsWith('@_'))
    .map(([key, attr]) => `${key}=${String(attr)}`)
    .toSorted((a, b) => a.localeCompare(b))
    .join(';');
};

const propertiesOf = (props: unknown, exclude: string): ReadonlyMap<string, string> => {
  const out = new Map<string, string>();
  if (props === null || typeof props !== 'object' || Array.isArray(props)) return out;
  for (const [key, value] of Object.entries(props as Record<string, unknown>)) {
    if (key.startsWith('@_') || key === '#text' || key === exclude) continue;
    out.set(key, attrSignature(value));
  }
  return out;
};

const changedProperties = (before: ReadonlyMap<string, string>, after: ReadonlyMap<string, string>): ReadonlyArray<string> =>
  [...new Set([...before.keys(), ...after.keys()])].filter((name) => before.get(name) !== after.get(name)).toSorted((a, b) => a.localeCompare(b));

const changesIn = (root: unknown, container: string, propsTag: string, changeTag: string, scope: FormatChangeScope): ReadonlyArray<FormatChange> => {
  const out: Array<FormatChange> = [];
  for (const node of findAll(root, container)) {
    const props = node[propsTag];
    if (!props || typeof props !== 'object') continue;
    const change = (props as XmlObject)[changeTag];
    if (!change || typeof change !== 'object') continue;
    const text = collectText(node, 'w:t');
    if (text === '') continue;
    out.push({
      scope,
      author: attrOf(change as XmlObject, 'w:author'),
      date: attrOf(change as XmlObject, 'w:date'),
      text,
      properties: changedProperties(propertiesOf((change as XmlObject)[propsTag], changeTag), propertiesOf(props, changeTag)),
    });
  }
  return out;
};

// A table carries its revisions at three levels, and Word records each the same
// way it records a run or a paragraph: the properties as they were BEFORE the
// edit, nested inside a `*Change` element under the current properties. So the
// same walker reads all five, and the scope is what tells a reader whether a
// reviewer restyled the whole table, one row, or one cell. Structural revisions
// (`w:cellIns`, `w:cellDel`, `w:cellMerge`) carry no before/after pair and are
// deliberately not reported here. Note the limit `attrSignature` imposes on all
// five: a property whose values hang off child elements rather than its own
// attributes (`w:tblBorders` is the common table case) compares as unchanged, so
// the revision is reported with its author and scope but names no property.
const extractFormatChanges = (root: unknown): ReadonlyArray<FormatChange> => [
  ...changesIn(root, 'w:r', 'w:rPr', 'w:rPrChange', 'run'),
  ...changesIn(root, 'w:p', 'w:pPr', 'w:pPrChange', 'paragraph'),
  ...changesIn(root, 'w:tbl', 'w:tblPr', 'w:tblPrChange', 'table'),
  ...changesIn(root, 'w:tr', 'w:trPr', 'w:trPrChange', 'row'),
  ...changesIn(root, 'w:tc', 'w:tcPr', 'w:tcPrChange', 'cell'),
];

const extractDocxMetadata = async (bytes: Uint8Array): Promise<Result<DocxMetadata, GraphError>> => {
  const zipR = await openOoxmlZip(bytes);
  if (!zipR.ok) return zipR;
  const zip = zipR.value;
  const documentXml = zip.read('word/document.xml');
  const document = parseXml(documentXml);
  const anchors = extractCommentAnchors(documentXml);
  const replacements = extractReplacements(documentXml);
  // Kind-prefixed so a deletion id can never mask an insertion id sharing it.
  const paired = new Set(replacements.flatMap((r) => [`del:${r.deletionId}`, `ins:${r.insertionId}`]));
  return ok({
    core: extractCoreProps(zip),
    app: extractAppProps(zip),
    custom: extractCustomProps(zip),
    people: extractPeople(parseXml(zip.read('word/people.xml'))),
    externalRels: extractExternalRels(zip),
    comments: extractComments(parseXml(zip.read('word/comments.xml')), anchors),
    insertions: extractTracked(document, 'w:ins').filter((t) => !paired.has(`ins:${t.id}`)),
    deletions: extractTracked(document, 'w:del').filter((t) => !paired.has(`del:${t.id}`)),
    replacements,
    moves: extractMoves(documentXml),
    formatChanges: extractFormatChanges(document),
    hiddenText: extractHidden(document),
    textBoxes: extractTextBoxes(zip),
    headersFooters: extractHeadersFooters(zip),
    fields: collectFields(zip),
    bookmarks: extractBookmarks(document),
    macros: extractMacros(zip),
  });
};

export { extractDocxMetadata };
export type { Bookmark, Comment, CustomProp, DocxMetadata, ExternalRel, Field, FormatChange, HeaderFooter, Move, Person, Replacement, TrackedChange };
