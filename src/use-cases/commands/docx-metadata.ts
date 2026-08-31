import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { OoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import { openOoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { extractAppProps, extractCoreProps, extractCustomProps, extractExternalRels, extractMacros } from './ooxml-metadata.ts';
import type { CustomProp, ExternalRel } from './ooxml-metadata.ts';
import { extractCommentAnchors } from './docx-comment-anchors.ts';
import type { OrderedNode } from './ooxml-xml-walker.ts';
import { attrOf, collectOrderedText, collectText, findAll, findAllTexts, orderedAttrOf, orderedSiblingGroups, parseXml, parseXmlOrdered } from './ooxml-xml-walker.ts';

/**
 * Pulls the side-channel content out of a .docx zip — every text-bearing
 * surface mammoth drops on the floor: core / app / custom doc properties,
 * people registry, external hyperlinks, comments, tracked changes (replacements,
 * then the insertions and deletions that pair with nothing),
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
    hiddenText: extractHidden(document),
    textBoxes: extractTextBoxes(zip),
    headersFooters: extractHeadersFooters(zip),
    fields: collectFields(zip),
    bookmarks: extractBookmarks(document),
    macros: extractMacros(zip),
  });
};

export { extractDocxMetadata };
export type { Bookmark, Comment, CustomProp, DocxMetadata, ExternalRel, Field, HeaderFooter, Person, Replacement, TrackedChange };
