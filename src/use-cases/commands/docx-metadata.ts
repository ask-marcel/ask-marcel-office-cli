import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { OoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import { openOoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { extractAppProps, extractCoreProps, extractCustomProps, extractExternalRels, extractMacros } from './ooxml-metadata.ts';
import type { CustomProp, ExternalRel } from './ooxml-metadata.ts';
import { attrOf, collectText, findAll, findAllTexts, parseXml } from './ooxml-xml-walker.ts';
import type { XmlObject } from './ooxml-xml-walker.ts';

/**
 * Pulls the side-channel content out of a .docx zip — every text-bearing
 * surface mammoth drops on the floor. High-value subset (10 sections, see
 * plan): core / app / custom doc properties, people registry, external
 * hyperlinks, comments, tracked changes (ins + del), hidden text (w:vanish),
 * field instructions (MERGEFIELD / HYPERLINK / DOCVARIABLE), bookmarks.
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
type Comment = { readonly id: string; readonly author: string; readonly initials: string; readonly date: string; readonly text: string };
type TrackedChange = { readonly id: string; readonly author: string; readonly date: string; readonly text: string };
type Field = { readonly source: string; readonly instruction: string };
type Bookmark = { readonly id: string; readonly name: string };

type DocxMetadata = {
  readonly core: CoreProps;
  readonly app: AppProps;
  readonly custom: ReadonlyArray<CustomProp>;
  readonly people: ReadonlyArray<Person>;
  readonly externalRels: ReadonlyArray<ExternalRel>;
  readonly comments: ReadonlyArray<Comment>;
  readonly insertions: ReadonlyArray<TrackedChange>;
  readonly deletions: ReadonlyArray<TrackedChange>;
  readonly hiddenText: ReadonlyArray<string>;
  readonly fields: ReadonlyArray<Field>;
  readonly bookmarks: ReadonlyArray<Bookmark>;
  readonly macros: ReadonlyArray<string>;
};

const extractPeople = (root: unknown): ReadonlyArray<Person> => {
  const persons = findAll(root, 'w15:person');
  return persons.map((p) => {
    const presence = findAll(p, 'w15:presenceInfo')[0] ?? ({} as XmlObject);
    return {
      author: attrOf(p, 'w15:author'),
      providerId: attrOf(presence, 'w15:providerId'),
      userId: attrOf(presence, 'w15:userId'),
    };
  });
};

const extractComments = (root: unknown): ReadonlyArray<Comment> => {
  const comments = findAll(root, 'w:comment');
  return comments.map((c) => ({
    id: attrOf(c, 'w:id'),
    author: attrOf(c, 'w:author'),
    initials: attrOf(c, 'w:initials'),
    date: attrOf(c, 'w:date'),
    text: collectText(c, 'w:t'),
  }));
};

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
    if (!Object.hasOwn(rPr as Record<string, unknown>, 'w:vanish')) continue;
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

const collectFields = (zip: OoxmlZip): ReadonlyArray<Field> => {
  const candidates = ['word/document.xml', 'word/header1.xml', 'word/header2.xml', 'word/header3.xml', 'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml'];
  const out: Array<Field> = [];
  for (const path of candidates) {
    const parsed = parseXml(zip.read(path));
    if (parsed === undefined) continue;
    for (const f of extractFieldsFromOne(parsed, path)) out.push(f);
  }
  return out;
};

const extractDocxMetadata = async (bytes: Uint8Array): Promise<Result<DocxMetadata, GraphError>> => {
  const zipR = await openOoxmlZip(bytes);
  if (!zipR.ok) return zipR;
  const zip = zipR.value;
  const document = parseXml(zip.read('word/document.xml'));
  return ok({
    core: extractCoreProps(zip),
    app: extractAppProps(zip),
    custom: extractCustomProps(zip),
    people: extractPeople(parseXml(zip.read('word/people.xml'))),
    externalRels: extractExternalRels(zip),
    comments: extractComments(parseXml(zip.read('word/comments.xml'))),
    insertions: extractTracked(document, 'w:ins'),
    deletions: extractTracked(document, 'w:del'),
    hiddenText: extractHidden(document),
    fields: collectFields(zip),
    bookmarks: extractBookmarks(document),
    macros: extractMacros(zip),
  });
};

export { extractDocxMetadata };
export type { Bookmark, Comment, CustomProp, DocxMetadata, ExternalRel, Field, Person, TrackedChange };
