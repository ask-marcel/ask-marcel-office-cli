import { posix } from 'node:path';
import type { OoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import { attrOf, collectTextLocal, findAll, findAllLocal, parseXml } from './ooxml-xml-walker.ts';
import type { XmlObject } from './ooxml-xml-walker.ts';

/**
 * PowerPoint comments come in two formats: legacy (`ppt/commentAuthors.xml`
 * authors by integer id + `ppt/comments/comment*.xml` `<cm authorId dt>` with a
 * `<text>` body) and modern (`ppt/authors.xml` authors by GUID +
 * `ppt/comments/*.xml` `<cm authorId created>` with a DrawingML `<t>` body).
 * Elements are matched by local name, so the prefixes a writer happens to bind
 * (`p:` / `p188:` / `a:` from PowerPoint, anything else from a third-party tool)
 * never decide whether a comment is found. Both generations share the local name
 * `cm`, so one pass reads them all; authors are resolved by id in either scheme.
 */

type CommentAuthor = { readonly id: string; readonly name: string; readonly initials: string; readonly email: string };
type PptxComment = { readonly author: string; readonly date: string; readonly text: string; readonly slide?: string };

const SLIDE_RE = /^ppt\/slides\/slide\d+\.xml$/;

// Map each comment part (ppt/comments/*.xml) to the slide that references it.
// A slide's `_rels` carries a `…/comments` relationship to its comment part —
// the same rels mechanism pptx-slides uses for speaker notes. Comments whose
// part isn't referenced by any slide stay unanchored.
const commentPartToSlide = (zip: OoxmlZip): ReadonlyMap<string, string> => {
  const map = new Map<string, string>();
  for (const slidePath of zip.list().filter((p) => SLIDE_RE.test(p))) {
    const relsPath = `${posix.dirname(slidePath)}/_rels/${posix.basename(slidePath)}.rels`;
    for (const rel of findAll(parseXml(zip.read(relsPath)), 'Relationship')) {
      if (!attrOf(rel, 'Type').endsWith('comments')) continue;
      const target = posix.normalize(posix.join(posix.dirname(slidePath), attrOf(rel, 'Target')));
      map.set(target, posix.basename(slidePath));
    }
  }
  return map;
};

// The commenter's directory identity. A modern <author> carries it on the element;
// a legacy <cmAuthor> carries it only inside a <presenceInfo> extension nested below.
// PowerPoint writes `S::<UPN>::<Entra object id>`, a third-party writer may store the
// address bare, and a non-AD provider stamps a display name that is no address at all.
const emailOf = (node: XmlObject): string => {
  const raw = attrOf(node, 'userId') || attrOf(findAllLocal(node, 'presenceInfo')[0] ?? {}, 'userId');
  const parts = raw.split('::');
  const candidate = parts.length === 3 ? (parts[1] ?? '') : raw;
  return candidate.includes('@') ? candidate : '';
};

const toAuthor = (node: XmlObject): CommentAuthor => ({ id: attrOf(node, 'id'), name: attrOf(node, 'name'), initials: attrOf(node, 'initials'), email: emailOf(node) });

const extractCommentAuthors = (zip: OoxmlZip): ReadonlyArray<CommentAuthor> => {
  const legacy = findAllLocal(parseXml(zip.read('ppt/commentAuthors.xml')), 'cmAuthor');
  const modern = findAllLocal(parseXml(zip.read('ppt/authors.xml')), 'author');
  return [...legacy, ...modern].map(toAuthor);
};

// `dt` / `<text>` are the legacy spellings, `created` / `<t>` the modern ones; a
// given comment carries one pair, so falling through picks the one that is there.
const commentsInPart = (root: unknown, nameById: Map<string, string>): ReadonlyArray<PptxComment> => {
  const resolve = (id: string): string => nameById.get(id) ?? id;
  return findAllLocal(root, 'cm').map((cm) => ({
    author: resolve(attrOf(cm, 'authorId')),
    date: attrOf(cm, 'dt') || attrOf(cm, 'created'),
    text: collectTextLocal(cm, 'text') || collectTextLocal(cm, 't'),
  }));
};

const extractComments = (zip: OoxmlZip, authors: ReadonlyArray<CommentAuthor>): ReadonlyArray<PptxComment> => {
  const nameById = new Map(authors.map((a) => [a.id, a.name]));
  const partToSlide = commentPartToSlide(zip);
  const out: Array<PptxComment> = [];
  for (const path of zip.list().filter((p) => /^ppt\/comments\/.*\.xml$/.test(p))) {
    const slide = partToSlide.get(path);
    for (const comment of commentsInPart(parseXml(zip.read(path)), nameById)) out.push(slide === undefined ? comment : { ...comment, slide });
  }
  return out;
};

export { extractCommentAuthors, extractComments };
export type { CommentAuthor, PptxComment };
