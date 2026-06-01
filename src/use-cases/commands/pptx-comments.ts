import { posix } from 'node:path';
import type { OoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import { attrOf, collectText, findAll, parseXml } from './ooxml-xml-walker.ts';
import type { XmlObject } from './ooxml-xml-walker.ts';

/**
 * PowerPoint comments come in two formats: legacy (`ppt/commentAuthors.xml`
 * authors by integer id + `ppt/comments/comment*.xml` `<p:cm authorId dt>`
 * with `<p:text>` body) and modern (`ppt/authors.xml` authors by GUID +
 * `ppt/comments/*.xml` `<p188:cm authorId created>` with DrawingML `<a:t>`
 * body). Both are scanned; authors are resolved by id in either scheme.
 * Modern-comment support is best-effort against the p188 (2018/8) schema.
 */

type CommentAuthor = { readonly id: string; readonly name: string; readonly initials: string };
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

const toAuthor = (node: XmlObject): CommentAuthor => ({ id: attrOf(node, 'id'), name: attrOf(node, 'name'), initials: attrOf(node, 'initials') });

const extractCommentAuthors = (zip: OoxmlZip): ReadonlyArray<CommentAuthor> => {
  const legacy = findAll(parseXml(zip.read('ppt/commentAuthors.xml')), 'p:cmAuthor');
  const modern = findAll(parseXml(zip.read('ppt/authors.xml')), 'p188:author');
  return [...legacy, ...modern].map(toAuthor);
};

const commentsInPart = (root: unknown, nameById: Map<string, string>): ReadonlyArray<PptxComment> => {
  const resolve = (id: string): string => nameById.get(id) ?? id;
  const legacy = findAll(root, 'p:cm').map((cm) => ({ author: resolve(attrOf(cm, 'authorId')), date: attrOf(cm, 'dt'), text: collectText(cm, 'p:text') }));
  const modern = findAll(root, 'p188:cm').map((cm) => ({ author: resolve(attrOf(cm, 'authorId')), date: attrOf(cm, 'created'), text: collectText(cm, 'a:t') }));
  return [...legacy, ...modern];
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
