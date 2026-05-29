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
type PptxComment = { readonly author: string; readonly date: string; readonly text: string };

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
  const out: Array<PptxComment> = [];
  for (const path of zip.list().filter((p) => /^ppt\/comments\/.*\.xml$/.test(p))) out.push(...commentsInPart(parseXml(zip.read(path)), nameById));
  return out;
};

export { extractCommentAuthors, extractComments };
export type { CommentAuthor, PptxComment };
