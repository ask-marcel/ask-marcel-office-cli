import type { OoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import { attrOf, collectText, findAll, findAllTexts, parseXml } from './ooxml-xml-walker.ts';

/**
 * Workbook annotations + the identities behind them. xlsx has two comment
 * formats: legacy cell comments (xl/comments*.xml, author by index into an
 * <authors> list) and modern threaded comments (xl/threadedComments/*, author
 * by personId resolved through xl/persons/person.xml). Both are user-authored
 * and invisible in the value-rendered body.
 */

type CellComment = { readonly cell: string; readonly author: string; readonly text: string };
type ThreadedComment = { readonly cell: string; readonly author: string; readonly date: string; readonly text: string };
type Person = { readonly id: string; readonly displayName: string; readonly userId: string };

const partsMatching = (zip: OoxmlZip, re: RegExp): ReadonlyArray<string> => zip.list().filter((p) => re.test(p));

const extractPeople = (zip: OoxmlZip): ReadonlyArray<Person> =>
  findAll(parseXml(zip.read('xl/persons/person.xml')), 'person').map((p) => ({
    id: attrOf(p, 'id'),
    displayName: attrOf(p, 'displayName'),
    userId: attrOf(p, 'userId'),
  }));

const legacyCommentsInPart = (root: unknown): ReadonlyArray<CellComment> => {
  const authors = findAllTexts(root, 'author');
  return findAll(root, 'comment').map((c) => ({ cell: attrOf(c, 'ref'), author: authors[Number(attrOf(c, 'authorId'))] ?? '', text: collectText(c, 't') }));
};

const extractLegacyComments = (zip: OoxmlZip): ReadonlyArray<CellComment> => {
  const out: Array<CellComment> = [];
  for (const path of partsMatching(zip, /^xl\/comments\d+\.xml$/)) out.push(...legacyCommentsInPart(parseXml(zip.read(path))));
  return out;
};

const threadedInPart = (root: unknown, nameById: Map<string, string>): ReadonlyArray<ThreadedComment> =>
  findAll(root, 'threadedComment').map((tc) => {
    const personId = attrOf(tc, 'personId');
    return { cell: attrOf(tc, 'ref'), author: nameById.get(personId) ?? personId, date: attrOf(tc, 'dT'), text: collectText(tc, 'text') };
  });

const extractThreadedComments = (zip: OoxmlZip, people: ReadonlyArray<Person>): ReadonlyArray<ThreadedComment> => {
  const nameById = new Map(people.map((p) => [p.id, p.displayName]));
  const out: Array<ThreadedComment> = [];
  for (const path of partsMatching(zip, /^xl\/threadedComments\/threadedComment\d+\.xml$/)) out.push(...threadedInPart(parseXml(zip.read(path)), nameById));
  return out;
};

export { extractLegacyComments, extractPeople, extractThreadedComments };
export type { CellComment, Person, ThreadedComment };
