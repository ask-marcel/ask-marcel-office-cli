import { posix } from 'node:path';
import type { OoxmlZip } from '../../infra/ooxml-zip-adapter.ts';
import { attrOf, collectText, findAll, findAllTexts, parseXml } from './ooxml-xml-walker.ts';

/**
 * Workbook annotations + the identities behind them. xlsx has two comment
 * formats: legacy cell comments (xl/comments*.xml, author by index into an
 * <authors> list) and modern threaded comments (xl/threadedComments/*, author
 * by personId resolved through xl/persons/person.xml). Both are user-authored
 * and invisible in the value-rendered body. A comment carries its sheet when the
 * package relationships name it (the workbook points at each sheet part, each
 * sheet part at its comments), which is how Excel writes every workbook.
 */

type CellComment = { readonly cell: string; readonly author: string; readonly text: string; readonly sheet?: string };
type ThreadedComment = { readonly cell: string; readonly author: string; readonly date: string; readonly text: string; readonly sheet?: string };
type Person = { readonly id: string; readonly displayName: string; readonly userId: string };

const partsMatching = (zip: OoxmlZip, re: RegExp): ReadonlyArray<string> => zip.list().filter((p) => re.test(p));

// A relationship target is package-absolute when it starts with `/`, else relative
// to the directory of the part that owns the relationship.
const resolveTarget = (ownerDir: string, target: string): string => (target.startsWith('/') ? target.slice(1) : posix.normalize(posix.join(ownerDir, target)));

const isCommentsRelationship = (type: string): boolean => type.endsWith('/comments') || type.endsWith('/threadedComment');

// Comments part (legacy or threaded) → the name of the sheet it annotates.
const commentPartToSheet = (zip: OoxmlZip): ReadonlyMap<string, string> => {
  const workbookRels = findAll(parseXml(zip.read('xl/_rels/workbook.xml.rels')), 'Relationship');
  const sheetPathById = new Map(workbookRels.map((rel) => [attrOf(rel, 'Id'), resolveTarget('xl', attrOf(rel, 'Target'))]));
  const map = new Map<string, string>();
  for (const sheet of findAll(parseXml(zip.read('xl/workbook.xml')), 'sheet')) {
    const sheetPath = sheetPathById.get(attrOf(sheet, 'r:id'));
    if (sheetPath === undefined) continue;
    const dir = posix.dirname(sheetPath);
    for (const rel of findAll(parseXml(zip.read(`${dir}/_rels/${posix.basename(sheetPath)}.rels`)), 'Relationship')) {
      if (isCommentsRelationship(attrOf(rel, 'Type'))) map.set(resolveTarget(dir, attrOf(rel, 'Target')), attrOf(sheet, 'name'));
    }
  }
  return map;
};

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
  const partToSheet = commentPartToSheet(zip);
  const out: Array<CellComment> = [];
  for (const path of partsMatching(zip, /^xl\/comments\d+\.xml$/)) out.push(...legacyCommentsInPart(parseXml(zip.read(path))).map((c) => ({ ...c, sheet: partToSheet.get(path) })));
  return out;
};

const threadedInPart = (root: unknown, nameById: Map<string, string>): ReadonlyArray<ThreadedComment> =>
  findAll(root, 'threadedComment').map((tc) => {
    const personId = attrOf(tc, 'personId');
    return { cell: attrOf(tc, 'ref'), author: nameById.get(personId) ?? personId, date: attrOf(tc, 'dT'), text: collectText(tc, 'text') };
  });

const extractThreadedComments = (zip: OoxmlZip, people: ReadonlyArray<Person>): ReadonlyArray<ThreadedComment> => {
  const nameById = new Map(people.map((p) => [p.id, p.displayName]));
  const partToSheet = commentPartToSheet(zip);
  const out: Array<ThreadedComment> = [];
  for (const path of partsMatching(zip, /^xl\/threadedComments\/threadedComment\d+\.xml$/))
    out.push(...threadedInPart(parseXml(zip.read(path)), nameById).map((c) => ({ ...c, sheet: partToSheet.get(path) })));
  return out;
};

export { extractLegacyComments, extractPeople, extractThreadedComments };
export type { CellComment, Person, ThreadedComment };
