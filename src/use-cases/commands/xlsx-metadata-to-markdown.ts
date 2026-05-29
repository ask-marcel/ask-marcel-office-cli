import type { CustomProp, ExternalRel } from './ooxml-metadata.ts';
import { renderKv, renderTable } from './ooxml-metadata-to-markdown.ts';
import type { CellComment, DefinedName, Person, Sheet, ThreadedComment, XlsxMetadata } from './xlsx-metadata.ts';

/**
 * Pure renderer: XlsxMetadata → `## Workbook metadata` block appended after
 * the sheet tables. Uses the shared OOXML render primitives and assembles the
 * workbook-specific section order on top. Empty sections emit `_(none)_`.
 */

const renderCustom = (props: ReadonlyArray<CustomProp>): string =>
  renderTable(
    props.map((p) => [p.name, p.value]),
    ['name', 'value']
  );

const renderRels = (rels: ReadonlyArray<ExternalRel>): string =>
  renderTable(
    rels.map((r) => [r.source, r.type, r.target]),
    ['source', 'type', 'target']
  );

const renderDefinedNames = (names: ReadonlyArray<DefinedName>): string =>
  renderTable(
    names.map((d) => [d.name, d.refersTo, d.hidden ? 'hidden' : '']),
    ['name', 'refersTo', 'hidden']
  );

const renderSheets = (sheets: ReadonlyArray<Sheet>): string =>
  renderTable(
    sheets.map((s) => [s.name, s.state]),
    ['name', 'state']
  );

const renderComments = (comments: ReadonlyArray<CellComment>): string =>
  renderTable(
    comments.map((c) => [c.cell, c.author, c.text]),
    ['cell', 'author', 'text']
  );

const renderThreaded = (comments: ReadonlyArray<ThreadedComment>): string =>
  renderTable(
    comments.map((c) => [c.cell, c.author, c.date, c.text]),
    ['cell', 'author', 'date', 'text']
  );

const renderPeople = (people: ReadonlyArray<Person>): string =>
  renderTable(
    people.map((p) => [p.id, p.displayName, p.userId]),
    ['id', 'displayName', 'userId']
  );

const formatXlsxMetadata = (meta: XlsxMetadata): string => {
  const sections: ReadonlyArray<readonly [string, string]> = [
    ['Core properties', renderKv(meta.core)],
    ['Application properties', renderKv(meta.app)],
    ['Custom properties', renderCustom(meta.custom)],
    ['External relationships', renderRels(meta.externalRels)],
    ['Defined names', renderDefinedNames(meta.definedNames)],
    ['Hidden / very-hidden sheets', renderSheets(meta.hiddenSheets)],
    ['Cell comments', renderComments(meta.comments)],
    ['Threaded comments', renderThreaded(meta.threadedComments)],
    ['People', renderPeople(meta.people)],
  ];
  const body = sections.map(([title, content]) => `### ${title}\n\n${content}`).join('\n\n');
  return `---\n\n## Workbook metadata\n\n${body}\n`;
};

export { formatXlsxMetadata };
