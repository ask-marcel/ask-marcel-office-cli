import type { Bookmark, Comment, CustomProp, DocxMetadata, ExternalRel, Field, FormatChange, HeaderFooter, Move, Person, Replacement, TrackedChange } from './docx-metadata.ts';
import { escapeCell, NONE, renderBullets, renderKv, renderMacros, renderTable } from './ooxml-metadata-to-markdown.ts';

/**
 * Pure renderer: DocxMetadata → markdown block to append after the
 * mammoth-converted body. Uses the shared OOXML render primitives
 * (escapeCell / renderKv / renderTable / renderBullets / NONE) and assembles
 * the docx-specific section order on top.
 */

const renderCustom = (props: ReadonlyArray<CustomProp>): string =>
  renderTable(
    props.map((p) => [p.name, p.value]),
    ['name', 'value']
  );

const renderPeople = (people: ReadonlyArray<Person>): string =>
  renderTable(
    people.map((p) => [p.author, p.providerId, p.userId]),
    ['author', 'providerId', 'userId']
  );

const renderRels = (rels: ReadonlyArray<ExternalRel>): string =>
  renderTable(
    rels.map((r) => [r.source, r.type, r.target]),
    ['source', 'type', 'target']
  );

const renderComments = (comments: ReadonlyArray<Comment>): string => {
  if (comments.length === 0) return NONE;
  return comments
    .map((c) => {
      const head = `- **id ${escapeCell(c.id)}** — ${escapeCell(c.author)} (${escapeCell(c.initials)}) @ ${escapeCell(c.date)}`;
      const on = c.anchor === undefined ? '' : ` on “${escapeCell(c.anchor)}”`;
      return `${head}${on}\n  > ${escapeCell(c.text)}`;
    })
    .join('\n');
};

const renderTracked = (changes: ReadonlyArray<TrackedChange>): string =>
  renderTable(
    changes.map((t) => [t.id, t.author, t.date, t.text]),
    ['id', 'author', 'date', 'text']
  );

// `before` and `after` rather than deleted/inserted: the pair is one edit, and
// naming the halves after the revision elements they came from would invite the
// same "two separate changes" reading that pairing exists to remove.
const renderReplacements = (replacements: ReadonlyArray<Replacement>): string =>
  renderTable(
    replacements.map((r) => [r.deletionId, r.insertionId, r.author, r.date, r.before, r.after]),
    ['deletionId', 'insertionId', 'author', 'date', 'before', 'after']
  );

const renderMoves = (moves: ReadonlyArray<Move>): string =>
  renderTable(
    moves.map((m) => [m.name, m.author, m.date, m.text, m.halves]),
    ['name', 'author', 'date', 'text', 'halves']
  );

// Properties joined rather than one row each: the edit is "these three things
// changed on this run", and splitting it into three rows would report one edit
// as three, which is the reading this whole section exists to prevent.
const renderFormatChanges = (changes: ReadonlyArray<FormatChange>): string =>
  renderTable(
    changes.map((c) => [c.scope, c.author, c.date, c.text, c.properties.join(', ')]),
    ['scope', 'author', 'date', 'text', 'properties']
  );

const renderFields = (fields: ReadonlyArray<Field>): string =>
  renderTable(
    fields.map((f) => [f.source, f.instruction]),
    ['source', 'instruction']
  );

const renderBookmarks = (bookmarks: ReadonlyArray<Bookmark>): string =>
  renderTable(
    bookmarks.map((b) => [b.id, b.name]),
    ['id', 'name']
  );

const renderHeadersFooters = (parts: ReadonlyArray<HeaderFooter>): string =>
  renderTable(
    parts.map((p) => [p.part, p.text]),
    ['part', 'text']
  );

const formatDocxMetadata = (meta: DocxMetadata): string => {
  const sections: ReadonlyArray<readonly [string, string]> = [
    ['Core properties', renderKv(meta.core)],
    ['Application properties', renderKv(meta.app)],
    ['Custom properties', renderCustom(meta.custom)],
    ['People registry', renderPeople(meta.people)],
    ['External relationships', renderRels(meta.externalRels)],
    ['Comments', renderComments(meta.comments)],
    ['Tracked changes — replacements', renderReplacements(meta.replacements)],
    ['Tracked changes — insertions', renderTracked(meta.insertions)],
    ['Tracked changes — deletions', renderTracked(meta.deletions)],
    ['Tracked changes — moves', renderMoves(meta.moves)],
    ['Tracked changes — formatting', renderFormatChanges(meta.formatChanges)],
    ['Hidden-formatted text (w:vanish)', renderBullets(meta.hiddenText)],
    ['Text boxes / shapes', renderBullets(meta.textBoxes)],
    ['Headers & footers', renderHeadersFooters(meta.headersFooters)],
    ['Fields (MERGEFIELD / HYPERLINK / DOCVARIABLE)', renderFields(meta.fields)],
    ['Bookmarks', renderBookmarks(meta.bookmarks)],
    ['Macros (VBA)', renderMacros(meta.macros)],
  ];
  const body = sections.map(([title, content]) => `### ${title}\n\n${content}`).join('\n\n');
  return `---\n\n## DOCX metadata\n\n${body}\n`;
};

export { formatDocxMetadata };
