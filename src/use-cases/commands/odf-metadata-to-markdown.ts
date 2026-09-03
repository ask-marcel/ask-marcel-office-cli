import type { OdfFormatChange, OdfMetadata, OdfReplacement, OdfTrackedChange, UserDefined } from './odf-metadata.ts';
import { renderBullets, renderKv, renderTable } from './ooxml-metadata-to-markdown.ts';

/**
 * Pure renderer: OdfMetadata → standalone `## OpenDocument metadata` document.
 * Uses the shared OOXML render primitives; empty sections emit `_(none)_`.
 */

const renderUserDefined = (props: ReadonlyArray<UserDefined>): string =>
  renderTable(
    props.map((p) => [p.name, p.value]),
    ['name', 'value']
  );

// Same section titles and columns as the docx block, so an agent reads one
// vocabulary across formats. ODF cannot name the properties a format change
// touched, so that table has no properties column.
const renderTracked = (changes: ReadonlyArray<OdfTrackedChange>): string =>
  renderTable(
    changes.map((t) => [t.id, t.author, t.date, t.text]),
    ['id', 'author', 'date', 'text']
  );

const renderReplacements = (replacements: ReadonlyArray<OdfReplacement>): string =>
  renderTable(
    replacements.map((r) => [r.deletionId, r.insertionId, r.author, r.date, r.before, r.after]),
    ['deletionId', 'insertionId', 'author', 'date', 'before', 'after']
  );

const renderFormatChanges = (changes: ReadonlyArray<OdfFormatChange>): string =>
  renderTable(
    changes.map((c) => [c.author, c.date, c.text]),
    ['author', 'date', 'text']
  );

const formatOdfMetadata = (meta: OdfMetadata): string => {
  const sections: ReadonlyArray<readonly [string, string]> = [
    ['Document properties', renderKv(meta.properties)],
    ['Keywords', renderBullets(meta.keywords)],
    ['User-defined properties', renderUserDefined(meta.userDefined)],
    ['Tracked changes — replacements', renderReplacements(meta.replacements)],
    ['Tracked changes — insertions', renderTracked(meta.insertions)],
    ['Tracked changes — deletions', renderTracked(meta.deletions)],
    ['Tracked changes — formatting', renderFormatChanges(meta.formatChanges)],
  ];
  const body = sections.map(([title, content]) => `### ${title}\n\n${content}`).join('\n\n');
  return `## OpenDocument metadata\n\n${body}\n`;
};

export { formatOdfMetadata };
