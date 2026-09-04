import type { OdfMetadata, UserDefined } from './odf-metadata.ts';
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

const formatOdfMetadata = (meta: OdfMetadata): string => {
  const sections: ReadonlyArray<readonly [string, string]> = [
    ['Document properties', renderKv(meta.properties)],
    ['Keywords', renderBullets(meta.keywords)],
    ['User-defined properties', renderUserDefined(meta.userDefined)],
  ];
  const body = sections.map(([title, content]) => `### ${title}\n\n${content}`).join('\n\n');
  return `## OpenDocument metadata\n\n${body}\n`;
};

export { formatOdfMetadata };
