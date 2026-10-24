/**
 * Renders the `## OneNote metadata` block appended to `get-onenote-page-as-markdown`
 * output under `--include-metadata true`. Mirrors the docx/xlsx/odf metadata-append
 * pattern: a fenced section listing the page's title, created / last-modified
 * timestamps, and the parent section + notebook names (expanded by the command).
 * Pure formatter — every field is optional and omitted when absent.
 */

type OnenotePage = {
  readonly title?: string;
  readonly createdDateTime?: string;
  readonly lastModifiedDateTime?: string;
  readonly parentSection?: { readonly displayName?: string };
  readonly parentNotebook?: { readonly displayName?: string };
};

const nonEmpty = (v: unknown): v is string => typeof v === 'string' && v !== '';

const formatOnenoteMetadata = (page: OnenotePage): string => {
  const lines: string[] = ['## OneNote metadata'];
  if (nonEmpty(page.title)) lines.push(`- **Title:** ${page.title}`);
  if (nonEmpty(page.createdDateTime)) lines.push(`- **Created:** ${page.createdDateTime}`);
  if (nonEmpty(page.lastModifiedDateTime)) lines.push(`- **Last modified:** ${page.lastModifiedDateTime}`);
  // Resolve the expanded names once so the value is reused in the push (no
  // redundant optional chain that would be an unkillable equivalent mutant).
  const notebook = page.parentNotebook?.displayName;
  if (nonEmpty(notebook)) lines.push(`- **Notebook:** ${notebook}`);
  const section = page.parentSection?.displayName;
  if (nonEmpty(section)) lines.push(`- **Section:** ${section}`);
  return lines.join('\n');
};

export { formatOnenoteMetadata };
export type { OnenotePage };
