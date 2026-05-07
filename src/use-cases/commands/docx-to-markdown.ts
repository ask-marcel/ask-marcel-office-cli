import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { mammothToHtml } from '../../infra/mammoth-adapter.ts';
import { htmlToMarkdown } from '../../infra/turndown-adapter.ts';

type MarkdownEnvelope = { readonly contentType: 'text/markdown'; readonly size: number; readonly text: string };

/**
 * Mammoth emits docx tables as `<table><tr>...</tr>...</table>` with no
 * `<thead>` (and often no `<tbody>` wrapper at all). Turndown's GFM
 * plugin only converts tables that have a `<thead>`, so promote the
 * first `<tr>` of each table into a `<thead>` before turndown sees it.
 *
 * Implemented as a linear-time string walker rather than a regex with
 * adjacent non-greedy quantifiers, which would be flagged by SonarJS's
 * slow-regex rule for backtracking risk.
 */
const TABLE_OPEN = '<table>';
const TABLE_CLOSE = '</table>';
const TBODY_OPEN = '<tbody>';
const TBODY_CLOSE = '</tbody>';
const TR_CLOSE = '</tr>';

const transformOneTable = (table: string): string => {
  let inner = table.slice(TABLE_OPEN.length, table.length - TABLE_CLOSE.length);
  if (inner.startsWith(TBODY_OPEN)) inner = inner.slice(TBODY_OPEN.length);
  if (inner.endsWith(TBODY_CLOSE)) inner = inner.slice(0, -TBODY_CLOSE.length);
  const firstRowEnd = inner.indexOf(TR_CLOSE);
  if (firstRowEnd === -1) return table;
  const firstRow = inner.slice(0, firstRowEnd + TR_CLOSE.length);
  const rest = inner.slice(firstRowEnd + TR_CLOSE.length);
  return `<table><thead>${firstRow}</thead><tbody>${rest}</tbody></table>`;
};

const promoteFirstRowToThead = (html: string): string => {
  let out = '';
  let cursor = 0;
  while (cursor < html.length) {
    const tableStart = html.indexOf(TABLE_OPEN, cursor);
    if (tableStart === -1) {
      out += html.slice(cursor);
      return out;
    }
    out += html.slice(cursor, tableStart);
    const tableEnd = html.indexOf(TABLE_CLOSE, tableStart);
    if (tableEnd === -1) {
      out += html.slice(tableStart);
      return out;
    }
    const fullTable = html.slice(tableStart, tableEnd + TABLE_CLOSE.length);
    out += transformOneTable(fullTable);
    cursor = tableEnd + TABLE_CLOSE.length;
  }
  return out;
};

const docxToMarkdown = async (bytes: Uint8Array): Promise<Result<MarkdownEnvelope, GraphError>> => {
  const html = await mammothToHtml(bytes);
  if (!html.ok) return html;
  const md = htmlToMarkdown(promoteFirstRowToThead(html.value));
  if (!md.ok) return md;
  return ok({ contentType: 'text/markdown', size: md.value.length, text: md.value });
};

export { docxToMarkdown, promoteFirstRowToThead };
export type { MarkdownEnvelope };
