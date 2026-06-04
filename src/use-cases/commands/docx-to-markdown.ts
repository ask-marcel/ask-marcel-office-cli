import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { mammothToHtml } from '../../infra/mammoth-adapter.ts';
import { htmlToMarkdown } from '../../infra/turndown-adapter.ts';
import { extractDocxMetadata } from './docx-metadata.ts';
import { formatDocxMetadata } from './docx-metadata-to-markdown.ts';

type MarkdownEnvelope = { readonly contentType: 'text/markdown'; readonly size: number; readonly text: string };
type DocxToMarkdownOptions = { readonly includeMetadata?: boolean; readonly inlineImages?: boolean };

// By default the docx's images are NOT embedded — mammoth inlines each as a huge
// base64 `data:` URI, which bloats the markdown and duplicates what
// `extract-drive-item-images` already returns (full-resolution originals as
// files). Replace every such image with an `[image: <alt>]` placeholder so its
// position in the text survives. `inlineImages: true` keeps the base64. The
// alt/data captures are bounded character classes (no catastrophic backtracking).
const stripInlineImages = (markdown: string): string =>
  markdown.replaceAll(/!\[([^\]]*)\]\(data:[^)]*\)/g, (_match, alt: string) => (alt.length > 0 ? `[image: ${alt}]` : '[image]'));

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

const docxToMarkdown = async (bytes: Uint8Array, opts: DocxToMarkdownOptions = {}): Promise<Result<MarkdownEnvelope, GraphError>> => {
  const html = await mammothToHtml(bytes);
  if (!html.ok) return html;
  const md = htmlToMarkdown(promoteFirstRowToThead(html.value));
  if (!md.ok) return md;
  let text = opts.inlineImages === true ? md.value : stripInlineImages(md.value);
  if (opts.includeMetadata === true) {
    const meta = await extractDocxMetadata(bytes);
    if (!meta.ok) return meta;
    text = `${text}\n\n${formatDocxMetadata(meta.value)}`;
  }
  // size = UTF-8 byte count (audit §2.1); `.length` is UTF-16 code units.
  return ok({ contentType: 'text/markdown', size: new TextEncoder().encode(text).byteLength, text });
};

export { docxToMarkdown, promoteFirstRowToThead, stripInlineImages };
export type { DocxToMarkdownOptions, MarkdownEnvelope };
