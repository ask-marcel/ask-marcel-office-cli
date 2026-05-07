import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { docxToMarkdown } from './docx-to-markdown.ts';
import { convertToMarkdown } from './markdown-pipeline.ts';
import { isPlainTextFilename } from './text-passthrough.ts';
import { csvToMarkdownTable, xlsxToMarkdown } from './xlsx-to-markdown.ts';

/**
 * Strategy dispatcher for `*-as-markdown` commands. Picks the right
 * conversion path based on the file extension and routes through it:
 *
 * - plain-text (txt/md/json/yaml/etc.)  → raw bytes envelope, no conversion
 * - csv                                 → csvToMarkdownTable (markdown table)
 * - docx                                → mammoth → turndown
 * - xlsx                                → sheetjs → markdown table per sheet
 * - loop / fluid / wbtx / whiteboard    → existing Graph ?format=html
 *                                         pipeline (the only inputs Graph
 *                                         HTML conversion actually accepts)
 * - everything else (pptx, pdf, rtf, odt, …) → err pointing at the
 *                                              corresponding *-as-pdf
 *                                              command (38 input extensions)
 */

const HTML_FORMAT_INPUTS: ReadonlySet<string> = new Set(['loop', 'fluid', 'wbtx', 'whiteboard']);

const PPTX_HINT =
  'pptx not supported by `*-as-markdown`. Use the corresponding `*-as-pdf` command — Graph PDF conversion preserves slide layout, and a vision-capable LLM reads it more reliably than flattened slide-by-slide bullets.';

const GENERIC_HINT = (ext: string): string =>
  `${ext} not supported by \`*-as-markdown\`. Use the corresponding \`*-as-pdf\` command — Graph \`?format=pdf\` accepts 38 input extensions including this one.`;

const extensionOf = (filename: string): string => {
  const dot = filename.lastIndexOf('.');
  if (dot === -1 || dot === filename.length - 1) return '';
  return filename.slice(dot + 1).toLowerCase();
};

type FetchOptions = { readonly elevated?: boolean };

/**
 * Real Graph responses for `/drives/{id}/items/{id}/content` are 302
 * redirects to a CDN URL. `getBinary` captures that and returns
 * `{ '@microsoft.graph.downloadUrl': '...' }` — NOT inline bytes.
 * To actually get the bytes we have to follow the URL via `fetchUrl`.
 *
 * The historical-version commands pass `elevated: true` so that the
 * initial Graph call is signed with an M365ChatClient token (on the
 * ODSP `logicalPermissions` allow-list); without that, Graph's 302
 * redirects to a streamContent URL whose embedded tempauth is signed
 * by Teams web client identity and rejected by SharePoint with 403.
 */
const fetchRawBytes = async (graph: GraphClient, contentPath: string, opts: FetchOptions = {}): Promise<Result<Uint8Array, GraphError>> => {
  const initial = opts.elevated ? await graph.getBinaryElevated(contentPath) : await graph.getBinary(contentPath);
  if (!initial.ok) return initial;
  const value = initial.value as Record<string, unknown>;

  const downloadUrl = value['@microsoft.graph.downloadUrl'];
  if (typeof downloadUrl === 'string') {
    const followed = await graph.fetchUrl(downloadUrl);
    if (!followed.ok) return followed;
    return decodeBlobBytes(followed.value as Record<string, unknown>);
  }
  return decodeBlobBytes(value);
};

const decodeBlobBytes = (blob: Record<string, unknown>): Result<Uint8Array, GraphError> => {
  const b64 = blob['base64'];
  if (typeof b64 === 'string') {
    const binary = atob(b64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return ok(bytes);
  }
  const text = blob['text'];
  if (typeof text === 'string') {
    return ok(new TextEncoder().encode(text));
  }
  return err({ type: 'api_error', status: 500, message: 'unexpected envelope: response had no @microsoft.graph.downloadUrl, no base64 bytes, and no text body' });
};

const officeToMarkdown = async (graph: GraphClient, contentPath: string, filename: string, opts: FetchOptions = {}): Promise<Result<unknown, GraphError>> => {
  if (isPlainTextFilename(filename)) {
    // Plain-text passthrough: follow the CDN redirect (just like csv/docx/xlsx
    // do) and return the bytes inline as `{ contentType: "text/plain", size,
    // text }` instead of the raw `{ "@microsoft.graph.downloadUrl": "..." }`
    // envelope. Audit §1.11: an LLM consuming the JSON shouldn't need a
    // separate fetch tool to actually read a txt/md/html body.
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    const text = new TextDecoder().decode(bytes.value);
    return ok({ contentType: 'text/plain', size: bytes.value.byteLength, text });
  }
  const ext = extensionOf(filename);

  if (ext === 'csv') {
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    const csv = new TextDecoder().decode(bytes.value);
    const md = csvToMarkdownTable(csv);
    return ok({ contentType: 'text/markdown', size: md.length, text: md });
  }

  if (ext === 'docx') {
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    return docxToMarkdown(bytes.value);
  }

  if (ext === 'xlsx') {
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    return xlsxToMarkdown(bytes.value);
  }

  if (HTML_FORMAT_INPUTS.has(ext)) {
    return convertToMarkdown(graph, `${contentPath}?format=html`);
  }

  if (ext === 'pptx') {
    return err({ type: 'api_error', status: 415, message: PPTX_HINT });
  }

  return err({ type: 'api_error', status: 415, message: GENERIC_HINT(ext === '' ? '<no-extension>' : ext) });
};

export { officeToMarkdown };
export type { FetchOptions };
