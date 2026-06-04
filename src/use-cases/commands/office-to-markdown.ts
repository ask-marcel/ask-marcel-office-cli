import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { docxToMarkdown } from './docx-to-markdown.ts';
import type { FetchOptions } from './fetch-raw-bytes.ts';
import { fetchRawBytes } from './fetch-raw-bytes.ts';
import { DOCX_FAMILY, ODF_FAMILY, PPTX_FAMILY, XLSX_FAMILY } from './office-extensions.ts';
import { odfToMarkdown } from './odf-to-markdown.ts';
import { pptxToMarkdown } from './pptx-to-markdown.ts';
import { convertToMarkdown } from './markdown-pipeline.ts';
import { decodeUtf8Text } from './text-passthrough.ts';
import { csvToMarkdownTable, xlsxToMarkdown } from './xlsx-to-markdown.ts';

/**
 * Strategy dispatcher for `*-as-markdown` commands. Picks the right
 * conversion path based on the file extension and routes through it:
 *
 * - csv                                 → csvToMarkdownTable (markdown table)
 * - docx                                → mammoth → turndown
 * - xlsx                                → sheetjs → markdown table per sheet
 * - odt / ods / odp                     → walk content.xml (headings/lists/
 *                                         tables/sheets/slides); +metadata block
 * - loop / fluid / wbtx / whiteboard    → existing Graph ?format=html
 *                                         pipeline (the only inputs Graph
 *                                         HTML conversion actually accepts)
 * - anything else                       → content-sniff: bytes that decode as
 *                                         valid UTF-8 are returned as text (any
 *                                         text file, no extension list); a
 *                                         known-binary ext or non-UTF-8 bytes
 *                                         err pointing at the *-as-pdf command
 */

const HTML_FORMAT_INPUTS: ReadonlySet<string> = new Set(['loop', 'fluid', 'wbtx', 'whiteboard']);

const PPTX_HINT =
  'pptx not supported by `*-as-markdown`. Use the corresponding `*-as-pdf` command — Graph PDF conversion preserves slide layout, and a vision-capable LLM reads it more reliably than flattened slide-by-slide bullets. Or pass `--include-metadata true` to extract the side-channel content (speaker notes, comments, hidden slides, properties, tags, links) as a `## PPTX metadata` document.';

// Extensions Graph's `?format=pdf` does NOT accept — pointing the user at
// `*-as-pdf` for these would trade one InputFormatNotSupported error for
// another. Surfaced here as a no-conversion-path-exists hint so the user
// can fetch raw bytes via `get-mail-attachment` / `download-onedrive-file-content`
// and process locally.
const PDF_UNSUPPORTED: ReadonlySet<string> = new Set([
  'zip',
  'rar',
  '7z',
  'tar',
  'gz',
  'tgz',
  'mp3',
  'mp4',
  'mov',
  'wav',
  'avi',
  'mkv',
  'png',
  'jpg',
  'jpeg',
  'gif',
  'webp',
  'bmp',
  'svg',
  'exe',
  'dmg',
  'iso',
]);

const GENERIC_HINT = (ext: string): string => {
  if (PDF_UNSUPPORTED.has(ext)) {
    return `${ext} cannot be converted to markdown OR pdf — Graph rejects this extension on both paths. Fetch the raw bytes via \`get-mail-attachment\` (mail context) or \`download-onedrive-file-content\` (drive context) and process locally; pair with \`--output-path\` to land them on disk.`;
  }
  return `${ext} not supported by \`*-as-markdown\`. Use the corresponding \`*-as-pdf\` command — Graph \`?format=pdf\` accepts 38 input extensions including this one.`;
};

const extensionOf = (filename: string): string => {
  const dot = filename.lastIndexOf('.');
  if (dot === -1 || dot === filename.length - 1) return '';
  return filename.slice(dot + 1).toLowerCase();
};

type OfficeToMarkdownOptions = FetchOptions & { readonly includeMetadata?: boolean; readonly inlineImages?: boolean; readonly maxCells?: number };

const officeToMarkdown = async (graph: GraphClient, contentPath: string, filename: string, opts: OfficeToMarkdownOptions = {}): Promise<Result<unknown, GraphError>> => {
  const ext = extensionOf(filename);

  if (ext === 'csv') {
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    const csv = new TextDecoder().decode(bytes.value);
    const md = csvToMarkdownTable(csv);
    // size = UTF-8 byte count of the markdown output (matches what
    // --output-path would write to disk). `md.length` is UTF-16 code units
    // and undercounts files with non-ASCII content (audit §2.1).
    return ok({ contentType: 'text/markdown', size: new TextEncoder().encode(md).byteLength, text: md });
  }

  if (DOCX_FAMILY.has(ext)) {
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    return docxToMarkdown(bytes.value, { includeMetadata: opts.includeMetadata, inlineImages: opts.inlineImages });
  }

  if (XLSX_FAMILY.has(ext)) {
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    return xlsxToMarkdown(bytes.value, { includeMetadata: opts.includeMetadata, maxCells: opts.maxCells });
  }

  if (HTML_FORMAT_INPUTS.has(ext)) {
    return convertToMarkdown(graph, `${contentPath}?format=html`);
  }

  if (PPTX_FAMILY.has(ext)) {
    if (opts.includeMetadata !== true) return err({ type: 'api_error', status: 415, message: PPTX_HINT });
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    return pptxToMarkdown(bytes.value);
  }

  if (ODF_FAMILY.has(ext)) {
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    return odfToMarkdown(bytes.value, { includeMetadata: opts.includeMetadata });
  }

  // Known-binary extensions: hint straight away, no wasted download.
  if (PDF_UNSUPPORTED.has(ext)) return err({ type: 'api_error', status: 415, message: GENERIC_HINT(ext) });

  // Fallback: content-sniff — any file whose bytes decode as valid UTF-8 comes
  // back as text (a .txt/.md/.conf, or even an extensionless README, with no
  // extension list to maintain); genuinely binary input gets the hint.
  const bytes = await fetchRawBytes(graph, contentPath, opts);
  if (!bytes.ok) return bytes;
  const text = decodeUtf8Text(bytes.value);
  if (text !== undefined) return ok({ contentType: 'text/plain', size: bytes.value.byteLength, text });
  return err({ type: 'api_error', status: 415, message: GENERIC_HINT(ext === '' ? '<no-extension>' : ext) });
};

export { officeToMarkdown };
export type { FetchOptions, OfficeToMarkdownOptions };
