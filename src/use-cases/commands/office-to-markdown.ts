import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { docToMarkdown } from './doc-to-markdown.ts';
import { docxToMarkdown } from './docx-to-markdown.ts';
import type { FetchOptions } from './fetch-raw-bytes.ts';
import { fetchRawBytes } from './fetch-raw-bytes.ts';
import { DOCX_FAMILY, ODF_FAMILY, PPTX_FAMILY, XLSX_FAMILY } from './office-extensions.ts';
import { odfToMarkdown } from './odf-to-markdown.ts';
import { pdfToMarkdown } from './pdf-to-markdown.ts';
import { pptxToMarkdown } from './pptx-to-markdown.ts';
import { convertToMarkdown } from './markdown-pipeline.ts';
import { decodeUtf8Text } from './text-passthrough.ts';
import { renderCsvCapped, xlsxToMarkdown } from './xlsx-to-markdown.ts';

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
 *                                         text file, no extension list); non-UTF-8
 *                                         (binary) bytes err pointing at the
 *                                         *-as-pdf command — no short-circuit list
 */

const HTML_FORMAT_INPUTS: ReadonlySet<string> = new Set(['loop', 'fluid', 'wbtx', 'whiteboard']);

const PDF_NO_TEXT_HINT =
  'pdf has no extractable text layer — it looks scanned / image-only (only page images, no embedded text). This command extracts the embedded text layer, not pixels. Use `download-drive-item-as-pdf` to fetch the PDF and read it with a vision-capable model, or run OCR.';

const LEGACY_PPT_HINT =
  'ppt (legacy PowerPoint 97-2003, OLE binary) cannot be converted to markdown here — there is no pure-JS parser for the format. Convert it to PDF first with `download-drive-item-as-pdf` (Graph renders legacy .ppt), then read the PDF with a vision-capable model.';

// The hint for anything that reaches the fallback as non-UTF-8 binary. No
// dedicated archive/media short-circuit list: a known-binary extension is
// fetched and content-sniffed like everything else, then lands here.
const GENERIC_HINT = (ext: string): string =>
  `${ext} not supported by \`*-as-markdown\`. Use the corresponding \`*-as-pdf\` command — Graph \`?format=pdf\` accepts 38 input extensions including this one.`;

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
    // renderCsvCapped (not csvToMarkdownTable): a large standalone `.csv` is
    // subject to the same `--max-cells` OOM cap as an xlsx sheet (audit A2).
    const md = renderCsvCapped(csv, opts.maxCells);
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
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    return pptxToMarkdown(bytes.value, { includeMetadata: opts.includeMetadata });
  }

  if (ODF_FAMILY.has(ext)) {
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    return odfToMarkdown(bytes.value, { includeMetadata: opts.includeMetadata });
  }

  if (ext === 'pdf') {
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    return pdfToMarkdown(bytes.value, PDF_NO_TEXT_HINT);
  }

  if (ext === 'xls') {
    // Legacy Excel (BIFF / OLE binary) — sheetjs `XLSX.read` auto-detects it; legacy
    // has no OOXML side-channel, so `--include-metadata` is not threaded here.
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    return xlsxToMarkdown(bytes.value, { maxCells: opts.maxCells });
  }

  if (ext === 'doc') {
    // Legacy Word (OLE binary) — word-extractor; mammoth only reads the .docx zip.
    const bytes = await fetchRawBytes(graph, contentPath, opts);
    if (!bytes.ok) return bytes;
    return docToMarkdown(bytes.value);
  }

  if (ext === 'ppt') return err({ type: 'api_error', status: 415, message: LEGACY_PPT_HINT });

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
