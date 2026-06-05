import type { Result } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { FetchOptions } from './fetch-raw-bytes.ts';
import { fetchRawBytes } from './fetch-raw-bytes.ts';
import { bytesToMarkdown } from './markdown-dispatch.ts';
import type { ConversionHints } from './markdown-dispatch.ts';
import { convertToMarkdown } from './markdown-pipeline.ts';
import { extensionOf } from './text-passthrough.ts';

/**
 * `*-as-markdown` dispatcher for a OneDrive / SharePoint drive item. Loop/Fluid/
 * Whiteboard need a Graph `?format=html` round-trip (the only inputs Graph's HTML
 * conversion accepts) and are handled here; everything else is fetched once and
 * routed through the shared `bytesToMarkdown` core (docx/xlsx/pptx/odf/csv/pdf/
 * legacy .xls/.doc, content-sniff fallback) — the same core convert-mail and
 * convert-drive-item-zip use, so all three agree on every extension.
 */

const HTML_FORMAT_INPUTS: ReadonlySet<string> = new Set(['loop', 'fluid', 'wbtx', 'whiteboard']);

const DRIVE_HINTS: ConversionHints = {
  pdfNoText:
    'pdf has no extractable text layer — it looks scanned / image-only (only page images, no embedded text). This command extracts the embedded text layer, not pixels. Use `download-drive-item-as-pdf` to fetch the PDF and read it with a vision-capable model, or run OCR.',
  legacyPpt:
    'ppt (legacy PowerPoint 97-2003, OLE binary) cannot be converted to markdown here — there is no pure-JS parser for the format. Convert it to PDF first with `download-drive-item-as-pdf` (Graph renders legacy .ppt), then read the PDF with a vision-capable model.',
  image: (ext) =>
    `${ext} is an image and cannot be converted to markdown. Fetch the bytes with \`download-onedrive-file-content\` and feed them to a vision-capable model, or pull images embedded in a document with \`extract-drive-item-images\`.`,
  generic: (ext) => `${ext} not supported by \`*-as-markdown\`. Use the corresponding \`*-as-pdf\` command — Graph \`?format=pdf\` accepts 38 input extensions including this one.`,
};

type OfficeToMarkdownOptions = FetchOptions & { readonly includeMetadata?: boolean; readonly inlineImages?: boolean; readonly maxCells?: number };

const officeToMarkdown = async (graph: GraphClient, contentPath: string, filename: string, opts: OfficeToMarkdownOptions = {}): Promise<Result<unknown, GraphError>> => {
  const ext = extensionOf(filename);
  if (HTML_FORMAT_INPUTS.has(ext)) return convertToMarkdown(graph, `${contentPath}?format=html`);
  const bytes = await fetchRawBytes(graph, contentPath, opts);
  if (!bytes.ok) return bytes;
  return bytesToMarkdown(bytes.value, filename, opts, DRIVE_HINTS);
};

export { officeToMarkdown };
export type { FetchOptions, OfficeToMarkdownOptions };
