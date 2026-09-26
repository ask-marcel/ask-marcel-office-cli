import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { FetchOptions } from './fetch-raw-bytes.ts';
import { fetchRawBytes } from './fetch-raw-bytes.ts';
import { bytesToMarkdown } from './markdown-dispatch.ts';
import type { ConversionHints } from './markdown-dispatch.ts';
import { convertToMarkdown } from './markdown-pipeline.ts';
import { extensionOf } from './text-passthrough.ts';
import { refuseSheet } from './xlsx-to-markdown.ts';

/**
 * `*-as-markdown` dispatcher for a OneDrive / SharePoint drive item. Loop/Fluid/
 * Whiteboard need a Graph `?format=html` round-trip (the only inputs Graph's HTML
 * conversion accepts) and are handled here; everything else is fetched once and
 * routed through the shared `bytesToMarkdown` core (docx/xlsx/pptx/odf/csv/pdf/
 * legacy .xls/.doc, content-sniff fallback) — the same core convert-mail and
 * convert-drive-item-zip-to-markdown use, so all three agree on every extension.
 */

const HTML_FORMAT_INPUTS: ReadonlySet<string> = new Set(['loop', 'fluid', 'wbtx', 'whiteboard']);

const DRIVE_HINTS: ConversionHints = {
  pdfNoText:
    'pdf has no extractable text layer — it looks scanned / image-only (only page images, no embedded text). This command extracts the embedded text layer, not pixels. Use `download-drive-item-as-pdf` to fetch the PDF and read it with a vision-capable model, or run OCR.',
  legacyPpt:
    'ppt (legacy PowerPoint 97-2003, OLE binary) cannot be converted to markdown here — there is no pure-JS parser for the format. Convert it to PDF first with `download-drive-item-as-pdf` (Graph renders legacy .ppt), then read the PDF with a vision-capable model.',
  image: (ext) =>
    `${ext} is an image and cannot be converted to markdown. Fetch the bytes with \`download-drive-item-content\` and feed them to a vision-capable model, or pull images embedded in a document with \`extract-drive-item-images\`.`,
  generic: (ext) => `${ext} not supported by \`*-as-markdown\`. Use the corresponding \`*-as-pdf\` command — Graph \`?format=pdf\` accepts 38 input extensions including this one.`,
};

type OfficeToMarkdownOptions = FetchOptions & {
  readonly includeMetadata?: boolean;
  readonly inlineImages?: boolean;
  readonly maxCells?: number;
  readonly keepQuoted?: boolean;
  readonly sheet?: string;
};

// Graph's HTML conversion of a fresh Loop page can come back empty while the
// page already holds content (a meeting-notes page with one 14 KB save rendered
// to nothing, 2026-09-19); the converter lags the saves. An empty body is
// therefore said, so the caller retries later instead of reading "no notes".
const EMPTY_LOOP_NOTE =
  'Graph returned no HTML for this page: its Loop converter lags the saves, sometimes by hours. Retry later; `list-drive-item-versions` shows when the page was last saved.';

const withEmptyLoopNote = (result: Result<unknown, GraphError>): Result<unknown, GraphError> => {
  if (!result.ok) return result;
  const envelope = result.value as { readonly text?: unknown };
  if (typeof envelope.text !== 'string' || envelope.text.trim() !== '') return result;
  return ok({ ...(result.value as Record<string, unknown>), note: EMPTY_LOOP_NOTE });
};

const officeToMarkdown = async (graph: GraphClient, contentPath: string, filename: string, opts: OfficeToMarkdownOptions = {}): Promise<Result<unknown, GraphError>> => {
  const ext = extensionOf(filename);
  if (opts.sheet !== undefined && HTML_FORMAT_INPUTS.has(ext)) return refuseSheet(`this file is a .${ext}`);
  if (HTML_FORMAT_INPUTS.has(ext)) return withEmptyLoopNote(await convertToMarkdown(graph, `${contentPath}?format=html`));
  const bytes = await fetchRawBytes(graph, contentPath, opts);
  if (!bytes.ok) return bytes;
  return bytesToMarkdown(bytes.value, filename, opts, DRIVE_HINTS);
};

export { officeToMarkdown };
