import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { docToMarkdown } from './doc-to-markdown.ts';
import { docxToMarkdown } from './docx-to-markdown.ts';
import { msgToMarkdown } from './msg-to-markdown.ts';
import { odfToMarkdown } from './odf-to-markdown.ts';
import { DOCX_FAMILY, ODF_FAMILY, PPTX_FAMILY, XLSX_FAMILY } from './office-extensions.ts';
import { pdfToMarkdown } from './pdf-to-markdown.ts';
import { pptxToMarkdown } from './pptx-to-markdown.ts';
import { decodeUtf8Text, extensionOf } from './text-passthrough.ts';
import { renderCsvCapped, xlsxToMarkdown } from './xlsx-to-markdown.ts';

// Image extensions that have no markdown text representation. NOTE: `svg` is NOT
// here — an SVG is XML text, so it content-sniffs to text/plain like any text file.
const IMAGE_EXTENSIONS: ReadonlySet<string> = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tiff', 'tif', 'ico']);

// Context-specific hint messages. The dispatch ladder is shared; each caller (drive /
// mail / zip) supplies its own wording (which sibling command to use, raw-bytes route).
type ConversionHints = {
  readonly pdfNoText: string;
  readonly legacyPpt: string;
  readonly image: (ext: string) => string;
  readonly generic: (ext: string) => string;
};

// `depth` is internal recursion plumbing for `.msg` attachments (a .msg can attach
// another .msg); callers never set it. It caps `.msg`-inside-`.msg` nesting.
type BytesToMarkdownOptions = { readonly includeMetadata?: boolean; readonly maxCells?: number; readonly inlineImages?: boolean; readonly depth?: number };

// Hints for files NESTED inside a container (zip entry, .msg attachment). The
// caller-specific hints point at sibling commands (`download-drive-item-as-pdf`,
// `extract-drive-item-images`, …) that can only address top-level drive items /
// attachments — they cannot reach a file INSIDE a container (QA-007). Nested
// conversions therefore always use this container-neutral wording.
const NESTED_HINTS: ConversionHints = {
  pdfNoText: 'pdf has no extractable text layer (scanned / image-only) — extract it from the archive/message first, then read it with a vision-capable model',
  legacyPpt: 'ppt (legacy PowerPoint, OLE binary) has no markdown path — extract it, convert it to PDF first (e.g. upload to OneDrive and use `download-drive-item-as-pdf`), then read it with a vision model',
  image: (ext) => `${ext} is an image — extract it from the archive/message first, then read it with a vision-capable model`,
  generic: (ext) => `${ext} is not a convertible Office/text format (images, binaries, and nested archives are not unpacked here)`,
};

const csvEnvelope = (bytes: Uint8Array, maxCells: number | undefined): Result<unknown, GraphError> => {
  const md = renderCsvCapped(new TextDecoder().decode(bytes), maxCells);
  return ok({ contentType: 'text/markdown', size: new TextEncoder().encode(md).byteLength, text: md });
};

/**
 * The single extension→converter dispatch for every markdown command, operating on
 * bytes already in hand: download-drive-item-as-markdown fetches them, convert-mail
 * decodes the attachment, convert-drive-item-zip unpacks the entry. Loop/Fluid/
 * Whiteboard (`?format=html`) need a Graph round-trip and are handled by the drive
 * caller BEFORE this — they never reach here. An Outlook `.msg` is rendered to markdown
 * (headers + body) with each of its own attachments recursed through this dispatch.
 * `hints` carry the caller-specific messages; a non-text result is an `err` the caller
 * surfaces (a 415, or a zip note).
 */
const bytesToMarkdown = async (bytes: Uint8Array, filename: string, opts: BytesToMarkdownOptions, hints: ConversionHints): Promise<Result<unknown, GraphError>> => {
  const ext = extensionOf(filename);
  if (ext === 'csv') return csvEnvelope(bytes, opts.maxCells);
  if (DOCX_FAMILY.has(ext)) return docxToMarkdown(bytes, { includeMetadata: opts.includeMetadata, inlineImages: opts.inlineImages });
  if (XLSX_FAMILY.has(ext)) return xlsxToMarkdown(bytes, { includeMetadata: opts.includeMetadata, maxCells: opts.maxCells });
  if (PPTX_FAMILY.has(ext)) return pptxToMarkdown(bytes, { includeMetadata: opts.includeMetadata });
  if (ODF_FAMILY.has(ext)) return odfToMarkdown(bytes, { includeMetadata: opts.includeMetadata });
  if (ext === 'pdf') return pdfToMarkdown(bytes, hints.pdfNoText);
  if (ext === 'xls') return xlsxToMarkdown(bytes, { maxCells: opts.maxCells }); // legacy Excel — no OOXML side-channel
  if (ext === 'doc') return docToMarkdown(bytes); // legacy Word — text only
  if (ext === 'ppt') return err({ type: 'api_error', status: 415, message: hints.legacyPpt });
  if (ext === 'msg') {
    // Outlook .msg: render headers + body and recurse each attachment through this
    // same dispatch (the zip pattern), incrementing depth so a .msg-in-.msg can't
    // loop. Attachments are NESTED files — container-neutral hints, not the
    // caller's (QA-007: a png inside a .msg must not point at drive-item commands).
    const depth = opts.depth ?? 0;
    return msgToMarkdown(bytes, { depth }, (childBytes, childName) => bytesToMarkdown(childBytes, childName, { ...opts, depth: depth + 1 }, NESTED_HINTS));
  }
  if (IMAGE_EXTENSIONS.has(ext)) return err({ type: 'api_error', status: 415, message: hints.image(ext) });
  const text = decodeUtf8Text(bytes);
  if (text !== undefined) return ok({ contentType: 'text/plain', size: bytes.byteLength, text });
  return err({ type: 'api_error', status: 415, message: hints.generic(ext === '' ? '<no-extension>' : ext) });
};

export { bytesToMarkdown, IMAGE_EXTENSIONS, NESTED_HINTS };
export type { BytesToMarkdownOptions, ConversionHints };
