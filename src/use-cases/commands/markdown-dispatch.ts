import type { Result } from '../../domain/result.ts';
import { err, map, ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { htmlToMarkdown, stripHtmlToText } from '../../infra/turndown-adapter.ts';
import { docToMarkdown } from './doc-to-markdown.ts';
import { docxToMarkdown, stripInlineImages } from './docx-to-markdown.ts';
import { emlToMarkdown } from './eml-to-markdown.ts';
import { msgToMarkdown } from './msg-to-markdown.ts';
import { odfToMarkdown } from './odf-to-markdown.ts';
import { DOCX_FAMILY, ODF_FAMILY, PPTX_FAMILY, XLSX_FAMILY } from './office-extensions.ts';
import { pdfToMarkdown } from './pdf-to-markdown.ts';
import { pptxToMarkdown } from './pptx-to-markdown.ts';
import { decodeUtf8Text, extensionOf } from './text-passthrough.ts';
import { refuseSheet, renderCsvCapped, xlsxToMarkdown } from './xlsx-to-markdown.ts';

// Image extensions that have no markdown text representation. NOTE: `svg` is NOT
// here — an SVG is XML text, so it content-sniffs to text/plain like any text file.
const IMAGE_EXTENSIONS: ReadonlySet<string> = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tiff', 'tif', 'ico']);

// A saved web page is markup, not prose: it goes through turndown like a mail body.
// The raw source stays one command away (`download-drive-item-content`).
const HTML_EXTENSIONS: ReadonlySet<string> = new Set(['html', 'htm']);

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
// `keepQuoted` is `.msg`-only (every other format has no reply chain); it rides the
// `--keep-quoted` flag on each command that can hand a `.msg` to the dispatch.
type BytesToMarkdownOptions = {
  readonly includeMetadata?: boolean;
  readonly maxCells?: number;
  readonly inlineImages?: boolean;
  readonly keepQuoted?: boolean;
  readonly sheet?: string;
  readonly depth?: number;
};

// Hints for files NESTED inside a container (zip entry, .msg attachment). The
// caller-specific hints point at sibling commands (`download-drive-item-as-pdf`,
// `extract-drive-item-images`, …) that can only address top-level drive items /
// attachments — they cannot reach a file INSIDE a container. Nested
// conversions therefore always use this container-neutral wording.
const NESTED_HINTS: ConversionHints = {
  pdfNoText:
    'pdf has no extractable text layer (scanned / image-only) — extract it from the archive/message first (for a mail or a forwarded mail, `extract-mail-attachment-images --pages` returns its page images), then read it with a vision-capable model',
  legacyPpt:
    'ppt (legacy PowerPoint, OLE binary) has no markdown path — extract it, convert it to PDF first (e.g. upload to OneDrive and use `download-drive-item-as-pdf`), then read it with a vision model',
  image: (ext) =>
    `${ext} is an image — extract it from the archive/message first (for a mail or a forwarded mail, \`extract-mail-attachment-images\` returns it), then read it with a vision-capable model`,
  generic: (ext) => `${ext} is not a convertible Office/text format (images, binaries, and nested archives are not unpacked here)`,
};

const markdownEnvelope = (md: string): { contentType: 'text/markdown'; size: number; text: string } => ({
  contentType: 'text/markdown',
  size: new TextEncoder().encode(md).byteLength,
  text: md,
});

const csvEnvelope = (bytes: Uint8Array, maxCells: number | undefined): Result<unknown, GraphError> =>
  ok(markdownEnvelope(renderCsvCapped(new TextDecoder().decode(bytes), maxCells)));

// The <head> (title, meta, stylesheets, scripts) is page chrome; turndown would print
// the title as a stray first line. `[\s>]` keeps a body <header> element.
const withoutHead = (html: string): string => html.replace(/<head[\s>][\s\S]*?<\/head>/i, '');

// turndown builds a DOM of the whole page: a 1 MB table page took 4 s and half a
// gigabyte of memory. Past this size the page is flattened to text in one pass.
const HTML_MARKDOWN_CAP = 1_000_000;

// An embedded `data:` image becomes `[image: alt]`, as in a Word file, unless
// `--inline-images true` asks for the base64.
const htmlEnvelope = (bytes: Uint8Array, inlineImages: boolean | undefined): Result<unknown, GraphError> => {
  const html = withoutHead(new TextDecoder().decode(bytes));
  if (bytes.byteLength <= HTML_MARKDOWN_CAP) return map(htmlToMarkdown(html), (md) => markdownEnvelope(inlineImages === true ? md : stripInlineImages(md)));
  const text = stripHtmlToText(html);
  return ok({
    contentType: 'text/plain',
    size: new TextEncoder().encode(text).byteLength,
    text,
    note: `This page is ${(bytes.byteLength / 1_000_000).toFixed(1)} MB, over the 1 MB cap for a markdown render, so it is flattened to plain text: each table row comes out on one line, its cells separated by |.`,
  });
};

/**
 * The single extension→converter dispatch for every markdown command, operating on
 * bytes already in hand: download-drive-item-as-markdown fetches them, convert-mail
 * decodes the attachment, convert-drive-item-zip-to-markdown unpacks the entry. Loop/Fluid/
 * Whiteboard (`?format=html`) need a Graph round-trip and are handled by the drive
 * caller BEFORE this — they never reach here. An Outlook `.msg` is rendered to markdown
 * (headers + body) with each of its own attachments recursed through this dispatch.
 * `hints` carry the caller-specific messages; a non-text result is an `err` the caller
 * surfaces (a 415, or a zip note).
 */
const bytesToMarkdown = async (bytes: Uint8Array, filename: string, opts: BytesToMarkdownOptions, hints: ConversionHints): Promise<Result<unknown, GraphError>> => {
  const ext = extensionOf(filename);
  if (opts.sheet !== undefined && !XLSX_FAMILY.has(ext) && ext !== 'xls') return refuseSheet(`this file is a .${ext}`);
  if (ext === 'csv') return csvEnvelope(bytes, opts.maxCells);
  if (HTML_EXTENSIONS.has(ext)) return htmlEnvelope(bytes, opts.inlineImages);
  if (DOCX_FAMILY.has(ext)) return docxToMarkdown(bytes, { includeMetadata: opts.includeMetadata, inlineImages: opts.inlineImages });
  if (XLSX_FAMILY.has(ext)) return xlsxToMarkdown(bytes, { includeMetadata: opts.includeMetadata, maxCells: opts.maxCells, sheet: opts.sheet });
  if (PPTX_FAMILY.has(ext)) return pptxToMarkdown(bytes, { includeMetadata: opts.includeMetadata });
  if (ODF_FAMILY.has(ext)) return odfToMarkdown(bytes, { includeMetadata: opts.includeMetadata });
  if (ext === 'pdf') return pdfToMarkdown(bytes, hints.pdfNoText);
  if (ext === 'xls') return xlsxToMarkdown(bytes, { maxCells: opts.maxCells, sheet: opts.sheet }); // legacy Excel — no OOXML side-channel
  if (ext === 'doc') return docToMarkdown(bytes); // legacy Word — text only
  if (ext === 'ppt') return err({ type: 'api_error', status: 415, code: 'unsupported_legacy_office', message: hints.legacyPpt });
  if (ext === 'msg' || ext === 'eml') {
    // An Outlook .msg or a raw RFC 822 .eml: render headers + body and recurse each
    // attachment through this same dispatch (the zip pattern), incrementing depth
    // so a message inside a message can't loop. Attachments are NESTED files —
    // container-neutral hints, not the caller's (a png inside a .msg must not
    // point at drive-item commands).
    const depth = opts.depth ?? 0;
    const render = ext === 'msg' ? msgToMarkdown : emlToMarkdown;
    return render(bytes, { depth, keepQuoted: opts.keepQuoted }, (childBytes, childName) => bytesToMarkdown(childBytes, childName, { ...opts, depth: depth + 1 }, NESTED_HINTS));
  }
  if (IMAGE_EXTENSIONS.has(ext)) return err({ type: 'api_error', status: 415, code: 'unsupported_image', message: hints.image(ext) });
  const text = decodeUtf8Text(bytes);
  if (text !== undefined) return ok({ contentType: 'text/plain', size: bytes.byteLength, text });
  return err({ type: 'api_error', status: 415, code: 'unsupported_format', message: hints.generic(ext === '' ? '<no-extension>' : ext) });
};

export { bytesToMarkdown, IMAGE_EXTENSIONS, NESTED_HINTS };
export type { BytesToMarkdownOptions, ConversionHints };
