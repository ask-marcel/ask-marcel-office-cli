import type { Converters } from '../use-cases/ports/converters.ts';
import { extractEml } from './eml-parser-adapter.ts';
import { extractDocText } from './legacy-doc-extractor.ts';
import { lineDiff } from './line-diff-adapter.ts';
import { mammothToHtml } from './mammoth-adapter.ts';
import { extractMsg } from './msg-reader-adapter.ts';
import { extractOoxmlMedia } from './ooxml-media-extractor.ts';
import { openOoxmlZip } from './ooxml-zip-adapter.ts';
import { extractPdfImages } from './pdf-image-extractor.ts';
import { extractPdfText } from './pdf-text-extractor.ts';
import { readSheetsAsCsv } from './sheetjs-adapter.ts';
import { htmlToMarkdown, stripHtmlToText } from './turndown-adapter.ts';
import { openZipEntries } from './zip-reader.ts';

/** The real converters, on the npm libraries; the composition root passes them to the commands (ADR 0004). */
export const CONVERTERS: Converters = {
  htmlToMarkdown,
  stripHtmlToText,
  mammothToHtml,
  openOoxmlZip,
  openZipEntries,
  extractOoxmlMedia,
  extractPdfImages,
  extractPdfText,
  extractDocText,
  extractMsg,
  extractEml,
  readSheetsAsCsv,
  lineDiff,
};
