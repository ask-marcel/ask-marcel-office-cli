import { describe, expect, it } from 'bun:test';
import { CONVERTERS } from './converters.ts';
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

describe('the real converters', () => {
  it('are the npm-backed adapters, one for each member of the port', () => {
    expect(CONVERTERS).toEqual({
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
    });
  });

  it('convert HTML to markdown through turndown', () => {
    expect(CONVERTERS.htmlToMarkdown('<p><strong>hi</strong></p>')).toEqual({ ok: true, value: '**hi**' });
  });
});
