import { describe, expect, it } from 'bun:test';
import { buildPdfNoImages, buildPdfWithText } from '../test-helpers/office-fixtures.ts';
import { extractPdfText } from './pdf-text-extractor.ts';

describe('extractPdfText', () => {
  it('returns the text layer of a born-digital PDF', async () => {
    const result = await extractPdfText(buildPdfWithText());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value).toContain('Hello from the'); // fixture's no-/Widths font drops the final glyph in pdfjs
  });

  it('returns an empty string for a PDF that paints no text (e.g. a scanned/image-only page)', async () => {
    const result = await extractPdfText(buildPdfNoImages());
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.trim()).toBe('');
  });

  it('returns an api_error for bytes that are not a parseable PDF', async () => {
    const result = await extractPdfText(Uint8Array.from([1, 2, 3, 4, 5]));
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    if (result.error.type !== 'api_error') return;
    expect(result.error.message).toContain('pdf text extraction failed');
  });
});
