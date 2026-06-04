import { describe, expect, it } from 'bun:test';
import { buildPdfNoImages, buildPdfWithText } from '../../test-helpers/office-fixtures.ts';
import { pdfToMarkdown } from './pdf-to-markdown.ts';

const HINT = 'no text layer — use `download-drive-item-as-pdf` and read it with a vision model';

describe('pdfToMarkdown', () => {
  it('returns the extracted text of a born-digital PDF as a text/plain envelope', async () => {
    const result = await pdfToMarkdown(buildPdfWithText(), HINT);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.contentType).toBe('text/plain');
    expect(result.value.text).toContain('Hello from the');
    expect(result.value.size).toBeGreaterThan(0);
  });

  it('errs 415 with the caller’s vision-model hint when the PDF has no text layer (scanned / image-only)', async () => {
    const result = await pdfToMarkdown(buildPdfNoImages(), HINT);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.type === 'api_error' ? result.error.status : -1).toBe(415);
    expect(result.error.message).toBe(HINT);
  });

  it('propagates the extraction error for bytes that are not a parseable PDF', async () => {
    const result = await pdfToMarkdown(Uint8Array.from([1, 2, 3, 4, 5]), HINT);
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('api_error');
    expect(result.error.message).toContain('pdf text extraction failed');
  });
});
