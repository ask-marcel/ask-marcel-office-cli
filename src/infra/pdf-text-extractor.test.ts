import { describe, expect, it } from 'bun:test';
import { buildPdfNoImages, buildPdfWithText } from '../test-helpers/office-fixtures.ts';
import { extractPdfText, pdfErrorMessage } from './pdf-text-extractor.ts';

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

describe('pdfErrorMessage (QA-006 — encrypted PDFs)', () => {
  it('a pdfjs PasswordException becomes an honest "password-protected" message with a way forward', () => {
    const passwordError = Object.assign(new Error('No password given'), { name: 'PasswordException' });
    const message = pdfErrorMessage(passwordError);
    expect(message).toContain('password-protected');
    expect(message).toContain('format=pdf cannot unlock');
    expect(message).not.toContain('No password given'); // no raw pdfjs internals
  });

  it('a password-mentioning message without the exception name still routes to the protected-pdf wording', () => {
    expect(pdfErrorMessage(new Error('Incorrect Password'))).toContain('password-protected');
  });

  it('any other parse failure keeps the generic extraction-failed message with the underlying detail', () => {
    expect(pdfErrorMessage(new Error('Invalid PDF structure'))).toBe('pdf text extraction failed: Invalid PDF structure');
    expect(pdfErrorMessage('plain string throw')).toBe('pdf text extraction failed: plain string throw');
  });
});
