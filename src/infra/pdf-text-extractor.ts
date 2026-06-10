import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { GraphError } from './graph-client.ts';

/**
 * Extract a PDF's text layer via unpdf (a pure-JS pdfjs build — no native deps,
 * runs under Bun). Returns the reading-order text of every page merged into one
 * string. A born-digital PDF (exported from Word/LaTeX/Chrome/…) yields its text;
 * a scanned / image-only PDF has no text layer, so the string comes back empty —
 * callers treat that as "needs OCR / a vision model", never as success-with-text.
 *
 * Scope note: this is the embedded text layer, NOT OCR. It also does not preserve
 * layout structure (headings, tables, columns flatten into reading order).
 *
 * `try/catch` is permitted here per the infra-boundary rule: pdfjs throws on
 * malformed input and we translate that into a Result rather than letting it escape.
 *
 * `verbosity: 0` (pdfjs `VerbosityLevel.ERRORS`) silences pdfjs's TrueType
 * font-hinting log spam — `Warning: TT: undefined function: N`, `Required "glyf"
 * table is not found`, `Indexing all PDF objects` — which it emits via its own
 * `console.warn` while parsing a PDF's embedded fonts. Font hinting is sub-pixel
 * glyph rendering, completely irrelevant to text extraction (the text still comes
 * out fine); the default level (WARNINGS) just floods the host app's logs once per
 * quirky font. We lower it at the boundary because the noise originates inside
 * pdfjs — the `no-console` rule keeps us from emitting or intercepting it ourselves.
 */
/**
 * Map a pdfjs throw to an actionable message. The special case is an encrypted
 * PDF: pdfjs raises `PasswordException` ("No password given") — QA run-1 found a
 * real tenant file surfacing that raw string with zero guidance. Password
 * handling is genuinely unsupported (and Graph's `format=pdf` cannot unlock
 * encrypted sources either), so say that honestly instead of leaking pdfjs
 * internals. Pure and exported: the branch is unit-tested with shaped errors —
 * no encrypted-PDF fixture exists in the toolchain.
 */
const pdfErrorMessage = (e: unknown): string => {
  const name = e instanceof Error ? e.name : '';
  const detail = e instanceof Error ? e.message : String(e);
  if (name === 'PasswordException' || /password/i.test(detail)) {
    return 'pdf is password-protected — its text layer cannot be read without the password, and Graph’s format=pdf cannot unlock encrypted sources either. Ask the sender for an unprotected copy, or open it locally with the password and re-save it without encryption.';
  }
  return `pdf text extraction failed: ${detail}`;
};

const extractPdfText = async (bytes: Uint8Array): Promise<Result<string, GraphError>> => {
  try {
    const { extractText, getDocumentProxy } = await import('unpdf');
    const doc = await getDocumentProxy(bytes, { verbosity: 0 });
    const { text } = await extractText(doc, { mergePages: true });
    return ok(text);
  } catch (e) {
    return err({ type: 'api_error', status: 415, message: pdfErrorMessage(e) });
  }
};

export { extractPdfText, pdfErrorMessage };
