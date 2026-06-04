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
 */
const extractPdfText = async (bytes: Uint8Array): Promise<Result<string, GraphError>> => {
  try {
    const { extractText, getDocumentProxy } = await import('unpdf');
    const doc = await getDocumentProxy(bytes);
    const { text } = await extractText(doc, { mergePages: true });
    return ok(text);
  } catch (e) {
    return err({ type: 'api_error', status: 500, message: `pdf text extraction failed: ${e instanceof Error ? e.message : String(e)}` });
  }
};

export { extractPdfText };
