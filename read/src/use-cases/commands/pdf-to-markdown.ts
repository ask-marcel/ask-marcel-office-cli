import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Converters } from '../ports/converters.ts';

type PdfTextEnvelope = { readonly contentType: 'text/plain'; readonly size: number; readonly text: string; readonly pageCount: number };

/**
 * Convert a PDF to a plain-text envelope by extracting its text layer (via unpdf).
 * A born-digital PDF yields its reading-order text; a scanned / image-only PDF has
 * no text layer, so the caller's `noTextHint` (pointing at the matching *-as-pdf
 * command + a vision model) is returned as a 415 rather than an empty body. Output
 * is text/plain, not structured markdown — pdfjs flattens layout into reading order.
 * `size` is the byte length of the produced text (what `--output-path` writes),
 * consistent with the other converters — NOT the source PDF's byte count.
 */
const pdfToMarkdown = async (converters: Converters, bytes: Uint8Array, noTextHint: string): Promise<Result<PdfTextEnvelope, GraphError>> => {
  const extracted = await converters.extractPdfText(bytes);
  if (!extracted.ok) return extracted;
  const text = extracted.value.text.trim();
  if (text === '') return err({ type: 'api_error', status: 415, message: noTextHint });
  return ok({ contentType: 'text/plain', size: new TextEncoder().encode(text).byteLength, text, pageCount: extracted.value.pageCount });
};

export { pdfToMarkdown };
