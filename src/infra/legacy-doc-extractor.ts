import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { GraphError } from './graph-client.ts';

/**
 * Extract the text of a legacy Word .doc — the pre-2007 OLE binary "Compound
 * File" format that mammoth (which only reads the OOXML .docx zip) cannot parse —
 * via word-extractor (pure-JS, walks the WordDocument OLE stream; no native deps).
 * Body text only: legacy .doc carries no structure this CLI surfaces (headings,
 * tables, and styling are lost), so callers return it as text/plain.
 *
 * try/catch is permitted here per the infra-boundary rule: word-extractor throws
 * on a non-.doc / corrupt OLE container, and we translate that into a Result.
 */
const extractDocText = async (bytes: Uint8Array): Promise<Result<string, GraphError>> => {
  try {
    const { default: WordExtractor } = await import('word-extractor');
    const doc = await new WordExtractor().extract(Buffer.from(bytes));
    return ok(doc.getBody());
  } catch (e) {
    return err({ type: 'api_error', status: 500, message: `doc text extraction failed: ${e instanceof Error ? e.message : String(e)}` });
  }
};

export { extractDocText };
