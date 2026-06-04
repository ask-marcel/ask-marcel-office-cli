import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { extractDocText } from '../../infra/legacy-doc-extractor.ts';

type DocTextEnvelope = { readonly contentType: 'text/plain'; readonly size: number; readonly text: string };

/**
 * Convert a legacy Word .doc to a plain-text envelope by extracting its body text
 * (via word-extractor). Output is text/plain, not markdown — legacy .doc carries
 * no structure this CLI recovers. A parse failure propagates as the infra api_error.
 */
const docToMarkdown = async (bytes: Uint8Array): Promise<Result<DocTextEnvelope, GraphError>> => {
  const extracted = await extractDocText(bytes);
  if (!extracted.ok) return extracted;
  // size = byte length of the extracted text (what --output-path writes), consistent
  // with the other converters — not the source .doc's byte count.
  return ok({ contentType: 'text/plain', size: new TextEncoder().encode(extracted.value).byteLength, text: extracted.value });
};

export { docToMarkdown };
