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
  const size = bytes.byteLength;
  const text = await extractDocText(bytes);
  if (!text.ok) return text;
  return ok({ contentType: 'text/plain', size, text: text.value });
};

export { docToMarkdown };
