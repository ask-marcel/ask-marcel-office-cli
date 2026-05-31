import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { MarkdownEnvelope } from './docx-to-markdown.ts';
import { odfContentToMarkdown } from './odf-content-to-markdown.ts';
import { formatOdfMetadata } from './odf-metadata-to-markdown.ts';
import { extractOdfMetadata } from './odf-metadata.ts';

/**
 * Converts an OpenDocument (.odt / .ods / .odp) body to markdown by walking
 * `content.xml` (headings / paragraphs / lists / tables for text docs, named
 * sheet-tables for spreadsheets, per-slide text for presentations — including
 * style-hidden content a rendered viewer would suppress). With
 * `--include-metadata true` the side-channel / authored metadata
 * (properties, keywords, user-defined fields) is appended as a trailing block.
 */

type OdfToMarkdownOptions = { readonly includeMetadata?: boolean };

const odfToMarkdown = async (bytes: Uint8Array, opts: OdfToMarkdownOptions = {}): Promise<Result<MarkdownEnvelope, GraphError>> => {
  const body = await odfContentToMarkdown(bytes);
  if (!body.ok) return body;
  let text = body.value;
  if (opts.includeMetadata === true) {
    const meta = await extractOdfMetadata(bytes);
    if (!meta.ok) return meta;
    const block = formatOdfMetadata(meta.value);
    text = text === '' ? block : `${text}\n\n${block}`;
  }
  // size = UTF-8 byte count (audit §2.1); `text.length` is UTF-16 code units.
  return ok({ contentType: 'text/markdown', size: new TextEncoder().encode(text).byteLength, text });
};

export { odfToMarkdown };
export type { OdfToMarkdownOptions };
