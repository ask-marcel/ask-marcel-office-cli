import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { MarkdownEnvelope } from './docx-to-markdown.ts';
import { formatOdfMetadata } from './odf-metadata-to-markdown.ts';
import { extractOdfMetadata } from './odf-metadata.ts';

/**
 * OpenDocument (.odt / .ods / .odp) has no markdown body in this CLI — those
 * inputs route to the `*-as-pdf` commands (Graph `?format=pdf` accepts them).
 * When `--include-metadata true` is set, this returns the side-channel /
 * authored content as a standalone markdown document instead, led by a note
 * pointing at the PDF sibling for the rendered document.
 */

const PDF_NOTE =
  '> OpenDocument body content is not converted here — use the corresponding `*-as-pdf` command for the rendered document. This document carries the side-channel / authored metadata only: properties, keywords, and user-defined custom fields.';

const odfToMarkdown = async (bytes: Uint8Array): Promise<Result<MarkdownEnvelope, GraphError>> => {
  const meta = await extractOdfMetadata(bytes);
  if (!meta.ok) return meta;
  const text = `${PDF_NOTE}\n\n${formatOdfMetadata(meta.value)}`;
  // size = UTF-8 byte count (audit §2.1); `text.length` is UTF-16 code units.
  return ok({ contentType: 'text/markdown', size: new TextEncoder().encode(text).byteLength, text });
};

export { odfToMarkdown };
