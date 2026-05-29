import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { MarkdownEnvelope } from './docx-to-markdown.ts';
import { formatPptxMetadata } from './pptx-metadata-to-markdown.ts';
import { extractPptxMetadata } from './pptx-metadata.ts';

/**
 * pptx has no convertible markdown body in this CLI (slide visuals go through
 * the `*-as-pdf` commands, which a vision model reads more reliably than
 * flattened bullets). When `--include-metadata true` is set, this returns the
 * side-channel / authored content as a standalone markdown document instead,
 * led by a note pointing at the PDF sibling for the rendered deck.
 */

const PDF_NOTE =
  '> PowerPoint slide visuals (layout, images, charts) are not converted here — use the corresponding `*-as-pdf` command for the rendered deck. This document carries the side-channel / authored content only: properties, speaker notes, comments, hidden slides, slide tags, and external links.';

const pptxToMarkdown = async (bytes: Uint8Array): Promise<Result<MarkdownEnvelope, GraphError>> => {
  const meta = await extractPptxMetadata(bytes);
  if (!meta.ok) return meta;
  const text = `${PDF_NOTE}\n\n${formatPptxMetadata(meta.value)}`;
  // size = UTF-8 byte count (audit §2.1); `text.length` is UTF-16 code units.
  return ok({ contentType: 'text/markdown', size: new TextEncoder().encode(text).byteLength, text });
};

export { pptxToMarkdown };
