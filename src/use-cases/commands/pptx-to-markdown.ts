import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { MarkdownEnvelope } from './docx-to-markdown.ts';
import { formatPptxMetadata } from './pptx-metadata-to-markdown.ts';
import { extractPptxMetadata } from './pptx-metadata.ts';
import type { Slide } from './pptx-slides.ts';

/**
 * Renders a .pptx to markdown: one `## Slide N` section per slide carrying the
 * visible slide text (title, bullets, text boxes, table cells) with the speaker
 * notes inline beneath it. With `includeMetadata`, the `## PPTX metadata` side-
 * channel block (properties, external links, slide tags, comments) is appended.
 *
 * Caveat: slide text comes out in document order, not guaranteed visual reading
 * order, and layout / images / charts are lost — use the `*-as-pdf` sibling +
 * a vision model when the rendered deck matters.
 */

const slideSection = (slide: Slide, index: number): string => {
  const parts = [`## Slide ${index + 1}${slide.hidden ? ' (hidden)' : ''}`];
  if (slide.text.trim() !== '') parts.push(slide.text);
  if (slide.notes.trim() !== '') parts.push(`**Speaker notes:** ${slide.notes}`);
  return parts.join('\n\n');
};

const pptxToMarkdown = async (bytes: Uint8Array, options: { readonly includeMetadata?: boolean } = {}): Promise<Result<MarkdownEnvelope, GraphError>> => {
  const meta = await extractPptxMetadata(bytes);
  if (!meta.ok) return meta;
  const body = meta.value.slides.map(slideSection).join('\n\n');
  const text = options.includeMetadata === true ? `${body}\n\n${formatPptxMetadata(meta.value)}` : body;
  // size = UTF-8 byte count (audit §2.1); `text.length` is UTF-16 code units.
  return ok({ contentType: 'text/markdown', size: new TextEncoder().encode(text).byteLength, text });
};

export { pptxToMarkdown };
