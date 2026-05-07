import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import { htmlToMarkdown } from '../../infra/turndown-adapter.ts';
import { embedInlineImages, type InlineAttachment } from './inline-image-embedder.ts';

/**
 * Orchestrate the four steps that turn a Graph `?format=html` content
 * call into a markdown envelope:
 *
 *   1. getBinary(contentPath)
 *   2. if Graph returned a 302 downloadUrl, follow via fetchUrl (host
 *      allow-list enforced inside fetchUrl, Hardening #3)
 *   3. embedInlineImages over any cid: refs (Hardening #1: image/* only)
 *   4. htmlToMarkdown via turndown
 *
 * Returns `{ contentType: 'text/markdown', size, text }` on success.
 */

type MarkdownEnvelope = { readonly contentType: 'text/markdown'; readonly size: number; readonly text: string };

/**
 * Microsoft Graph's `?format=html` conversion supports a *narrow* input
 * set: only `loop`, `fluid`, `wbtx`, and `whiteboard` source extensions
 * (per https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format).
 * Office documents (docx, pptx, xlsx, pdf, rtf, etc.) are NOT in that set
 * and cause Office Online's conversion sandbox to return
 * `Sandbox_InputFormatNotSupported`. This has been the documented
 * behavior since the API shipped — it is not an outage.
 *
 * `format=pdf` has a much wider input set (38 extensions including all
 * Office formats), so the right user-facing fix is to redirect them to
 * the `*-as-pdf` sibling.
 */
const augmentSandboxError = (e: GraphError): GraphError => {
  if (e.type !== 'api_error' || !e.message.startsWith('Sandbox_InputFormatNotSupported')) return e;
  return {
    type: 'api_error',
    status: e.status,
    message:
      'Microsoft Graph `?format=html` only supports four input extensions: loop, fluid, wbtx, whiteboard (https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format). Office documents (docx, pptx, xlsx, pdf) are NOT in that set and Office Online correctly rejects them with Sandbox_InputFormatNotSupported. For Office sources, use the corresponding `*-as-pdf` command — `?format=pdf` accepts a much wider input set.',
  };
};

const extractHtml = async (graph: GraphClient, binaryResult: Record<string, unknown>): Promise<Result<string, GraphError>> => {
  if (typeof binaryResult.text === 'string') {
    return ok(binaryResult.text);
  }
  if (typeof binaryResult['@microsoft.graph.downloadUrl'] === 'string') {
    const followed = await graph.fetchUrl(binaryResult['@microsoft.graph.downloadUrl']);
    if (!followed.ok) return followed;
    const v = followed.value as Record<string, unknown>;
    if (typeof v.text === 'string') return ok(v.text);
    return err({ type: 'api_error', status: 500, message: 'unexpected response shape from fetchUrl: missing text field' });
  }
  return err({ type: 'api_error', status: 500, message: 'unexpected response shape from getBinary: no text or downloadUrl' });
};

const convertToMarkdown = async (
  graph: GraphClient,
  contentPath: string,
  inlineAttachments: ReadonlyArray<InlineAttachment> = []
): Promise<Result<MarkdownEnvelope, GraphError>> => {
  const binary = await graph.getBinary(contentPath);
  if (!binary.ok) return err(augmentSandboxError(binary.error));
  const html = await extractHtml(graph, binary.value as Record<string, unknown>);
  if (!html.ok) return err(augmentSandboxError(html.error));
  const inlined = inlineAttachments.length > 0 ? embedInlineImages(html.value, inlineAttachments) : html.value;
  const md = htmlToMarkdown(inlined);
  if (!md.ok) return md;
  return ok({ contentType: 'text/markdown', size: md.value.length, text: md.value });
};

export { convertToMarkdown };
export type { MarkdownEnvelope };
