import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import { extractEml } from '../../infra/eml-parser-adapter.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { renderMsg } from './msg-to-markdown.ts';
import type { MsgAttachmentConverter, MsgToMarkdownOptions } from './msg-to-markdown.ts';

/**
 * Render a raw RFC 822 message (`.eml`) exactly as an Outlook `.msg` renders: the
 * parser maps it onto the same shape, so the subject, the header block, the
 * quote stripping (`keepQuoted`), the `cid:` placeholders and the recursive
 * attachments are the `.msg` renderer's own. `recurse` and `depth` work as in
 * `msgToMarkdown`.
 */
const emlToMarkdown = async (bytes: Uint8Array, opts: MsgToMarkdownOptions, recurse: MsgAttachmentConverter): Promise<Result<unknown, GraphError>> => {
  const parsed = await extractEml(bytes);
  if (!parsed.ok) return parsed;
  const text = await renderMsg(parsed.value, opts, recurse);
  return ok({ contentType: 'text/markdown', size: new TextEncoder().encode(text).byteLength, text });
};

export { emlToMarkdown };
