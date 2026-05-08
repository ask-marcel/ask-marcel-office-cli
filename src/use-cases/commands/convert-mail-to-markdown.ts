import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { htmlToMarkdown } from '../../infra/turndown-adapter.ts';
import { embedInlineImages, type InlineAttachment } from './inline-image-embedder.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({ messageId: z.string().min(1) });

type Recipient = { readonly emailAddress?: { readonly name?: string; readonly address?: string } };

const formatAddress = (a: { readonly name?: string; readonly address?: string } | undefined): string | undefined => {
  if (!a?.address) return undefined;
  return a.name ? `${a.name} <${a.address}>` : a.address;
};

const formatRecipients = (rs: ReadonlyArray<Recipient> | undefined): string | undefined => {
  if (!rs || rs.length === 0) return undefined;
  const parts = rs.map((r) => formatAddress(r.emailAddress)).filter((s): s is string => s !== undefined);
  return parts.length > 0 ? parts.join(', ') : undefined;
};

const renderHeaders = (m: {
  readonly subject?: string;
  readonly from?: Recipient;
  readonly toRecipients?: ReadonlyArray<Recipient>;
  readonly ccRecipients?: ReadonlyArray<Recipient>;
  readonly receivedDateTime?: string;
}): string => {
  const lines: string[] = [];
  if (m.subject !== undefined) lines.push(`**Subject:** ${m.subject}`);
  const from = formatAddress(m.from?.emailAddress);
  if (from !== undefined) lines.push(`**From:** ${from}`);
  const to = formatRecipients(m.toRecipients);
  if (to !== undefined) lines.push(`**To:** ${to}`);
  const cc = formatRecipients(m.ccRecipients);
  if (cc !== undefined) lines.push(`**Cc:** ${cc}`);
  if (m.receivedDateTime !== undefined) lines.push(`**Date:** ${m.receivedDateTime}`);
  return lines.join('\n');
};

type AttachmentLike = {
  readonly '@odata.type'?: string;
  readonly contentId?: string;
  readonly contentType?: string;
  readonly contentBytes?: string;
  readonly isInline?: boolean;
};

const collectInlineImageAttachments = (attachments: ReadonlyArray<AttachmentLike> | undefined): ReadonlyArray<InlineAttachment> => {
  if (!attachments) return [];
  return attachments
    .filter(
      (a) =>
        a.isInline === true &&
        typeof a.contentBytes === 'string' &&
        a.contentBytes !== '' &&
        typeof a.contentType === 'string' &&
        a.contentType.toLowerCase().startsWith('image/') &&
        typeof a.contentId === 'string' &&
        a.contentId !== ''
    )
    .map(
      (a): InlineAttachment => ({
        contentId: a.contentId as string,
        contentType: a.contentType as string,
        contentBytes: a.contentBytes as string,
      })
    );
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { messageId } = parsed.data;

  const fetched = await graph.get(`/me/messages/${messageId}?$expand=attachments`);
  if (!fetched.ok) return fetched;

  const m = fetched.value as {
    readonly subject?: string;
    readonly from?: Recipient;
    readonly toRecipients?: ReadonlyArray<Recipient>;
    readonly ccRecipients?: ReadonlyArray<Recipient>;
    readonly receivedDateTime?: string;
    readonly body?: { readonly contentType?: string; readonly content?: string };
    readonly attachments?: ReadonlyArray<AttachmentLike>;
  };

  const headers = renderHeaders(m);
  const inlineImages = collectInlineImageAttachments(m.attachments);
  const rawHtml = m.body?.content ?? '';
  const inlined = inlineImages.length > 0 ? embedInlineImages(rawHtml, inlineImages) : rawHtml;
  let bodyMd: string;
  if (m.body?.contentType === 'html') {
    const converted = htmlToMarkdown(inlined);
    if (!converted.ok) return converted;
    bodyMd = converted.value;
  } else {
    bodyMd = inlined;
  }
  const text = [headers, bodyMd].filter((s) => s !== '').join('\n\n');

  // size = UTF-8 byte count (audit §2.1); `text.length` is UTF-16 code units.
  return ok({ contentType: 'text/markdown', size: new TextEncoder().encode(text).byteLength, text });
};

const meta: CommandMeta = {
  summary:
    'Render a single Outlook email as markdown — headers in this order: `**Subject:**`, `**From:**`, `**To:**`, `**Cc:**` (only when present), `**Date:**` — followed by the body run through turndown. Inline images attached with `isInline:true` and an `image/*` content-type are embedded as base64 `data:` URIs so the output is self-contained (Hardening #1: non-image inline attachments are NOT embedded). One Graph round-trip via `?$expand=attachments`.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-get',
  options: [{ name: 'message-id', key: 'messageId', required: true, description: 'Outlook message ID. Returned by `list-mail-messages` or `list-mail-folder-messages`.' }],
  example: "ask-marcel convert-mail-to-markdown --message-id 'AAMkAD...'",
  responseShape: '`{ contentType: "text/markdown", size, text }` — headers + turndown-rendered body with inline images embedded.',
};

export { execute, meta, schema };
