import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { htmlToMarkdown } from '../../infra/turndown-adapter.ts';
import { embedInlineImages, type InlineAttachment } from './inline-image-embedder.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({ messageId: z.string().min(1) });

// Audit v1.0.0 — multi-MB attachments timeout fix. `?$expand=attachments`
// inlines every attachment's `contentBytes` (base64) into the message
// envelope. For an email with a 4 MB PDF attachment the response balloons
// past Graph's ~3 MB tolerance; Graph times out at 60s or truncates the
// JSON mid-stream. We now stage the fetch:
//   1. /me/messages/{id}                     (no $expand)         — body + hasAttachments
//   2. /me/messages/{id}/attachments?$select (metadata only)      — only if hasAttachments
//   3. /me/messages/{id}/attachments/{a-id}  (per inline image)   — only for small inline images
// File attachments are listed in the markdown by name + size + id (so the
// caller can fetch them on demand via `convert-mail-attachment-to-pdf` or
// `get-mail-attachment`); their bytes never traverse this command.

const INLINE_IMAGE_SIZE_LIMIT_BYTES = 2_000_000;

const ATTACHMENT_METADATA_SELECT = '$select=id,name,contentType,size,isInline,contentId';

// Single predicate replacing the `typeof x === 'string' && x !== ''` pattern
// that was repeated across attachment field checks. Collapses ~5 separate
// 4-mutant predicates into one place — the helper itself remains under test
// via every call site's behaviour.
const nonEmpty = (v: unknown): v is string => typeof v === 'string' && v !== '';

type Recipient = { readonly emailAddress?: { readonly name?: string; readonly address?: string } };

const formatAddress = (a: { readonly name?: string; readonly address?: string } | undefined): string | undefined => {
  if (!nonEmpty(a?.address)) return undefined;
  return nonEmpty(a?.name) ? `${a.name} <${a.address}>` : a.address;
};

const formatRecipients = (rs: ReadonlyArray<Recipient> | undefined): string | undefined => {
  if (!rs || rs.length === 0) return undefined;
  const parts = rs.map((r) => formatAddress(r.emailAddress)).filter(nonEmpty);
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
  if (nonEmpty(m.receivedDateTime)) lines.push(`**Date:** ${m.receivedDateTime}`);
  return lines.join('\n');
};

type AttachmentMeta = {
  readonly id?: string;
  readonly name?: string;
  readonly contentType?: string;
  readonly size?: number;
  readonly isInline?: boolean;
  readonly contentId?: string;
};

// Decimal units (1 KB = 1000 bytes, 1 MB = 1_000_000 bytes) — matches
// Outlook / Microsoft 365 user-facing size displays for email attachments.
const formatBytes = (n: number): string => {
  if (n < 1000) return `${n} B`;
  if (n < 1_000_000) return `${(n / 1000).toFixed(1)} KB`;
  return `${(n / 1_000_000).toFixed(1)} MB`;
};

const isInlineImage = (a: AttachmentMeta): boolean => a.isInline === true && nonEmpty(a.contentType) && a.contentType.toLowerCase().startsWith('image/') && nonEmpty(a.contentId);

type EmbedFetchResult = { readonly meta: AttachmentMeta; readonly inline?: InlineAttachment; readonly oversize: boolean };

const fetchInlineImageBytes = async (graph: GraphClient, messageId: string, meta: AttachmentMeta): Promise<EmbedFetchResult> => {
  if ((meta.size ?? 0) > INLINE_IMAGE_SIZE_LIMIT_BYTES) return { meta, oversize: true };
  if (!nonEmpty(meta.id)) return { meta, oversize: false };
  const fetched = await graph.get(`/me/messages/${messageId}/attachments/${meta.id}`);
  if (!fetched.ok) return { meta, oversize: false };
  const body = fetched.value as { readonly contentBytes?: string };
  if (!nonEmpty(body.contentBytes)) return { meta, oversize: false };
  return {
    meta,
    oversize: false,
    inline: { contentId: meta.contentId ?? '', contentType: meta.contentType ?? '', contentBytes: body.contentBytes },
  };
};

const renderOversizePlaceholders = (html: string, embeds: ReadonlyArray<EmbedFetchResult>): string => {
  let out = html;
  for (const e of embeds) {
    if (!e.oversize || !nonEmpty(e.meta.contentId)) continue;
    const label = `[inline image too large to embed: ${e.meta.name ?? 'image'} (${formatBytes(e.meta.size ?? 0)})]`;
    out = out.replaceAll(`cid:${e.meta.contentId}`, label);
  }
  return out;
};

const renderFileAttachmentsList = (attachments: ReadonlyArray<AttachmentMeta>): string => {
  const fileAttachments = attachments.filter((a) => !isInlineImage(a) && nonEmpty(a.name));
  if (fileAttachments.length === 0) return '';
  const items = fileAttachments.map((a) => {
    const size = typeof a.size === 'number' ? ` (${formatBytes(a.size)}` : '';
    const type = nonEmpty(a.contentType) ? `, ${a.contentType}` : '';
    const id = nonEmpty(a.id) ? `, id: ${a.id}` : '';
    return `- ${a.name ?? ''}${size}${type}${id})`;
  });
  return ['**Attachments:**', ...items, '_Use `convert-mail-attachment-to-pdf` or `get-mail-attachment` with the attachment id to fetch._'].join('\n');
};

const execute = async (graph: GraphClient, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { messageId } = parsed.data;

  const fetched = await graph.get(`/me/messages/${messageId}`);
  if (!fetched.ok) return fetched;

  const m = fetched.value as {
    readonly subject?: string;
    readonly from?: Recipient;
    readonly toRecipients?: ReadonlyArray<Recipient>;
    readonly ccRecipients?: ReadonlyArray<Recipient>;
    readonly receivedDateTime?: string;
    readonly body?: { readonly contentType?: string; readonly content?: string };
    readonly hasAttachments?: boolean;
  };

  let attachments: ReadonlyArray<AttachmentMeta> = [];
  let attachmentsListNote: string | undefined;
  if (m.hasAttachments === true) {
    const listed = await graph.get(`/me/messages/${messageId}/attachments?${ATTACHMENT_METADATA_SELECT}`);
    if (listed.ok) {
      attachments = ((listed.value as { readonly value?: ReadonlyArray<AttachmentMeta> }).value ?? []) as ReadonlyArray<AttachmentMeta>;
    } else {
      attachmentsListNote = `attachments-list fetch failed (${listed.error.type}: ${listed.error.message}) — markdown body returned without attachment metadata`;
    }
  }

  const inlineImageCandidates = attachments.filter(isInlineImage);
  const embedResults = await Promise.all(inlineImageCandidates.map((meta) => fetchInlineImageBytes(graph, messageId, meta)));
  const inlineImages = embedResults.flatMap((r) => (r.inline ? [r.inline] : []));

  const headers = renderHeaders(m);
  const rawHtml = m.body?.content ?? '';
  const withPlaceholders = renderOversizePlaceholders(rawHtml, embedResults);
  const inlined = inlineImages.length > 0 ? embedInlineImages(withPlaceholders, inlineImages) : withPlaceholders;
  let bodyMd: string;
  if (m.body?.contentType === 'html') {
    const converted = htmlToMarkdown(inlined);
    if (!converted.ok) return converted;
    bodyMd = converted.value;
  } else {
    bodyMd = inlined;
  }
  const fileList = renderFileAttachmentsList(attachments);
  const text = [headers, bodyMd, fileList].filter((s) => s !== '').join('\n\n');

  // size = UTF-8 byte count (audit §2.1); `text.length` is UTF-16 code units.
  const envelope: { contentType: string; size: number; text: string; note?: string } = {
    contentType: 'text/markdown',
    size: new TextEncoder().encode(text).byteLength,
    text,
  };
  if (attachmentsListNote !== undefined) envelope.note = attachmentsListNote;
  return ok(envelope);
};

const meta: CommandMeta = {
  summary:
    'Render a single Outlook email as markdown — headers (`**Subject:**`, `**From:**`, `**To:**`, `**Cc:**` only when present, `**Date:**`), followed by the body run through turndown. Inline images attached with `isInline:true` and an `image/*` content-type (size ≤ 2 MB) are embedded as base64 `data:` URIs so the output is self-contained (Hardening #1: non-image inline attachments are NOT embedded; oversize inline images are replaced with a placeholder note). File attachments are listed below the body by name + size + id; their bytes are NOT fetched here — call `convert-mail-attachment-to-pdf` or `get-mail-attachment` with the id when you actually need them. Staged-fetch design (audit v1.0.0): one call for the body, one for the attachments-metadata list (only if `hasAttachments:true`), and one per small inline image — replaces the old `?$expand=attachments` which timed out / truncated on messages with multi-MB attachments.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-get',
  options: [{ name: 'message-id', key: 'messageId', required: true, description: 'Outlook message ID. Returned by `list-mail-messages` or `list-mail-folder-messages`.' }],
  example: "ask-marcel convert-mail-to-markdown --message-id 'AAMkAD...'",
  responseShape:
    '`{ contentType: "text/markdown", size, text, note? }` — headers + turndown-rendered body + (when present) a file-attachments list. The optional `note` carries a partial-success hint when the attachments-metadata fetch fails after the body succeeded.',
  producesBytes: true,
};

export { execute, meta, schema };
