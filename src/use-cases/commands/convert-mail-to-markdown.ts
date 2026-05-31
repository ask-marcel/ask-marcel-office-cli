import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandMeta } from './command-types.ts';
import { htmlToMarkdown } from '../../infra/turndown-adapter.ts';
import { embedInlineImages, type InlineAttachment } from './inline-image-embedder.ts';
import { formatZodError } from './format-zod-error.ts';
import { stripQuotedReplies } from './mail-quote-stripper.ts';

// Audit v1.4.0 fresh-pass #6: `--inline-images false` opts out of the
// per-image bytes-fetch + base64 embedding. An LLM that just wants the text
// body of an email with several inline images was paying ~36 KB for a 6 KB
// body. With `--inline-images false`, the body is rendered with the raw
// `cid:<contentId>` references preserved (no `data:` URIs); the file-
// attachments list still surfaces the inline images so the caller knows
// they exist. Default stays `true` for backward compatibility.
const schema = z.object({
  messageId: z.string().min(1),
  inlineImages: z.enum(['true', 'false']).optional(),
  keepQuoted: z.enum(['true', 'false']).optional(),
});

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

// Audit Alex-session §2 follow-up: `contentId` only exists on the
// `microsoft.graph.fileAttachment` subtype, NOT on the base
// `microsoft.graph.attachment`. The `/me/messages/{id}/attachments`
// endpoint returns polymorphic entries (fileAttachment | itemAttachment |
// referenceAttachment); requesting `contentId` bare returns
// `Could not find a property named contentId on type microsoft.graph.attachment`
// and Graph fails the whole list-fetch. The CLI used to swallow this in
// a `note` field, dropping every attachment's metadata.
//
// Graph's polymorphic-cast syntax for $select on derived types is
// `microsoft.graph.<derived-type>/<field>` — the cast applies the field
// projection only to entries of that subtype, leaving other subtypes
// unaffected. Documented in Graph OData / cast operator reference.
const ATTACHMENT_METADATA_SELECT = '$select=id,name,contentType,size,isInline,microsoft.graph.fileAttachment/contentId';

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

// Schema validates the attachments-list Graph response at the boundary.
// Without this, a malformed shape (e.g., `{ value: "not an array" }` from a
// tenant glitch) launders through the `as` cast and throws TypeError
// downstream on `.filter()`. The schema-failure path surfaces a precise
// note in the markdown envelope instead of an unhandled exception.
//
// Regression note (v1.4.0 follow-up): every field uses `.nullish()`
// (= `.optional().nullable()`) rather than `.optional()`. Graph's
// polymorphic-cast response (`microsoft.graph.fileAttachment/contentId`)
// returns `contentId: null` on every non-fileAttachment entry — `.optional()`
// rejects `null`, which made the schema fail and the "malformed shape" note
// fire on every real call. The downstream `nonEmpty` predicate already
// treats `null` as "empty" (it requires `typeof v === 'string'`), so
// loosening the input type is safe and matches the wire reality.
const attachmentMetaSchema = z.object({
  id: z.string().nullish(),
  name: z.string().nullish(),
  contentType: z.string().nullish(),
  size: z.number().nullish(),
  isInline: z.boolean().nullish(),
  contentId: z.string().nullish(),
});

const attachmentsListSchema = z.object({
  value: z.array(attachmentMetaSchema).optional(),
});

type AttachmentMeta = z.infer<typeof attachmentMetaSchema>;

// Decimal units (1 KB = 1000 bytes, 1 MB = 1_000_000 bytes) — matches
// Outlook / Microsoft 365 user-facing size displays for email attachments.
const formatBytes = (n: number): string => {
  if (n < 1000) return `${n} B`;
  if (n < 1_000_000) return `${(n / 1000).toFixed(1)} KB`;
  return `${(n / 1_000_000).toFixed(1)} MB`;
};

// Narrowed shape produced by `isInlineImage`-filtered attachments: contentType
// and contentId are guaranteed non-empty by that predicate, so downstream code
// doesn't need defensive `?? ''` defaults. Documents the invariant in the type
// system instead of in a comment.
type InlineImageCandidate = AttachmentMeta & { readonly contentType: string; readonly contentId: string };

const isInlineImage = (a: AttachmentMeta): a is InlineImageCandidate =>
  a.isInline === true && nonEmpty(a.contentType) && a.contentType.toLowerCase().startsWith('image/') && nonEmpty(a.contentId);

type EmbedFetchResult = { readonly meta: InlineImageCandidate; readonly inline?: InlineAttachment; readonly oversize: boolean };

const fetchInlineImageBytes = async (graph: GraphClient, messageId: string, meta: InlineImageCandidate): Promise<EmbedFetchResult> => {
  if ((meta.size ?? 0) > INLINE_IMAGE_SIZE_LIMIT_BYTES) return { meta, oversize: true };
  if (!nonEmpty(meta.id)) return { meta, oversize: false };
  const fetched = await graph.get(`/me/messages/${messageId}/attachments/${meta.id}`);
  if (!fetched.ok) return { meta, oversize: false };
  const body = fetched.value as { readonly contentBytes?: string };
  if (!nonEmpty(body.contentBytes)) return { meta, oversize: false };
  return {
    meta,
    oversize: false,
    inline: { contentId: meta.contentId, contentType: meta.contentType, contentBytes: body.contentBytes },
  };
};

const renderOversizePlaceholders = (html: string, embeds: ReadonlyArray<EmbedFetchResult>): string => {
  let out = html;
  for (const e of embeds) {
    if (!e.oversize) continue;
    const label = `[inline image too large to embed: ${e.meta.name ?? 'image'} (${formatBytes(e.meta.size ?? 0)})]`;
    out = out.replaceAll(`cid:${e.meta.contentId}`, label);
  }
  return out;
};

// `includeInlineImages` is set when CMtM is in `--inline-images false` mode:
// the body keeps raw `cid:<contentId>` refs (no data-URI embedding), so the
// caller would lose visibility of inline images entirely if the attachment
// list still filtered them out. Surface them alongside regular file
// attachments so the LLM can decide whether to fetch the bytes separately.
const renderFileAttachmentsList = (attachments: ReadonlyArray<AttachmentMeta>, includeInlineImages: boolean = false): string => {
  const fileAttachments = attachments.filter((a) => (includeInlineImages || !isInlineImage(a)) && nonEmpty(a.name));
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
  // Default `true` keeps backward compat: existing callers see the same
  // base64-embedded output. Pass `--inline-images false` to skip the per-
  // image bytes fetch + embedding (cheaper for an LLM that just wants text).
  const embedInlineImagesEnabled = parsed.data.inlineImages !== 'false';
  // Quoted reply chains / forwarded-message blocks are stripped by default
  // (they duplicate content already in earlier messages and blow the context
  // budget). Pass `--keep-quoted true` to preserve the full body.
  const keepQuoted = parsed.data.keepQuoted === 'true';

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
    if (!listed.ok) {
      attachmentsListNote = `attachments-list fetch failed (${listed.error.type}: ${listed.error.message}) — markdown body returned without attachment metadata`;
    } else {
      const parsed = attachmentsListSchema.safeParse(listed.value);
      if (parsed.success) {
        attachments = parsed.data.value ?? [];
      } else {
        attachmentsListNote = `attachments-list returned a malformed shape (${formatZodError(parsed.error)}) — markdown body returned without attachment metadata`;
      }
    }
  }

  // When inline-image embedding is disabled, skip the per-image bytes
  // fetch AND the oversize-placeholder pass — the body keeps its raw
  // `cid:<contentId>` references, and the file-attachments list (rendered
  // below) still surfaces the inline images by name + id so the caller
  // knows they exist.
  const inlineImageCandidates = embedInlineImagesEnabled ? attachments.filter(isInlineImage) : [];
  const embedResults = await Promise.all(inlineImageCandidates.map((meta) => fetchInlineImageBytes(graph, messageId, meta)));
  const inlineImages = embedResults.flatMap((r) => (r.inline ? [r.inline] : []));

  const headers = renderHeaders(m);
  const rawHtml = m.body?.content ?? '';
  const withPlaceholders = renderOversizePlaceholders(rawHtml, embedResults);
  const inlined = inlineImages.length > 0 ? embedInlineImages(withPlaceholders, inlineImages) : withPlaceholders;
  let bodyMd: string;
  let quotedStripped = false;
  if (m.body?.contentType === 'html') {
    // Strip quoted reply chains before turndown (the Outlook / Gmail markers
    // are still HTML at this point). v1 covers HTML bodies only — plain-text
    // reply detection is a separate heuristic.
    const stripped = keepQuoted ? { html: inlined, stripped: false } : stripQuotedReplies(inlined);
    quotedStripped = stripped.stripped;
    const converted = htmlToMarkdown(stripped.html);
    if (!converted.ok) return converted;
    bodyMd = converted.value;
  } else {
    bodyMd = inlined;
  }
  const fileList = renderFileAttachmentsList(attachments, !embedInlineImagesEnabled);
  const text = [headers, bodyMd, fileList].filter((s) => s !== '').join('\n\n');

  // size = UTF-8 byte count (audit §2.1); `text.length` is UTF-16 code units.
  const envelope: { contentType: string; size: number; text: string; note?: string } = {
    contentType: 'text/markdown',
    size: new TextEncoder().encode(text).byteLength,
    text,
  };
  const notes: string[] = [];
  if (attachmentsListNote !== undefined) notes.push(attachmentsListNote);
  if (quotedStripped) notes.push('quoted reply chain stripped — pass --keep-quoted true to include it');
  if (notes.length > 0) envelope.note = notes.join('; ');
  return ok(envelope);
};

const meta: CommandMeta = {
  summary:
    'Render a single Outlook email as markdown — headers (`**Subject:**`, `**From:**`, `**To:**`, `**Cc:**` only when present, `**Date:**`), followed by the body run through turndown. By default, inline images (`isInline:true` + `image/*` content-type, size ≤ 2 MB) are embedded as base64 `data:` URIs so the output is self-contained (non-image inline attachments are NOT embedded; oversize inline images are replaced with a placeholder note). For LLM callers that only want the text body, pass `--inline-images false` to skip the per-image bytes fetch entirely — the body keeps raw `cid:<contentId>` references and the inline images surface in the file-attachments list so you can decide whether to fetch them separately via `get-mail-attachment`. File attachments are always listed below the body by name + size + id; their bytes are NOT fetched here — call `convert-mail-attachment-to-pdf` or `get-mail-attachment` with the id when you actually need them. Staged-fetch design (audit v1.0.0): one call for the body, one for the attachments-metadata list (only if `hasAttachments:true`), and one per small inline image — replaces the old `?$expand=attachments` which timed out / truncated on messages with multi-MB attachments.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-get',
  options: [
    { name: 'message-id', key: 'messageId', required: true, description: 'Outlook message ID. Returned by `list-mail-messages` or `list-mail-folder-messages`.' },
    {
      name: 'inline-images',
      key: 'inlineImages',
      required: false,
      description:
        'Pass `--inline-images false` to skip the per-image bytes fetch + base64 embedding. Default is `true` (embed). Disabling cuts the response size dramatically on emails with several inline images (a 6 KB body with 6 inline images shipped at ~36 KB by default; with `--inline-images false` it stays close to 6 KB). The body keeps raw `cid:<contentId>` references and the inline images surface in the file-attachments list instead, so the LLM caller can still see what is there and fetch any specific image via `get-mail-attachment` on demand.',
    },
    {
      name: 'keep-quoted',
      key: 'keepQuoted',
      required: false,
      description:
        'Quoted reply chains and forwarded-message blocks are stripped by default (they duplicate content already present in earlier messages and inflate the context budget). The stripped tail is replaced with a single visible marker naming this flag, so nothing is removed silently. Pass `--keep-quoted true` to preserve the full body. Only well-known structural markers are cut (Outlook `divRplyFwdMsg` / `appendonsend` / `stopSpelling`, Gmail `gmail_quote`); HTML bodies only.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: "ask-marcel convert-mail-to-markdown --message-id 'AAMkAD...'",
  responseShape:
    '`{ contentType: "text/markdown", size, text, note? }` — headers + turndown-rendered body + (when present) a file-attachments list. The optional `note` carries a partial-success hint when the attachments-metadata fetch fails after the body succeeded, and/or a flag that a quoted reply chain was stripped (use `--keep-quoted true` to include it).',
  producesBytes: true,
};

export { execute, meta, schema };
