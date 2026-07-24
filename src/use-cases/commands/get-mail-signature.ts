import { z } from 'zod';
import { err, ok, type Result } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { ATTACHMENT_METADATA_SELECT, attachmentsListSchema, fetchInlineImageBytes, formatBytes, isInlineImage } from './convert-mail-to-markdown.ts';
import { formatZodError } from './format-zod-error.ts';
import { embedInlineImages } from './inline-image-embedder.ts';
import { extractSignatureBlock } from './signature-extractor.ts';

const schema = z.object({
  messageId: z.string().optional(),
});

const SCAN_LIMIT = 10;
// The space in `$orderby` is percent-encoded by hand: this path is hardcoded,
// so there is no builder to do it.
const SENT_SCAN_PATH = `/me/mailFolders/sentitems/messages?$top=${SCAN_LIMIT}&$orderby=sentDateTime%20desc&$select=id,sentDateTime`;

const sentListSchema = z.object({ value: z.array(z.object({ id: z.string() })).optional() });
const messageSchema = z.object({
  body: z.object({ contentType: z.string(), content: z.string() }).optional(),
  sentDateTime: z.string().optional(),
  hasAttachments: z.boolean().optional(),
});

type InlineResult = { readonly html: string; readonly count: number; readonly note?: string };

/**
 * Embeds the inline images the signature actually references. Deliberately does
 * NOT run `replaceUnresolvedCidImages`: a text placeholder is right for a body
 * being read, but this block is destined for a draft, and swapping the img for
 * `[inline image: logo]` would destroy the reference the caller could still
 * resolve. Anything not embedded keeps its cid: and is named in the note.
 */
const inlineSignatureImages = async (graph: GraphClient, messageId: string, block: string): Promise<InlineResult> => {
  const listed = await graph.get(`/me/messages/${messageId}/attachments?${ATTACHMENT_METADATA_SELECT}`);
  if (!listed.ok) return { html: block, count: 0, note: 'inline images could not be listed, so any cid: references are unresolved' };
  const parsed = attachmentsListSchema.safeParse(listed.value);
  if (!parsed.success) return { html: block, count: 0, note: 'the inline-image list came back in an unreadable shape, so any cid: references are unresolved' };

  const referenced = (parsed.data.value ?? []).filter(isInlineImage).filter((a) => block.includes(`cid:${a.contentId}`));
  const fetched = await Promise.all(referenced.map((meta) => fetchInlineImageBytes(graph, messageId, meta)));
  const embeddable = fetched.flatMap((f) => (f.inline === undefined ? [] : [f.inline]));
  const skipped = fetched.filter((f) => f.inline === undefined);

  const html = embeddable.length > 0 ? embedInlineImages(block, embeddable) : block;
  if (skipped.length === 0) return { html, count: embeddable.length };
  const named = skipped.map((f) => `${f.meta.name ?? 'image'} (${formatBytes(f.meta.size ?? 0)})`).join(', ');
  return { html, count: embeddable.length, note: `${skipped.length} inline image${skipped.length === 1 ? '' : 's'} left as a cid: reference: ${named}` };
};

// A signature references its logo as `<img src="cid:...">`. The single attachments
// call is gated on THIS, not on Graph's `hasAttachments`: an inline image carries
// `isInline` + a contentId but does NOT flip `hasAttachments`, so a signature whose
// only images are inline rides in a message reported `hasAttachments:false`. Gating
// on the flag skipped the embed and left the logo a raw cid: reference — broken
// once pasted into a fresh draft (reported 2026-07-19).
const referencesCidImage = (block: string): boolean => /\bsrc\s*=\s*["']cid:/i.test(block);

const noSignatureFound = (messageId: string | undefined, scanned: number): GraphError => {
  const plural = scanned === 1 ? '' : 's';
  const scanMessage = `No OWA signature block (\`<div id="Signature">\`) was found in the last ${scanned} sent message${plural}. Mail composed in Outlook desktop does not carry the marker, so a signature may exist without being findable this way - pass --message-id to pin a message you know was sent from Outlook on the web.`;
  const pinnedMessage = `Message ${messageId} carries no OWA signature block (\`<div id="Signature">\`). Mail composed in Outlook desktop does not carry the marker; pin a message sent from Outlook on the web instead.`;
  return { type: 'validation_error', message: messageId === undefined ? scanMessage : pinnedMessage };
};

const buildEnvelope = (messageId: string, block: string, sentDateTime: string | undefined, inlined: InlineResult): Record<string, unknown> => ({
  contentType: 'text/html',
  // UTF-8 bytes; `text.length` would be UTF-16 code units.
  size: new TextEncoder().encode(inlined.html).byteLength,
  text: inlined.html,
  sourceMessageId: messageId,
  ...(sentDateTime === undefined ? {} : { sentDateTime }),
  inlinedImages: inlined.count,
  ...(inlined.note === undefined ? {} : { note: inlined.note }),
});

const readCandidates = async (graph: GraphClient, messageId: string | undefined): Promise<Result<ReadonlyArray<string>, GraphError>> => {
  if (messageId !== undefined) return ok([messageId]);
  const listed = await graph.get(SENT_SCAN_PATH);
  if (!listed.ok) return listed;
  const parsed = sentListSchema.safeParse(listed.value);
  const ids = (parsed.success ? (parsed.data.value ?? []) : []).map((m) => m.id);
  if (ids.length === 0)
    return err({ type: 'validation_error', message: 'Found no sent messages to read a signature from. Send one from Outlook on the web first, or pass --message-id.' });
  return ok(ids);
};

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { messageId } = parsed.data;

  const candidates = await readCandidates(graph, messageId);
  if (!candidates.ok) return candidates;

  // Sequential on purpose: the newest sent message almost always carries the
  // signature, so this is one body read in the common case. Fanning out over ten
  // would cost ten reads every time to save latency the caller rarely needs.
  for (const id of candidates.value) {
    const fetched = await graph.get(`/me/messages/${id}?$select=body,sentDateTime,hasAttachments`);
    if (!fetched.ok) return fetched;
    const message = messageSchema.safeParse(fetched.value);
    if (!message.success || message.data.body === undefined) continue;
    const block = extractSignatureBlock(message.data.body.content);
    if (block === undefined) continue;
    const inlined = referencesCidImage(block) ? await inlineSignatureImages(graph, id, block) : { html: block, count: 0 };
    return ok(buildEnvelope(id, block, message.data.sentDateTime, inlined));
  }
  return err(noSignatureFound(messageId, candidates.value.length));
};

const meta: CommandMeta = {
  summary:
    'Read your own email signature as HTML, lifted from a message you already sent. Graph-created drafts carry NO signature (create-mail-draft, create-reply-draft, and create-forward-draft all produce unsigned bodies), so this is where you get one: take the `text` this returns, append it to your reply text, and hand the result to `update-mail-draft` in comment mode (HTML body-content-type) to place it above the quoted history. Scans your last 10 sent messages newest-first and returns the first `<div id="Signature">` block it finds, stopping there, with any logo the block references embedded as a base64 data: URI so the HTML renders on its own. Read-only. NOTE: the marker is written by Outlook on the web and new Outlook; mail composed in Outlook desktop does not carry it, so pin a webmail-sent message with --message-id if the scan finds nothing.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate:
    '/me/mailFolders/sentitems/messages (scan, skipped when {message-id} is given) then /me/messages/{message-id}?$select=body,sentDateTime,hasAttachments (+ /attachments per referenced logo)',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-get',
  options: [
    {
      name: 'message-id',
      key: 'messageId',
      required: false,
      description:
        'Read the signature from THIS message instead of scanning the sent folder. Use when the scan finds nothing (the message was composed in Outlook desktop) or to pin a specific signature. Source from list-mail-folder-messages --mail-folder-id sentitems.',
      argumentHint: { kind: 'idOrName' },
    },
  ],
  example: 'ask-marcel-office get-mail-signature',
  producesBytes: true,
  responseShape:
    '`{ contentType: "text/html", size, text, sourceMessageId, sentDateTime?, inlinedImages, note? }`. `text` is the signature block itself (the `<div id="Signature">` element, not the whole body), ready to append to a reply. `inlinedImages` counts the logos embedded as data: URIs; any image too large (> 2 MB) or unfetchable keeps its raw `cid:` reference and is named in `note` — no placeholder is substituted, so the reference stays resolvable via get-mail-attachment. `--output-path` writes the HTML to a file.',
};

export { execute, meta, schema };
