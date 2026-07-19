import { z } from 'zod';
import { err, ok, type Result } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { boundaryMarkerRefusal, commentCarriesQuoteBoundary, insertCommentAboveQuote } from './draft-comment-splicer.ts';
import { slimDraftResult } from './draft-response.ts';
import { formatZodError } from './format-zod-error.ts';
import { parseRecipients } from './parse-recipients.ts';

const schema = z.object({
  forwardMessageId: z.string().min(1),
  toRecipients: z.string().min(1),
  ccRecipients: z.string().optional(),
  bodyContent: z.string().min(1),
  subject: z.string().optional(),
  bodyContentType: z.enum(['Text', 'HTML']).optional(),
});

const isUnsentDraft = (value: unknown): value is { id: string } =>
  typeof value === 'object' && value !== null && 'id' in value && typeof value.id === 'string' && 'isDraft' in value && value.isDraft === true;

const draftBodySchema = z.object({ body: z.object({ contentType: z.string(), content: z.string() }) });

type DraftBody = z.infer<typeof draftBodySchema>['body'];

// Mirrors create-reply-draft's reader: the body normally rides back on the
// create response, so the common case costs no extra call. Duplicated rather
// than shared while there are only two of them (Rule of Three).
const readDraftBody = async (graph: GraphClient, created: unknown, draftId: string): Promise<Result<DraftBody, GraphError>> => {
  const inline = draftBodySchema.safeParse(created);
  if (inline.success) return ok(inline.data.body);
  const fetched = await graph.get(`/me/messages/${draftId}?$select=body`);
  if (!fetched.ok) return fetched;
  const parsed = draftBodySchema.safeParse(fetched.value);
  if (!parsed.success) {
    return err({
      type: 'api_error',
      status: 500,
      code: 'draft_body_unreadable',
      message: `Draft ${draftId} was created but its body could not be read back, so the comment was never written into it. The draft exists - review it in Outlook Drafts, or set its body with \`update-mail-draft --message-id ${draftId} --body-content ...\`.`,
    });
  }
  return ok(parsed.data.body);
};

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success)
    return err({
      type: 'validation_error',
      message: formatZodError(parsed.error),
    });
  const { forwardMessageId, toRecipients, ccRecipients, bodyContent, subject, bodyContentType } = parsed.data;

  const asHtml = bodyContentType === 'HTML';

  // A comment carrying a boundary marker of its own would be kept verbatim by the
  // splice, and the NEXT `update-mail-draft --comment` edit would cut the draft AT
  // the pasted marker, dropping the forwarded original below it. Refuse before
  // creating anything, so there is no orphan draft to clean up.
  if (asHtml && commentCarriesQuoteBoundary(bodyContent)) {
    return err({ type: 'validation_error', message: boundaryMarkerRefusal('--body-content') });
  }

  // Graph's createForward mints the draft in one call. On the Text path the
  // comment travels via `comment`, never a body PATCH: PATCHing `body` with
  // freshly built HTML replaces the whole draft body and drops the entire
  // forwarded original (it did - fixed 2026-07-13 after a live smoke). The HTML
  // path posts an EMPTY comment (Graph HTML-escapes what `comment` carries) and
  // splices the markup in below, keeping the quote inside the body it patches.
  const created = await graph.post(`/me/messages/${forwardMessageId}/createForward`, {
    comment: asHtml ? '' : bodyContent,
    toRecipients: parseRecipients(toRecipients),
  });
  if (!created.ok) return created;

  // Defense in depth: createForward is documented to return an UNSENT draft.
  // If Graph ever hands back anything else, stop before writing into it.
  if (!isUnsentDraft(created.value)) {
    return err({
      type: 'api_error',
      status: 500,
      code: 'not_an_unsent_draft',
      message: `createForward did not return an unsent draft for message ${forwardMessageId} - refusing to patch. Inspect the message id and retry.`,
    });
  }

  const draftId = created.value.id;

  // Text path, unchanged: optional cc / subject override PATCHes ONLY those -
  // never `body`, which carries the comment + quoted original. Nothing to set ->
  // the draft is already complete, return it as-is.
  if (!asHtml) {
    const patch: Record<string, unknown> = {};
    if (ccRecipients) patch.ccRecipients = parseRecipients(ccRecipients);
    if (subject) patch.subject = subject;
    if (Object.keys(patch).length === 0) return slimDraftResult(created);
    return slimDraftResult(await graph.patch(`/me/messages/${draftId}`, patch));
  }

  const draftBody = await readDraftBody(graph, created.value, draftId);
  if (!draftBody.ok) return draftBody;

  // A text-bodied original cannot take HTML: splicing markup in would show it as
  // literal characters, and converting the body to HTML would rewrite the quote.
  if (draftBody.value.contentType.toLowerCase() !== 'html') {
    return err({
      type: 'validation_error',
      message: `The forwarded draft's body is ${draftBody.value.contentType}, not HTML, so HTML cannot be placed above the quoted original without rewriting it. Draft ${draftId} was already created - set its text with \`update-mail-draft --message-id ${draftId} --comment "..."\`, or delete it in Outlook Drafts and retry without --body-content-type HTML.`,
    });
  }

  // The quote rides along inside `content`, so this PATCH is the sanctioned kind:
  // never PATCH body WITHOUT the quote in it. Cc and subject merge into the same
  // call rather than costing a second round trip.
  const spliced = insertCommentAboveQuote(draftBody.value.content, bodyContent);
  const patch: Record<string, unknown> = { body: { contentType: 'HTML', content: spliced.html } };
  if (ccRecipients) patch.ccRecipients = parseRecipients(ccRecipients);
  if (subject) patch.subject = subject;
  return slimDraftResult(await graph.patch(`/me/messages/${draftId}`, patch));
};

const meta: CommandMeta = {
  summary:
    'Create an UNSENT forward draft of an existing message. POST /me/messages/{id}/createForward mints the draft (FW: subject, quoted original) with your comment placed above the quote and the recipients set, in one call. Redirects a thread to the right owner without leaving the CLI. The draft is saved in Drafts and can be reviewed, edited, and sent from any Outlook client; the CLI still cannot send.',
  category: 'mail',
  graphMethod: 'POST',
  graphPathTemplate: '/me/messages/{forward-message-id}/createForward (+ optional body-free PATCH for cc / subject)',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-createforward',
  options: [
    {
      name: 'forward-message-id',
      key: 'forwardMessageId',
      required: true,
      aliases: [{ name: 'id', key: 'id' }],
      description: 'The message being forwarded. Source from list-mail-folder-messages or search-mail-messages. Accepts `--id` as an alias.',
      argumentHint: { kind: 'idOrName' },
    },
    {
      name: 'to-recipients',
      key: 'toRecipients',
      required: true,
      description:
        'Comma-separated list of recipient email addresses to forward to (e.g. "alice@example.com,bob@example.com"). Required: a forward without a recipient is not actionable.',
    },
    {
      name: 'cc-recipients',
      key: 'ccRecipients',
      required: false,
      description: 'Comma-separated list of CC recipient email addresses.',
    },
    {
      name: 'body-content',
      key: 'bodyContent',
      required: true,
      description: 'The comment text, placed above the quoted forwarded message. Plain text by default; pass --body-content-type HTML to send it as markup.',
    },
    {
      name: 'subject',
      key: 'subject',
      required: false,
      description: 'Optional subject override. Omit to keep the inherited "FW: ..." subject.',
    },
    {
      name: 'body-content-type',
      key: 'bodyContentType',
      required: false,
      description:
        'Format of --body-content: Text (default) or HTML. Text is handed to Graph as the forward comment, which HTML-escapes it, so markup shows as literal characters. HTML instead creates the draft with an empty comment and splices your markup in above the forwarded original, leaving it and its styles byte-identical. Rejected when your markup itself contains a quote boundary marker (a pasted reply chain), and when the original is a plain-text message.',
      argumentHint: { kind: 'magicValue', values: ['Text', 'HTML'] },
    },
  ],
  example:
    'ask-marcel-office create-forward-draft --forward-message-id "AAMkAD..." --to-recipients "bob@example.com" --body-content "Bob owns this now, forwarding for your action."',
  bodyTemplate:
    "Text: POST { comment: '{body-content}', toRecipients: '{to-recipients}' } then optional PATCH { ccRecipients?: '{cc-recipients}', subject?: '{subject}' }. HTML ({body-content-type}): POST { comment: '', toRecipients: '{to-recipients}' } then ONE PATCH { body: { contentType: 'HTML', content: <'{body-content}' spliced above the quote> }, ccRecipients?: '{cc-recipients}', subject?: '{subject}' }",
  mutates: true,
  scopesRequired: ['Mail.ReadWrite'],
  responseShape:
    "A confirmation of the write, NOT the whole message: `{ id, subject, toRecipients, ccRecipients, bccRecipients, importance, bodyPreview, isDraft, webLink, conversationId }` (only the fields Graph returned; `{ ok: true }` when Graph answers 204). The `body` is deliberately omitted — you just wrote it, and echoing a long thread's quoted history back cost ~174 KB of context per call. Read the full body with `get-mail-message --id <the returned id>` when you actually need it; `bodyPreview` is Graph's ~255-char summary, enough to confirm WHICH draft answered. The `id` is the draft — refine it with `update-mail-draft`, or open Outlook Drafts to review and send.",
};

export { execute, meta, schema };
