import { z } from 'zod';
import { err, ok, type Result } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { boundaryMarkerRefusal, commentCarriesQuoteBoundary, insertCommentAboveQuote } from './draft-comment-splicer.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({
  replyToMessageId: z.string().min(1),
  bodyContent: z.string().min(1),
  subject: z.string().optional(),
  replyAll: z.enum(['true', 'false']).optional(),
  bodyContentType: z.enum(['Text', 'HTML']).optional(),
});

const isUnsentDraft = (value: unknown): value is { id: string } =>
  typeof value === 'object' && value !== null && 'id' in value && typeof value.id === 'string' && 'isDraft' in value && value.isDraft === true;

const draftBodySchema = z.object({ body: z.object({ contentType: z.string(), content: z.string() }) });

type DraftBody = z.infer<typeof draftBodySchema>['body'];

// Graph returns the freshly minted draft's body on the create response, so the
// common case costs no extra call. It is not contractually guaranteed, so fall
// back to reading it back rather than guessing at the scaffolding.
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
      message: `Draft ${draftId} was created but its body could not be read back, so the reply text was never written into it. The draft exists - review it in Outlook Drafts, or set its body with \`update-mail-draft --message-id ${draftId} --body-content ...\`.`,
    });
  }
  return ok(parsed.data.body);
};

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { replyToMessageId, bodyContent, subject, replyAll, bodyContentType } = parsed.data;

  // Reply-all stays the default: dropping recipients is a deliberate act, so it
  // takes an explicit `false`. Anything else (absent, or `true`) replies to all.
  const action = replyAll === 'false' ? 'createReply' : 'createReplyAll';
  const asHtml = bodyContentType === 'HTML';

  // A reply that carries a boundary marker of its own would be kept verbatim by
  // the splice, and the NEXT `update-mail-draft --comment` edit would cut the
  // draft AT the pasted marker, silently dropping everything below it. Refuse
  // before creating anything, so there is no orphan draft to clean up.
  if (asHtml && commentCarriesQuoteBoundary(bodyContent)) {
    return err({ type: 'validation_error', message: boundaryMarkerRefusal('--body-content') });
  }

  // Graph mints the threaded draft (inherited recipients, RE: subject, quoted
  // history) in one call. On the Text path the reply travels via `comment`,
  // never a body PATCH: PATCHing `body` with freshly built HTML replaces the
  // whole draft body and drops the quoted thread (fixed 2026-07-13). The HTML
  // path posts an EMPTY comment (Graph HTML-escapes what `comment` carries) and
  // splices the markup in below, keeping the quote inside the body it patches.
  const created = await graph.post(`/me/messages/${replyToMessageId}/${action}`, { comment: asHtml ? '' : bodyContent });
  if (!created.ok) return created;

  // Defense in depth: both reply actions are documented to return an UNSENT
  // draft. If Graph ever hands back anything else, stop before writing into it.
  if (!isUnsentDraft(created.value)) {
    return err({
      type: 'api_error',
      status: 500,
      code: 'not_an_unsent_draft',
      message: `${action} did not return an unsent draft for message ${replyToMessageId} - refusing to patch. Inspect the message id and retry.`,
    });
  }

  const draftId = created.value.id;

  // Text path, unchanged: optional subject override PATCHes ONLY the subject -
  // never `body`, which carries the reply text + quoted history. No override ->
  // return the draft as Graph made it.
  if (!asHtml) {
    const patch: Record<string, unknown> = {};
    if (subject) patch.subject = subject;
    if (Object.keys(patch).length === 0) return created;
    return graph.patch(`/me/messages/${draftId}`, patch);
  }

  const draftBody = await readDraftBody(graph, created.value, draftId);
  if (!draftBody.ok) return draftBody;

  // A text-bodied thread cannot take HTML: splicing markup in would show it as
  // literal characters, and converting the body to HTML would rewrite the quote.
  if (draftBody.value.contentType.toLowerCase() !== 'html') {
    return err({
      type: 'validation_error',
      message: `The thread's draft body is ${draftBody.value.contentType}, not HTML, so HTML cannot be placed above its quoted history without rewriting the quote. Draft ${draftId} was already created - reply to it with \`update-mail-draft --message-id ${draftId} --comment "..."\`, or delete it in Outlook Drafts and retry without --body-content-type HTML.`,
    });
  }

  // The quote rides along inside `content`, so this PATCH is the sanctioned kind:
  // never PATCH body WITHOUT the quote in it. Subject merges into the same call.
  const spliced = insertCommentAboveQuote(draftBody.value.content, bodyContent);
  const patch: Record<string, unknown> = { body: { contentType: 'HTML', content: spliced.html } };
  if (subject) patch.subject = subject;
  return graph.patch(`/me/messages/${draftId}`, patch);
};

const meta: CommandMeta = {
  summary:
    'Create an UNSENT reply draft threaded on an existing message. POST /me/messages/{id}/createReplyAll mints the draft (inherited recipients, RE: subject, quoted history) with your reply text placed above the quote, in one call. Reply-all by default - dropping recipients is a deliberate act, so pass --reply-all false to reply to the sender only, which switches the action to createReply. The draft is saved in Drafts and can be reviewed, edited, and sent from any Outlook client; the CLI still cannot send.',
  category: 'mail',
  graphMethod: 'POST',
  graphPathTemplate: '/me/messages/{reply-to-message-id}/createReplyAll, or /createReply when {reply-all} is false (+ optional body-free PATCH for subject)',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-createreplyall',
  options: [
    {
      name: 'reply-to-message-id',
      key: 'replyToMessageId',
      required: true,
      aliases: [{ name: 'id', key: 'id' }],
      description: 'The message being replied to. Source from list-mail-folder-messages or search-mail-messages. Accepts `--id` as an alias.',
      argumentHint: { kind: 'idOrName' },
    },
    {
      name: 'body-content',
      key: 'bodyContent',
      required: true,
      description: 'The reply text, placed above the quoted history. Plain text by default; pass --body-content-type HTML to send it as markup.',
    },
    {
      name: 'body-content-type',
      key: 'bodyContentType',
      required: false,
      description:
        'Format of --body-content: Text (default) or HTML. Text is handed to Graph as the reply comment, which HTML-escapes it, so markup shows as literal characters. HTML instead creates the draft with an empty comment and splices your markup in above the quote, leaving the quoted thread and its styles byte-identical. Rejected when your markup itself contains a quote boundary marker (a pasted reply chain), and when the thread is a plain-text one.',
      argumentHint: { kind: 'magicValue', values: ['Text', 'HTML'] },
    },
    {
      name: 'subject',
      key: 'subject',
      required: false,
      description: 'Optional subject override. Omit to keep the inherited "RE: ..." subject.',
    },
    {
      name: 'reply-all',
      key: 'replyAll',
      required: false,
      description:
        'Who the draft replies to. Defaults to true (everyone on the thread: sender + To + Cc, via createReplyAll). Pass `false` to reply to the sender only, via createReply. Only an explicit `false` narrows the recipients - anything else keeps reply-all.',
      argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
    },
  ],
  example: 'ask-marcel-office create-reply-draft --reply-to-message-id "AAMkAD..." --body-content "Confirmed for Contoso, aligned with the group choice."',
  bodyTemplate:
    "Text: POST { comment: '{body-content}' } then optional PATCH { subject?: '{subject}' }. HTML ({body-content-type}): POST { comment: '' } then ONE PATCH { body: { contentType: 'HTML', content: <'{body-content}' spliced above the quote> }, subject?: '{subject}' }",
  mutates: true,
  scopesRequired: ['Mail.ReadWrite'],
  responseShape:
    'The updated draft message object (or `{ ok: true }` when Graph answers 204): `{ id, subject, body, toRecipients, ccRecipients, isDraft: true, … }`. The `id` is the draft - update further with update-mail-draft, or open Outlook Drafts to review and send.',
};

export { execute, meta, schema };
