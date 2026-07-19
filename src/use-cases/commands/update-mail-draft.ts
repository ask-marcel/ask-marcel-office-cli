import { z } from 'zod';
import { err, ok, type Result } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { boundaryMarkerRefusal, commentCarriesQuoteBoundary, escapeTextAsHtml, replaceCommentAboveQuote, replacePlainTextCommentAboveQuote } from './draft-comment-splicer.ts';
import { slimDraftResult } from './draft-response.ts';
import { formatZodError } from './format-zod-error.ts';
import { parseRecipients } from './parse-recipients.ts';

const schema = z.object({
  messageId: z.string().min(1),
  subject: z.string().optional(),
  bodyContent: z.string().optional(),
  comment: z.string().min(1).optional(),
  bodyContentType: z.enum(['Text', 'HTML']).optional(),
  toRecipients: z.string().optional(),
  ccRecipients: z.string().optional(),
  bccRecipients: z.string().optional(),
  importance: z.enum(['Low', 'Normal', 'High']).optional(),
});

const draftSchema = z.object({ isDraft: z.boolean(), body: z.object({ contentType: z.string(), content: z.string() }) });

type DraftBody = z.infer<typeof draftSchema>['body'];

const noQuoteRefusal = (messageId: string): GraphError => ({
  type: 'validation_error',
  message: `Draft ${messageId} has no quoted reply history to preserve, so there is nothing for --comment to sit above. Use --body-content to replace the whole body instead.`,
});

// Rewrites the reply text ABOVE the quote, keeping the quote. The draft's own
// contentType is passed through verbatim: the caller asked to change the text,
// not the format. A body with no boundary is refused rather than overwritten -
// replacing an unbounded body is exactly the 2026-07-13 regression.
const reviseBodyAboveQuote = (draft: DraftBody, messageId: string, comment: string, asHtml: boolean): Result<DraftBody, GraphError> => {
  if (draft.contentType.toLowerCase() !== 'html') {
    if (asHtml) {
      return err({
        type: 'validation_error',
        message: `Draft ${messageId} has a ${draft.contentType} body, so HTML would show as literal characters. Drop --body-content-type HTML to revise it as plain text, or use --body-content to replace the whole body with HTML.`,
      });
    }
    const revised = replacePlainTextCommentAboveQuote(draft.content, comment);
    if (!revised.boundaryFound) return err(noQuoteRefusal(messageId));
    return ok({ contentType: draft.contentType, content: revised.text });
  }
  const revised = replaceCommentAboveQuote(draft.content, asHtml ? comment : escapeTextAsHtml(comment));
  if (!revised.boundaryFound) return err(noQuoteRefusal(messageId));
  return ok({ contentType: draft.contentType, content: revised.html });
};

const readDraft = async (graph: GraphClient, messageId: string): Promise<Result<z.infer<typeof draftSchema>, GraphError>> => {
  const fetched = await graph.get(`/me/messages/${messageId}?$select=body,isDraft`);
  if (!fetched.ok) return fetched;
  const parsed = draftSchema.safeParse(fetched.value);
  if (!parsed.success) {
    return err({
      type: 'validation_error',
      message: `Message ${messageId} did not come back with a readable body, so its reply text cannot be revised. Use get-mail-message to inspect it.`,
    });
  }
  // A sent message has no reply-above-the-quote to revise, and rewriting one
  // would be editing history rather than a draft.
  if (!parsed.data.isDraft) {
    return err({ type: 'validation_error', message: `Message ${messageId} is not a draft, so its body cannot be revised. Only unsent drafts can be updated.` });
  }
  return ok(parsed.data);
};

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { messageId, subject, bodyContent, comment, bodyContentType, toRecipients, ccRecipients, bccRecipients, importance } = parsed.data;

  // At least one field must be provided for the update. Keyed on presence, not
  // truthiness: an empty string is a real instruction (clear this list), and a
  // truthiness gate would silently read it as "flag absent" and refuse.
  if ([subject, bodyContent, comment, toRecipients, ccRecipients, bccRecipients, importance].every((v) => v === undefined)) {
    return err({
      type: 'validation_error',
      message:
        'At least one field must be provided to update (--subject, --body-content, --comment, --to-recipients, --cc-recipients, --bcc-recipients, or --importance). Pass an empty string to a recipient flag to clear that list.',
    });
  }

  // They mean opposite things - revise the text above the quote, vs replace the
  // whole body - so the CLI must not guess which was meant.
  if (comment !== undefined && bodyContent !== undefined) {
    return err({
      type: 'validation_error',
      message:
        '--comment and --body-content are mutually exclusive. --comment rewrites only the reply text ABOVE the quoted history and keeps the quote; --body-content replaces the entire body, quote included. Pass whichever one you meant.',
    });
  }

  if (comment !== undefined && bodyContentType === 'HTML' && commentCarriesQuoteBoundary(comment)) {
    return err({ type: 'validation_error', message: boundaryMarkerRefusal('--comment') });
  }

  const body: Record<string, unknown> = {};
  if (subject !== undefined) body.subject = subject;
  if (bodyContent !== undefined) body.body = { contentType: bodyContentType ?? 'Text', content: bodyContent };
  // `parseRecipients('')` yields `[]`, which is exactly Graph's clear payload, so
  // an inherited Cc list can finally be dropped without leaving the CLI.
  if (toRecipients !== undefined) body.toRecipients = parseRecipients(toRecipients);
  if (ccRecipients !== undefined) body.ccRecipients = parseRecipients(ccRecipients);
  if (bccRecipients !== undefined) body.bccRecipients = parseRecipients(bccRecipients);
  if (importance !== undefined) body.importance = importance;

  if (comment === undefined) return slimDraftResult(await graph.patch(`/me/messages/${messageId}`, body));

  // The comment path costs one read: the reply text can only be replaced in
  // place if we know what the quote below it looks like.
  const draft = await readDraft(graph, messageId);
  if (!draft.ok) return draft;
  const revised = reviseBodyAboveQuote(draft.value.body, messageId, comment, bodyContentType === 'HTML');
  if (!revised.ok) return revised;
  // The quote rides along inside `content`, so this PATCH is the sanctioned kind:
  // never PATCH body WITHOUT the quote in it.
  body.body = revised.value;
  return slimDraftResult(await graph.patch(`/me/messages/${messageId}`, body));
};

const meta: CommandMeta = {
  summary:
    'Update an existing mail draft. PATCH /me/messages/{id} — modifies a draft created by create-mail-draft (or any existing draft in the Drafts folder). Only the fields you pass are updated; omitted fields are left unchanged. At least one field must be provided. On a THREADED draft (one made by create-reply-draft / create-forward-draft), revise your text with --comment, which rewrites only what sits above the quoted history and leaves the quote byte-identical; --body-content would replace the whole body and drop the thread. Passing an EMPTY string to a recipient flag clears that list, which is how you drop recipients a reply-all or forward inherited; omitting the flag leaves the list alone. Returns a slim confirmation (id, subject, recipients, importance, bodyPreview, …) - NOT the full body, which you just wrote; read it back with get-mail-message if you need the whole draft before sending.',
  category: 'mail',
  graphMethod: 'PATCH',
  graphPathTemplate: '/me/messages/{message-id} (+ a GET of body,isDraft first when {comment} is used)',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-update',
  options: [
    {
      name: 'message-id',
      key: 'messageId',
      required: true,
      aliases: [{ name: 'id', key: 'id' }],
      description: 'Draft message ID to update. Source from create-mail-draft response or list-mail-folder-messages --mail-folder-id drafts. Accepts `--id` as an alias.',
      argumentHint: { kind: 'idOrName' },
    },
    {
      name: 'subject',
      key: 'subject',
      required: false,
      description: 'New email subject line. Omit to keep the current subject.',
    },
    {
      name: 'body-content',
      key: 'bodyContent',
      required: false,
      description:
        'New email body content. Replaces the ENTIRE body, quoted history included. On a threaded reply or forward draft this is almost never what you want - use --comment to revise only your own text and keep the quote. Pass --body-content-type HTML for rich text. Mutually exclusive with --comment.',
    },
    {
      name: 'comment',
      key: 'comment',
      required: false,
      description:
        'Rewrite ONLY the reply text above the quoted history on a threaded draft, keeping the quote and its styles byte-identical. This is the flag for revising a draft made by create-reply-draft or create-forward-draft; repeated edits replace your text rather than stacking. Refused when the draft has no quoted history (use --body-content), when it is not a draft, and when HTML markup you pass carries a quote boundary marker of its own. Mutually exclusive with --body-content.',
    },
    {
      name: 'body-content-type',
      key: 'bodyContentType',
      required: false,
      description:
        'Format of --body-content, or of --comment in comment mode: Text (default) or HTML. With --comment on an HTML draft, Text is escaped into the draft (markup shows as characters) and HTML is spliced in as markup; HTML is refused on a plain-text draft. The draft keeps its own body format either way - this flag describes what you are passing, not what the draft becomes.',
      argumentHint: { kind: 'magicValue', values: ['Text', 'HTML'] },
    },
    {
      name: 'to-recipients',
      key: 'toRecipients',
      required: false,
      description:
        'Comma-separated list of recipient email addresses. Replaces the entire toRecipients list. Pass an empty string to CLEAR the list (the only way to drop recipients a reply or forward inherited).',
    },
    {
      name: 'cc-recipients',
      key: 'ccRecipients',
      required: false,
      description:
        'Comma-separated list of CC recipient email addresses. Replaces the entire ccRecipients list. Pass an empty string to CLEAR the list (the only way to drop recipients a reply or forward inherited).',
    },
    {
      name: 'bcc-recipients',
      key: 'bccRecipients',
      required: false,
      description:
        'Comma-separated list of BCC recipient email addresses. Replaces the entire bccRecipients list. Pass an empty string to CLEAR the list (the only way to drop recipients a reply or forward inherited).',
    },
    {
      name: 'importance',
      key: 'importance',
      required: false,
      description: 'Email importance: Low, Normal, or High.',
      argumentHint: { kind: 'magicValue', values: ['Low', 'Normal', 'High'] },
    },
  ],
  example: 'ask-marcel-office update-mail-draft --message-id "AAMkAD..." --subject "Updated: Q3 Report" --to-recipients "alice@example.com,charlie@example.com"',
  bodyTemplate:
    "{ subject?: '{subject}', body?: { contentType: '{body-content-type}', content: '{body-content}' }, toRecipients?: '{to-recipients}', ccRecipients?: '{cc-recipients}', bccRecipients?: '{bcc-recipients}', importance?: '{importance}' } — only provided fields are sent. With '{comment}': body.content is the draft's own body with the text above the quote replaced, and body.contentType is the draft's own, unchanged",
  mutates: true,
  scopesRequired: ['Mail.ReadWrite'],
  responseShape:
    "A confirmation of the write, NOT the whole message: `{ id, subject, toRecipients, ccRecipients, bccRecipients, importance, bodyPreview, isDraft, webLink, conversationId }` (only the fields Graph returned; `{ ok: true }` when Graph answers 204). The `body` is deliberately omitted — you just wrote it, and echoing a long thread's quoted history back cost ~174 KB of context per call. Read the full body with `get-mail-message --id <the returned id>` when you actually need it; `bodyPreview` is Graph's ~255-char summary, enough to confirm WHICH draft answered. The `id` is the draft — refine it with `update-mail-draft`, or open Outlook Drafts to review and send. Dedup caveat: a `conversationId` `$filter` on the Drafts folder is not a reliable 'does a draft already exist on this thread' check. Reply and forward drafts do not always inherit the inbound message's `conversationId` (one thread can split across several), and Graph `$filter` on Drafts is not read-your-writes consistent, so a just-created draft can be missed. To find existing drafts, use the `find-mail-drafts` command, which scans recent drafts and matches client-side on subject and recipients; to revise a draft this session created, reuse the returned `id`.",
};

export { execute, meta, schema };
