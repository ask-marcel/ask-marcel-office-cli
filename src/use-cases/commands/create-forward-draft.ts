import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { parseRecipients } from './parse-recipients.ts';

const schema = z.object({
  forwardMessageId: z.string().min(1),
  toRecipients: z.string().min(1),
  ccRecipients: z.string().optional(),
  bodyContent: z.string().min(1),
  subject: z.string().optional(),
});

const isUnsentDraft = (value: unknown): value is { id: string } =>
  typeof value === 'object' && value !== null && 'id' in value && typeof value.id === 'string' && 'isDraft' in value && value.isDraft === true;

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success)
    return err({
      type: 'validation_error',
      message: formatZodError(parsed.error),
    });
  const { forwardMessageId, toRecipients, ccRecipients, bodyContent, subject } = parsed.data;

  // Graph's createForward mints the draft with the comment placed ABOVE the
  // quoted forwarded message in one call. The comment MUST travel via `comment`,
  // never a body PATCH: PATCHing `body` replaces the whole draft body and drops
  // the entire forwarded original (it did - fixed 2026-07-13 after a live smoke).
  const created = await graph.post(`/me/messages/${forwardMessageId}/createForward`, {
    comment: bodyContent,
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

  // Optional cc recipients / subject override: PATCH ONLY these - never `body`,
  // which carries the comment + quoted original. Nothing to set -> the draft is
  // already complete, return it as-is.
  const patch: Record<string, unknown> = {};
  if (ccRecipients) patch.ccRecipients = parseRecipients(ccRecipients);
  if (subject) patch.subject = subject;
  if (Object.keys(patch).length === 0) return created;
  return graph.patch(`/me/messages/${created.value.id}`, patch);
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
      description: 'The comment text, placed above the quoted forwarded message by Graph.',
    },
    {
      name: 'subject',
      key: 'subject',
      required: false,
      description: 'Optional subject override. Omit to keep the inherited "FW: ..." subject.',
    },
  ],
  example:
    'ask-marcel-office create-forward-draft --forward-message-id "AAMkAD..." --to-recipients "bob@example.com" --body-content "Bob owns this now, forwarding for your action."',
  bodyTemplate: "POST { comment: '{body-content}', toRecipients: '{to-recipients}' } then optional PATCH { ccRecipients?: '{cc-recipients}', subject?: '{subject}' }",
  mutates: true,
  scopesRequired: ['Mail.ReadWrite'],
  responseShape:
    'The updated draft message object (or `{ ok: true }` when Graph answers 204): `{ id, subject, body, toRecipients, ccRecipients, isDraft: true, … }`. The `id` is the draft - update further with update-mail-draft, or open Outlook Drafts to review and send.',
};

export { execute, meta, schema };
