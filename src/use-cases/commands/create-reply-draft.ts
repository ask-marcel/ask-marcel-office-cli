import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({
  replyToMessageId: z.string().min(1),
  bodyContent: z.string().min(1),
  subject: z.string().optional(),
});

const isUnsentDraft = (value: unknown): value is { id: string } =>
  typeof value === 'object' && value !== null && 'id' in value && typeof value.id === 'string' && 'isDraft' in value && value.isDraft === true;

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { replyToMessageId, bodyContent, subject } = parsed.data;

  // Graph mints the threaded reply-all draft (inherited recipients, RE: subject,
  // quoted history) with the reply text placed ABOVE the quote in one call. The
  // text MUST travel via `comment`, never a body PATCH: PATCHing `body` replaces
  // the whole draft body and drops the quoted thread (fixed 2026-07-13).
  const created = await graph.post(`/me/messages/${replyToMessageId}/createReplyAll`, { comment: bodyContent });
  if (!created.ok) return created;

  // Defense in depth: createReplyAll is documented to return an UNSENT draft.
  // If Graph ever hands back anything else, stop before writing into it.
  if (!isUnsentDraft(created.value)) {
    return err({
      type: 'api_error',
      status: 500,
      code: 'not_an_unsent_draft',
      message: `createReplyAll did not return an unsent draft for message ${replyToMessageId} - refusing to patch. Inspect the message id and retry.`,
    });
  }

  // Optional subject override: PATCH ONLY the subject - never `body`, which
  // carries the reply text + quoted history. No override -> return the draft.
  const patch: Record<string, unknown> = {};
  if (subject) patch.subject = subject;
  if (Object.keys(patch).length === 0) return created;
  return graph.patch(`/me/messages/${created.value.id}`, patch);
};

const meta: CommandMeta = {
  summary:
    'Create an UNSENT reply-all draft threaded on an existing message. POST /me/messages/{id}/createReplyAll mints the draft (inherited recipients, RE: subject, quoted history) with your reply text placed above the quote, in one call. Reply-all by design - dropping recipients is a deliberate act for the human in Outlook, not a default. The draft is saved in Drafts and can be reviewed, edited, and sent from any Outlook client; the CLI still cannot send.',
  category: 'mail',
  graphMethod: 'POST',
  graphPathTemplate: '/me/messages/{reply-to-message-id}/createReplyAll (+ optional body-free PATCH for subject)',
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
      description: 'The reply text, placed above the quoted history by Graph.',
    },
    {
      name: 'subject',
      key: 'subject',
      required: false,
      description: 'Optional subject override. Omit to keep the inherited "RE: ..." subject.',
    },
  ],
  example: 'ask-marcel-office create-reply-draft --reply-to-message-id "AAMkAD..." --body-content "Confirmed for Concur, aligned with the group choice."',
  bodyTemplate: "POST { comment: '{body-content}' } then optional PATCH { subject?: '{subject}' }",
  mutates: true,
  scopesRequired: ['Mail.ReadWrite'],
  responseShape:
    'The updated draft message object (or `{ ok: true }` when Graph answers 204): `{ id, subject, body, toRecipients, ccRecipients, isDraft: true, … }`. The `id` is the draft - update further with update-mail-draft, or open Outlook Drafts to review and send.',
};

export { execute, meta, schema };
