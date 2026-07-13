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
  bodyContentType: z.enum(['Text', 'HTML']).optional(),
  subject: z.string().optional(),
});

const isUnsentDraft = (value: unknown): value is { id: string } =>
  typeof value === 'object' && value !== null && 'id' in value && typeof value.id === 'string' && 'isDraft' in value && value.isDraft === true;

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { forwardMessageId, toRecipients, ccRecipients, bodyContent, bodyContentType, subject } = parsed.data;

  // Step 1: Graph mints the forward draft - FW: subject, quoted history. The
  // empty body is deliberate: comment-based createForward renders poorly, so
  // the comment lands via PATCH instead (same rationale as create-reply-draft).
  const created = await graph.post(`/me/messages/${forwardMessageId}/createForward`, {});
  if (!created.ok) return created;

  // Defense in depth: createForward is documented to return an UNSENT draft.
  // If Graph ever hands back anything else, stop before writing into it.
  if (!isUnsentDraft(created.value)) {
    return err({
      type: 'api_error',
      status: 500,
      message: `createForward did not return an unsent draft for message ${forwardMessageId} - refusing to patch. Inspect the message id and retry.`,
    });
  }

  // Step 2: place the comment above the quote and set the recipients (plus an
  // optional subject override) into the draft.
  const patch: Record<string, unknown> = {
    body: { contentType: bodyContentType ?? 'Text', content: bodyContent },
    toRecipients: parseRecipients(toRecipients),
  };
  if (ccRecipients) patch.ccRecipients = parseRecipients(ccRecipients);
  if (subject) patch.subject = subject;

  return graph.patch(`/me/messages/${created.value.id}`, patch);
};

const meta: CommandMeta = {
  summary:
    'Create an UNSENT forward draft of an existing message. POST /me/messages/{id}/createForward mints the draft (FW: subject, quoted history), then PATCH places the comment above the quote and sets the recipients. Redirects a thread to the right owner without leaving the CLI. The draft is saved in Drafts and can be reviewed, edited, and sent from any Outlook client; the CLI still cannot send.',
  category: 'mail',
  graphMethod: 'POST',
  graphPathTemplate: '/me/messages/{forward-message-id}/createForward (then PATCH the returned draft)',
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
      description: 'The comment text, placed above the quoted message being forwarded. Plain text by default; pass --body-content-type HTML for rich text.',
    },
    {
      name: 'body-content-type',
      key: 'bodyContentType',
      required: false,
      description: 'Comment body format: Text (default) or HTML.',
      argumentHint: { kind: 'magicValue', values: ['Text', 'HTML'] },
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
  bodyTemplate:
    "POST {} then PATCH { body: { contentType: '{body-content-type}', content: '{body-content}' }, toRecipients: '{to-recipients}', ccRecipients?: '{cc-recipients}', subject?: '{subject}' }",
  mutates: true,
  scopesRequired: ['Mail.ReadWrite'],
  responseShape:
    'The updated draft message object (or `{ ok: true }` when Graph answers 204): `{ id, subject, body, toRecipients, ccRecipients, isDraft: true, … }`. The `id` is the draft - update further with update-mail-draft, or open Outlook Drafts to review and send.',
};

export { execute, meta, schema };
