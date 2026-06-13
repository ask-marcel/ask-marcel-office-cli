import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({
  messageId: z.string().min(1),
  subject: z.string().optional(),
  bodyContent: z.string().optional(),
  bodyContentType: z.enum(['Text', 'HTML']).optional(),
  toRecipients: z.string().optional(),
  ccRecipients: z.string().optional(),
  bccRecipients: z.string().optional(),
  importance: z.enum(['Low', 'Normal', 'High']).optional(),
});

const parseRecipients = (csv: string): Array<{ emailAddress: { address: string } }> =>
  csv
    .split(',')
    .map((s) => s.trim())
    .filter((s) => s.length > 0)
    .map((address) => ({ emailAddress: { address } }));

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { messageId, subject, bodyContent, bodyContentType, toRecipients, ccRecipients, bccRecipients, importance } = parsed.data;

  // At least one field must be provided for the update.
  if (!subject && !bodyContent && !toRecipients && !ccRecipients && !bccRecipients && !importance) {
    return err({ type: 'validation_error', message: 'At least one field must be provided to update (--subject, --body-content, --to-recipients, --cc-recipients, --bcc-recipients, or --importance)' });
  }

  const body: Record<string, unknown> = {};
  if (subject !== undefined) body.subject = subject;
  if (bodyContent !== undefined) body.body = { contentType: bodyContentType ?? 'Text', content: bodyContent };
  if (toRecipients) body.toRecipients = parseRecipients(toRecipients);
  if (ccRecipients) body.ccRecipients = parseRecipients(ccRecipients);
  if (bccRecipients) body.bccRecipients = parseRecipients(bccRecipients);
  if (importance) body.importance = importance;

  return graph.patch(`/me/messages/${messageId}`, body);
};

const meta: CommandMeta = {
  summary:
    'Update an existing mail draft. PATCH /me/messages/{id} — modifies a draft created by create-mail-draft (or any existing draft in the Drafts folder). Only the fields you pass are updated; omitted fields are left unchanged. At least one field must be provided. Returns the updated message object. Use get-mail-message to verify the final state before sending.',
  category: 'mail',
  graphMethod: 'PATCH',
  graphPathTemplate: '/me/messages/{message-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-update',
  options: [
    {
      name: 'message-id',
      key: 'messageId',
      required: true,
      description: 'Draft message ID to update. Source from create-mail-draft response or list-mail-folder-messages --mail-folder-id drafts.',
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
      description: 'New email body content. Replaces the entire body. Pass --body-content-type HTML for rich text.',
    },
    {
      name: 'body-content-type',
      key: 'bodyContentType',
      required: false,
      description: 'Body format for the new body: Text (default) or HTML. Only used when --body-content is provided.',
      argumentHint: { kind: 'magicValue', values: ['Text', 'HTML'] },
    },
    {
      name: 'to-recipients',
      key: 'toRecipients',
      required: false,
      description: 'Comma-separated list of recipient email addresses. Replaces the entire toRecipients list.',
    },
    {
      name: 'cc-recipients',
      key: 'ccRecipients',
      required: false,
      description: 'Comma-separated list of CC recipient email addresses. Replaces the entire ccRecipients list.',
    },
    {
      name: 'bcc-recipients',
      key: 'bccRecipients',
      required: false,
      description: 'Comma-separated list of BCC recipient email addresses. Replaces the entire bccRecipients list.',
    },
    {
      name: 'importance',
      key: 'importance',
      required: false,
      description: 'Email importance: Low, Normal, or High.',
      argumentHint: { kind: 'magicValue', values: ['Low', 'Normal', 'High'] },
    },
  ],
  example:
    'ask-marcel update-mail-draft --message-id "AAMkAD..." --subject "Updated: Q3 Report" --to-recipients "alice@example.com,charlie@example.com"',
  bodyTemplate:
    "{ subject?: '{subject}', body?: { contentType: '{body-content-type}', content: '{body-content}' }, toRecipients?: '{to-recipients}', ccRecipients?: '{cc-recipients}', bccRecipients?: '{bcc-recipients}', importance?: '{importance}' } — only provided fields are sent",
  scopesRequired: ['Mail.ReadWrite'],
  responseShape:
    'The updated Microsoft Graph message object: `{ id, subject, body, from, toRecipients, ccRecipients, bccRecipients, receivedDateTime, isDraft, … }`. Graph returns 204 No Content on success with no body — the CLI surfaces `{ ok: true }` in that case.',
};

export { execute, meta, schema };
