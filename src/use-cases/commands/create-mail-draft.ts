import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({
  subject: z.string().min(1),
  bodyContent: z.string().min(1),
  bodyContentType: z.enum(['Text', 'HTML']).optional(),
  toRecipients: z.string().min(1),
  ccRecipients: z.string().optional(),
  bccRecipients: z.string().optional(),
  importance: z.enum(['Low', 'Normal', 'High']).optional(),
  mailFolderId: z.string().optional(),
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
  const { subject, bodyContent, bodyContentType, toRecipients, ccRecipients, bccRecipients, importance, mailFolderId } = parsed.data;

  const body: Record<string, unknown> = {
    subject,
    body: {
      contentType: bodyContentType ?? 'Text',
      content: bodyContent,
    },
    toRecipients: parseRecipients(toRecipients),
  };

  if (ccRecipients) body.ccRecipients = parseRecipients(ccRecipients);
  if (bccRecipients) body.bccRecipients = parseRecipients(bccRecipients);
  if (importance) body.importance = importance;

  const path = mailFolderId ? `/me/mailFolders/${mailFolderId}/messages` : '/me/messages';
  return graph.post(path, body);
};

const meta: CommandMeta = {
  summary:
    'Create a new mail draft. POST /me/messages (or /me/mailFolders/{id}/messages when --mail-folder-id is set). The draft is saved in the Drafts folder (or the specified folder) and can be sent later via the Outlook client or Graph sendMail. Recipients are comma-separated email addresses. Returns the created message object with its id — use this id with update-mail-draft to modify the draft before sending.',
  category: 'mail',
  graphMethod: 'POST',
  graphPathTemplate: '/me/messages (or /me/mailFolders/{mail-folder-id}/messages)',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-post-messages',
  options: [
    {
      name: 'subject',
      key: 'subject',
      required: true,
      description: 'Email subject line.',
    },
    {
      name: 'body-content',
      key: 'bodyContent',
      required: true,
      description: 'Email body content. Plain text by default; pass --body-content-type HTML for rich text.',
    },
    {
      name: 'body-content-type',
      key: 'bodyContentType',
      required: false,
      description: 'Body format: Text (default) or HTML.',
      argumentHint: { kind: 'magicValue', values: ['Text', 'HTML'] },
    },
    {
      name: 'to-recipients',
      key: 'toRecipients',
      required: true,
      description: 'Comma-separated list of recipient email addresses (e.g. "alice@example.com,bob@example.com").',
    },
    {
      name: 'cc-recipients',
      key: 'ccRecipients',
      required: false,
      description: 'Comma-separated list of CC recipient email addresses.',
    },
    {
      name: 'bcc-recipients',
      key: 'bccRecipients',
      required: false,
      description: 'Comma-separated list of BCC recipient email addresses.',
    },
    {
      name: 'importance',
      key: 'importance',
      required: false,
      description: 'Email importance: Low, Normal (default), or High.',
      argumentHint: { kind: 'magicValue', values: ['Low', 'Normal', 'High'] },
    },
    {
      name: 'mail-folder-id',
      key: 'mailFolderId',
      required: false,
      description: 'Target folder ID to create the draft in. Defaults to the Drafts folder. Source from list-mail-folders.',
      argumentHint: { kind: 'idOrName' },
    },
  ],
  example:
    'ask-marcel create-mail-draft --subject "Q3 Report" --body-content "Please review the attached report." --to-recipients "alice@example.com,bob@example.com" --importance High',
  bodyTemplate:
    "{ subject: '{subject}', body: { contentType: '{body-content-type}', content: '{body-content}' }, toRecipients: '{to-recipients}', ccRecipients: '{cc-recipients}', bccRecipients: '{bcc-recipients}', importance: '{importance}' }",
  scopesRequired: ['Mail.ReadWrite'],
  responseShape:
    'The created Microsoft Graph message object: `{ id, subject, body, from, toRecipients, ccRecipients, bccRecipients, receivedDateTime, isDraft, … }`. The `id` field is the draft message ID — use it with `update-mail-draft` to modify before sending.',
};

export { execute, meta, schema };
