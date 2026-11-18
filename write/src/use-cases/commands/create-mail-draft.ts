import { z } from 'zod';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { WriteCommand, WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { slimDraftResult } from './draft-response.ts';
import { BODY_FILE_FIELD, bodyFileOption, createBodyFileExecute, createBodyFilePlan } from './draft-body-file.ts';
import type { BodyFileSpec } from './draft-body-file.ts';
import { createDraftPlan, DRAFT_CREATE_UNDO } from './draft-runs.ts';
import type { DraftTarget } from './draft-runs.ts';
import { formatZodError } from '@ask-marcel/office-core/use-cases/commands/format-zod-error.ts';
import { parseRecipients } from './parse-recipients.ts';

// The params after --body-file is read: the body text is in bodyContent.
const textSchema = z.object({
  subject: z.string().min(1),
  bodyContent: z.string().min(1),
  bodyContentType: z.enum(['Text', 'HTML']).optional(),
  toRecipients: z.string().min(1),
  ccRecipients: z.string().optional(),
  bccRecipients: z.string().optional(),
  importance: z.enum(['Low', 'Normal', 'High']).optional(),
  mailFolderId: z.string().optional(),
});

const schema = textSchema.extend({ bodyContent: textSchema.shape.bodyContent.optional(), bodyFile: BODY_FILE_FIELD });

const BODY: BodyFileSpec = { key: 'bodyContent', required: true, exclusive: ['bodyContent'] };

const textExecute: WriteCommand['execute'] = async (graph, params) => {
  const parsed = textSchema.safeParse(params);
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

  return slimDraftResult(await graph.post(draftPath(mailFolderId), body));
};

const draftPath = (mailFolderId: string | undefined): string => (mailFolderId ? `/me/mailFolders/${mailFolderId}/messages` : '/me/messages');

// What the draft is made in: the dry-run shows it, and a run records it.
const target = (params: Record<string, string>): Result<DraftTarget, GraphError> => {
  const parsed = textSchema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  return ok({ target: draftPath(parsed.data.mailFolderId), before: {} });
};

const execute = createBodyFileExecute(textExecute, BODY);

const plan = createBodyFilePlan(createDraftPlan({ action: 'create-draft', target, execute: textExecute }), BODY);

const meta: WriteCommandMeta = {
  summary:
    'Create a new mail draft. POST /me/messages (or /me/mailFolders/{id}/messages when --mail-folder-id is set). The draft is saved in the Drafts folder (or the specified folder) and can be sent later via the Outlook client or Graph sendMail. Recipients are comma-separated email addresses. Returns a slim confirmation (id, subject, recipients, importance, bodyPreview, …) - NOT the body you just wrote; read the full draft back with get-mail-message if you need it. Use the returned id with update-mail-draft to modify before sending.',
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
      required: false,
      description: 'Email body content. Plain text by default; pass --body-content-type HTML for rich text. Give it or --body-file.',
    },
    bodyFileOption('--body-content'),
    {
      name: 'body-content-type',
      key: 'bodyContentType',
      required: false,
      description: 'Format of the body text (--body-content or --body-file): Text (default) or HTML.',
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
    '{write-bin} create-mail-draft --subject "Q3 Report" --body-content "Please review the attached report." --to-recipients "alice@example.com,bob@example.com" --importance High',
  bodyTemplate:
    "{ subject: '{subject}', body: { contentType: '{body-content-type}', content: '{body-content}' (or the text of {body-file}) }, toRecipients: '{to-recipients}', ccRecipients: '{cc-recipients}', bccRecipients: '{bcc-recipients}', importance: '{importance}' }",
  effect: 'draft',
  undo: DRAFT_CREATE_UNDO,
  scopesRequired: ['Mail.ReadWrite'],
  responseShape:
    'A confirmation of the write, NOT the whole message: `{ id, subject, toRecipients, ccRecipients, bccRecipients, importance, bodyPreview, isDraft, webLink, conversationId }` (only the fields Graph returned; `{ ok: true }` when Graph answers 204). The `body` is deliberately omitted — you just supplied it; read it back with `get-mail-message --id <the returned id>` if you need it. The `id` is the draft — modify it with `update-mail-draft` before sending.',
};

export { execute, meta, plan, schema };
