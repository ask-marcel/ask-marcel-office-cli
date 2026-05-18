import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

// Graph's `/me/mailFolders/inbox/messageRules` silently ignores every
// standard OData passthrough (verified live — `--top 1` against a
// 2-rule mailbox still returns both). Don't advertise flags Graph drops.
const schema = z.object({ mailFolderId: z.string().min(1).default('inbox') });
const { execute } = buildCommand((p) => `/me/mailFolders/${p.mailFolderId}/messageRules`, schema);

const meta: CommandMeta = {
  summary:
    'List the message rules on the Outlook Inbox. Microsoft Graph only supports message rules on the Inbox folder; passing any other folder ID (drafts, sentitems, archive, a custom folder) returns `MailFolderNotSupportedError` from Graph. `--mail-folder-id` defaults to `inbox` because that is the only value Graph accepts; the flag is kept (optional) for callers that want to pass a resolved Inbox ID explicitly. Note: Graph silently ignores every OData passthrough on this endpoint, so the CLI does NOT expose them — the full rule set is always returned.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/mailFolders/{mail-folder-id}/messageRules',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/mailfolder-list-messagerules',
  options: [
    {
      name: 'mail-folder-id',
      key: 'mailFolderId',
      required: false,
      description:
        'mailFolder ID. Optional; defaults to `inbox`. In practice only `inbox` (or its resolved ID) works — Graph rejects every other folder for messageRules with `MailFolderNotSupportedError`.',
    },
  ],
  example: 'ask-marcel list-mail-rules',
  responseShape: 'collection of Microsoft Graph `messageRule` resources under `value[]`',
};

export { execute, meta, schema };
