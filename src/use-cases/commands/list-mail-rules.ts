import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ mailFolderId: z.string().min(1).default('inbox') });
const { execute, schema } = buildListCommand((p) => `/me/mailFolders/${p.mailFolderId}/messageRules`, baseSchema);

const meta: CommandMeta = {
  summary:
    'List the message rules on the Outlook Inbox. Microsoft Graph only supports message rules on the Inbox folder; passing any other folder ID (drafts, sentitems, archive, a custom folder) returns an `ErrorInvalidParameter` from Graph. `--mail-folder-id` defaults to `inbox` because that is the only value Graph accepts; the flag is kept (optional) for callers that want to pass a resolved Inbox ID explicitly.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/mailFolders/{mail-folder-id}/messageRules',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/mailfolder-list-messagerules',
  options: [
    {
      name: 'mail-folder-id',
      key: 'mailFolderId',
      required: false,
      description: 'mailFolder ID. Optional; defaults to `inbox`. In practice only `inbox` (or its resolved ID) works — Graph rejects every other folder for messageRules.',
    },
    ...odataQueryOptions,
  ],
  example: 'ask-marcel list-mail-rules',
  responseShape: 'collection of Microsoft Graph `messageRule` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
