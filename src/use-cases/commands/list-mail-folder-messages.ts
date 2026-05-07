import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ mailFolderId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/me/mailFolders/${p.mailFolderId}/messages`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the messages inside a specific Outlook mail folder (Inbox, custom folder, etc.).',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/mailFolders/{mail-folder-id}/messages',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/mailfolder-list-messages',
  options: [
    {
      name: 'mail-folder-id',
      key: 'mailFolderId',
      required: true,
      description: 'mailFolder ID. Returned by `ask-marcel list-mail-folders`. Well-known names also work, e.g. `inbox`, `sentitems`, `drafts`.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-mail-folder-messages --mail-folder-id 'inbox'",
  responseShape: 'collection of Microsoft Graph `message` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
