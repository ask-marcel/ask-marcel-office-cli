import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ mailFolderId: z.string().min(1) });
const { execute } = buildCommand((p) => `/me/mailFolders/${p.mailFolderId}/messages/delta()`, schema);

const meta: CommandMeta = {
  summary:
    'Track incremental changes (added / updated / deleted messages) within a single mail folder using Microsoft Graph delta tokens. The first call returns the current snapshot plus a `@odata.deltaLink`; subsequent calls with that link return only what has changed since.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/mailFolders/{mail-folder-id}/messages/delta()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-delta',
  options: [
    {
      name: 'mail-folder-id',
      key: 'mailFolderId',
      required: true,
      description: 'Mail folder ID or well-known name (`inbox`, `archive`, `sentitems`, `deleteditems`, `junkemail`, `drafts`). Returned by `list-mail-folders`.',
    },
  ],
  example: "ask-marcel list-mail-folder-messages-delta --mail-folder-id 'inbox'",
  responseShape: 'collection of Microsoft Graph `message` resources plus `@odata.deltaLink` / `@odata.nextLink`',
  pagination: true,
};

export { execute, meta, schema };
