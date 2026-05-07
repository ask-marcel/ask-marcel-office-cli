import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/mailFolders', baseSchema);

const meta: CommandMeta = {
  summary: 'List the top-level mail folders in the signed-in user’s Outlook mailbox (Inbox, Sent Items, etc.).',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/mailFolders',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-mail-folders',
  responseShape: 'collection of Microsoft Graph `mailFolder` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
