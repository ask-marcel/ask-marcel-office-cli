import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { INCLUDE_HIDDEN_FOLDERS_OPTION } from './include-hidden-folders.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ includeHiddenFolders: z.enum(['true', 'false']).optional() }).strict();
// Plain (non-OData) query param, so it cannot ride appendOData's `$`-prefixed
// builder and has to be emitted by the path itself.
const { execute, schema } = buildListCommand((p) => (p.includeHiddenFolders === 'true' ? '/me/mailFolders?includeHiddenFolders=true' : '/me/mailFolders'), baseSchema);

const meta: CommandMeta = {
  summary: 'List the top-level mail folders in the signed-in user’s Outlook mailbox (Inbox, Sent Items, etc.).',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/mailFolders',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders',
  options: [INCLUDE_HIDDEN_FOLDERS_OPTION, ...odataQueryOptions],
  example: 'ask-marcel-office list-mail-folders',
  responseShape: 'collection of Microsoft Graph `mailFolder` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
