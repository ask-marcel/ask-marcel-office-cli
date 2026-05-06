import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/mailFolders/delta()', schema);

const meta: CommandMeta = {
  summary:
    'Track incremental changes to the mail-folder tree itself (folders added / renamed / deleted). The first call returns the current snapshot plus a `@odata.deltaLink`; subsequent calls with that link return only what has changed. Companion to `list-mail-folder-messages-delta` which tracks message changes inside one folder.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/me/mailFolders/delta()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/mailfolder-delta',
  options: [],
  example: 'ask-marcel list-mail-folders-delta',
  responseShape: 'collection of Microsoft Graph `mailFolder` resources plus `@odata.deltaLink` / `@odata.nextLink`',
  pagination: true,
};

export { execute, meta, schema };
