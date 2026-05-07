import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/drive/root/delta()', baseSchema);

const meta: CommandMeta = {
  summary:
    "Track incremental changes (added / modified / deleted items) anywhere under the signed-in user's OneDrive root. The first call returns a snapshot plus `@odata.deltaLink`; subsequent calls with that link return only what has changed since. Cross-folder companion to `get-drive-delta` (which scopes to one specific folder).",
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drive/root/delta()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-delta',
  options: [...odataQueryOptions],
  example: 'ask-marcel get-drive-root-delta',
  responseShape: 'collection of Microsoft Graph `driveItem` resources plus `@odata.deltaLink` / `@odata.nextLink`',
  pagination: true,
};

export { execute, meta, schema };
