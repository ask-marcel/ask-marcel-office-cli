import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { pickODataOptions } from './odata-query.ts';

// Live probe 2026-07-23: `--filter name eq '<no match>'` and `--orderby name
// asc` each returned the SAME first item as an unfiltered call, so Graph drops
// both on driveItem delta; neither is declared. `$top` IS honoured here as a
// genuine page size (unlike the mail delta), so it stays on the query string.
// Passing an undeclared flag is refused by the registry-level guard.
const baseSchema = z.object({});
const { execute, schema } = buildPickODataListCommand(() => '/me/drive/root/delta()', baseSchema, ['top', 'select', 'expand']);

const meta: CommandMeta = {
  summary:
    "Track incremental changes (added / modified / deleted items) anywhere under the signed-in user's OneDrive root. **Takes zero required arguments** — acts implicitly on the signed-in user's primary OneDrive; use `get-drive-delta` to target a specific drive by ID. The first call returns a snapshot plus `@odata.deltaLink`; subsequent calls with that link return only what has changed since. Cross-folder companion to `get-drive-delta` (which scopes to one specific folder).",
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drive/root/delta()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-delta',
  options: [...pickODataOptions(['top', 'select', 'expand'])],
  example: 'ask-marcel-office get-drive-root-delta',
  responseShape:
    'collection of Microsoft Graph `driveItem` resources under `data.value[]`. Cursor tokens are hoisted to envelope level: top-level `nextLink` while paging, then top-level `deltaLink` on the final page.',
  pagination: true,
  paginationStrategy: 'deltaLink',
};

export { execute, meta, schema };
