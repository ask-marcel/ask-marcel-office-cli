import { z } from 'zod';
import { err } from '../../domain/result.ts';
import { buildPickODataListCommand } from './build-command.ts';
import type { Command, CommandMeta } from './command-types.ts';
import type { WithdrawnFlag } from './delta-withdrawn-flags.ts';
import { withdrawnFlagRefusal } from './delta-withdrawn-flags.ts';
import { pickODataOptions } from './odata-query.ts';

// Live probe 2026-07-23: `--filter name eq '<no match>'` and `--orderby name
// asc` each returned the SAME first item as an unfiltered call, so Graph drops
// both on driveItem delta. `$top` is honoured here as a genuine page size
// (unlike the mail delta), so it stays on the query string.
const WITHDRAWN_FLAGS: ReadonlyArray<WithdrawnFlag> = [
  { key: 'filter', reason: 'Graph ignores `$filter` on driveItem delta: a predicate matching nothing still returned the full first page. Filter client-side.' },
  { key: 'orderby', reason: 'Graph ignores `$orderby` on driveItem delta: the first item was identical with and without it. Sort client-side.' },
];

const baseSchema = z.object({});
const built = buildPickODataListCommand(() => '/me/drive/root/delta()', baseSchema, ['top', 'select', 'expand']);
const schema = built.schema;

const execute: Command['execute'] = async (graph, params) => {
  const withdrawn = withdrawnFlagRefusal('get-drive-root-delta', WITHDRAWN_FLAGS, params);
  if (withdrawn !== undefined) return err({ type: 'validation_error', message: withdrawn });
  return built.execute(graph, params);
};

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
