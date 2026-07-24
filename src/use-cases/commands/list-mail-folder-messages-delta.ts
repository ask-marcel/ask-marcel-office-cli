import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, pickODataOptions, pickODataShape } from './odata-query.ts';

// What this endpoint actually honours, established by live probe 2026-07-23
// against a 67-message Inbox (NOT by reading the docs):
//   $select   honoured  (--select id,subject returned only those fields)
//   $filter   honoured  (a future-dated predicate returned 0 items, not 10)
//   $expand   honoured  (--expand attachments returned a 750 KB page)
//   $skip     IGNORED   (--skip 5 returned the same first message)
//   $orderby  REJECTED  (ErrorInvalidUrlQuery; only the default
//                        `receivedDateTime desc` is accepted, so the flag can
//                        only ever be a no-op or an error)
//   $top      POISONOUS (see the header translation below)
// Only the honoured three are declared; passing `--skip` or `--orderby` is
// refused by the registry-level unknown-parameter guard.
const schema = z.object({ mailFolderId: z.string().min(1) }).extend(pickODataShape(['top', 'select', 'filter', 'expand']));

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { mailFolderId, top, ...odata } = parsed.data;

  // `$top` as a query parameter does NOT bound a page here — Graph treats the
  // satisfied `$top` as "this sync is complete" and answers with an
  // `@odata.deltaLink` instead of a `nextLink`. On the probed mailbox,
  // `--top 2` returned 2 of 67 messages and a delta token certifying a sync
  // that never happened; following that cursor returned 0 and the other 65
  // were never delivered. The same bound sent as `Prefer: odata.maxpagesize`
  // pages normally, which is also how the calendar delta commands carry it.
  const headers: Record<string, string> = {};
  if (top !== undefined) headers['Prefer'] = `odata.maxpagesize=${top}`;
  return graph.get(appendOData(`/me/mailFolders/${mailFolderId}/messages/delta()`, odata), headers);
};

const meta: CommandMeta = {
  summary:
    'Track incremental changes (added / updated / deleted messages) within a single mail folder using Microsoft Graph delta tokens. The first call returns the current snapshot plus a `@odata.deltaLink`; subsequent calls with that link return only what has changed since. `--top` is translated into the `Prefer: odata.maxpagesize=N` header: as a `$top` query parameter Graph reads a satisfied count as "sync complete" and hands back a deltaLink after N items, silently abandoning the rest of the folder. `$skip` and `$orderby` are NOT exposed — Graph ignores the former on this endpoint and rejects the latter unless it merely restates the default `receivedDateTime desc`.',
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
    ...pickODataOptions(['top', 'select', 'filter', 'expand']),
  ],
  example: "ask-marcel-office list-mail-folder-messages-delta --mail-folder-id 'inbox'",
  responseShape:
    'collection of Microsoft Graph `message` resources under `data.value[]`. Cursor tokens are hoisted to envelope level: top-level `nextLink` while paging, then top-level `deltaLink` on the final page (CLI strips the original `@odata.*` keys from `data`). A `deltaLink` on the FIRST page means the folder is fully synced, not that it was truncated.',
  pagination: true,
  paginationStrategy: 'preferMaxPageSize',
};

export { execute, meta, schema };
