import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();

// `/me/people` only emits a self-advancing `@odata.nextLink` when `$top` is on
// the request. Called bare, Graph returns 10 people AND a `?$skip=0` cursor
// that echoes the same `$skip` on every page — following it re-fetches page 1
// forever. Forcing a default `$top` (matching Graph's implicit page size of 10,
// so the default response is unchanged) makes Graph increment `$skip` by the
// page size, so `next-page` walks correctly. User `--top` overrides. See
// `build-command.ts` `withDefaultTop`.
const { execute, schema } = buildListCommand(() => '/me/people', baseSchema, { defaultTop: '10' });

const meta: CommandMeta = {
  summary:
    "List people relevant to the signed-in user — colleagues they email and meet with most. Microsoft's relevance ranking, not the full directory. Returns `displayName`, `emailAddresses`, `jobTitle`, `companyName`, etc.",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me/people',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-people',
  options: [...odataQueryOptions],
  example: 'ask-marcel-office list-relevant-people',
  responseShape: 'collection of Microsoft Graph `person` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
