import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { ReadCommandMeta } from './command-types.ts';
import { pickODataOptions } from './odata-query.ts';

// Graph refuses `$filter` on this collection. Probed live 2026-09-06 on a group
// the signed-in user belongs to: a predicate over `topic` answers
// `ErrorUnsupportedPathForQuery`, and unlike the sibling `threads` collection
// there is not even an `IsLocked` form that works. `$top`, `$skip`, `$orderby`,
// `$select` and `$expand` are honoured and stay.
const HONOURED = ['top', 'skip', 'select', 'orderby', 'expand'] as const;

const baseSchema = z.object({ groupId: z.string().min(1) });
const { execute, schema } = buildPickODataListCommand((p) => `/groups/${p.groupId}/conversations`, baseSchema, HONOURED);

const meta: ReadCommandMeta = {
  summary:
    "List conversations in a unified (Microsoft 365) group inbox. Each conversation aggregates one or more threads. Only Microsoft 365 groups have a mailbox — security and distribution groups return `MailboxNotEnabledForRESTAPI`. Verify the group is unified before calling. Bodies live two levels down: `list-group-threads` then `list-group-thread-posts`, or `--expand 'threads($expand=posts)'` here to fetch conversations, threads and posts in one call.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/conversations',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/group-list-conversations',
  options: [
    {
      name: 'group-id',
      key: 'groupId',
      required: true,
      description: 'Azure AD group object ID for a unified (Microsoft 365) group.',
    },
    ...pickODataOptions(HONOURED),
  ],
  example: "ask-marcel-office list-group-conversations --group-id 'a1b2c3d4-...'",
  responseShape: 'collection of Microsoft Graph `conversation` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Group.Read.All'],
};

export { execute, meta, schema };
