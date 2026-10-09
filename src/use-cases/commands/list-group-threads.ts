import { z } from 'zod';
import { buildPickODataListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { pickODataOptions } from './odata-query.ts';

// Graph refuses `$filter` on this collection. Probed live 2026-09-06 on a group
// the signed-in user belongs to: any predicate answers `ConversationFilterOther`,
// and even `isLocked eq false` answers `ConversationFilterIsLockedEqualsFalse`,
// since the only accepted forms are `IsLocked eq true` and `IsLocked ne false`.
// Filtering a group inbox to its locked threads is not a use this CLI has, so
// the flag is dropped rather than advertised with a caveat: the manifest
// promises only what the endpoint honours. `$top`, `$skip`, `$orderby`,
// `$select` and `$expand` are all honoured and stay.
const HONOURED = ['top', 'skip', 'select', 'orderby', 'expand'] as const;

const baseSchema = z.object({ groupId: z.string().min(1) });
const { execute, schema } = buildPickODataListCommand((p) => `/groups/${p.groupId}/threads`, baseSchema, HONOURED);

const meta: CommandMeta = {
  summary:
    "List threads in a unified (Microsoft 365) group inbox. Threads are flatter than conversations — one per topic, useful when conversation-level grouping isn't needed. Only Microsoft 365 groups have a mailbox — security and distribution groups return `MailboxNotEnabledForRESTAPI`. Each thread carries only a truncated `preview` of its latest post: read the full posts with `list-group-thread-posts`, or pass `--expand posts` here to inline them.",
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/threads',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/group-list-threads',
  options: [
    {
      name: 'group-id',
      key: 'groupId',
      required: true,
      description: 'Azure AD group object ID for a unified (Microsoft 365) group.',
    },
    ...pickODataOptions(HONOURED),
  ],
  example: "ask-marcel-office list-group-threads --group-id 'a1b2c3d4-...'",
  responseShape: 'collection of Microsoft Graph `conversationThread` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Group.Read.All'],
};

export { execute, meta, schema };
