import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ groupId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/groups/${p.groupId}/threads`, baseSchema);

const meta: CommandMeta = {
  summary: "List threads in a unified (Microsoft 365) group inbox. Threads are flatter than conversations — one per topic, useful when conversation-level grouping isn't needed.",
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
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-group-threads --group-id 'a1b2c3d4-...'",
  responseShape: 'collection of Microsoft Graph `conversationThread` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
