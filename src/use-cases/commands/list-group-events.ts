import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ groupId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/groups/${p.groupId}/events`, baseSchema);

const meta: CommandMeta = {
  summary:
    "List events from a unified (Microsoft 365) group's calendar. Only Microsoft 365 groups have a calendar — security and distribution groups return an empty `value[]` or 404.",
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/events',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/group-list-events',
  options: [
    {
      name: 'group-id',
      key: 'groupId',
      required: true,
      description: 'Azure AD group object ID for a unified (Microsoft 365) group. Use `list-groups` to find one.',
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-group-events --group-id 'a1b2c3d4-...'",
  responseShape: 'collection of Microsoft Graph `event` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
