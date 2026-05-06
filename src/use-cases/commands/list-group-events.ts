import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ groupId: z.string().min(1) });
const { execute } = buildCommand((p) => `/groups/${p.groupId}/events`, schema);

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
  ],
  example: "ask-marcel list-group-events --group-id 'a1b2c3d4-...'",
  responseShape: 'collection of Microsoft Graph `event` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
