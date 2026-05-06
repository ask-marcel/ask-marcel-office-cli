import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ groupId: z.string().min(1) });
const { execute } = buildCommand((p) => `/groups/${p.groupId}/conversations`, schema);

const meta: CommandMeta = {
  summary: 'List conversations in a unified (Microsoft 365) group inbox. Each conversation aggregates one or more threads.',
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
  ],
  example: "ask-marcel list-group-conversations --group-id 'a1b2c3d4-...'",
  responseShape: 'collection of Microsoft Graph `conversation` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
