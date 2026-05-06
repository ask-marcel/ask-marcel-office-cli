import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ groupId: z.string().min(1) });
const { execute } = buildCommand((p) => `/groups/${p.groupId}/owners`, schema);

const meta: CommandMeta = {
  summary: 'List the owners of an Azure AD / Microsoft 365 group.',
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}/owners',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/group-list-owners',
  options: [
    {
      name: 'group-id',
      key: 'groupId',
      required: true,
      description: 'Azure AD group object ID. Use `list-groups` to find one.',
    },
  ],
  example: "ask-marcel list-group-owners --group-id 'a1b2c3d4-...'",
  responseShape: 'collection of Microsoft Graph `directoryObject` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
