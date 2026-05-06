import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ groupId: z.string().min(1) });
const { execute } = buildCommand((p) => `/groups/${p.groupId}`, schema);

const meta: CommandMeta = {
  summary: 'Return metadata for a single Azure AD / Microsoft 365 group.',
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/groups/{group-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/group-get',
  options: [
    {
      name: 'group-id',
      key: 'groupId',
      required: true,
      description: 'Azure AD group object ID. Use `list-groups` to find one.',
    },
  ],
  example: "ask-marcel get-group --group-id 'a1b2c3d4-...'",
  responseShape: 'single Microsoft Graph `group` resource',
};

export { execute, meta, schema };
