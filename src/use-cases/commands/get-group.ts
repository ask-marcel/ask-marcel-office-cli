import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

const baseSchema = z.object({ groupId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/groups/${p.groupId}`, baseSchema);

const meta: CommandMeta = {
  summary: 'Return metadata for a single Azure AD / Microsoft 365 group. Use `--select` to slim large group payloads (the full group resource includes 30+ fields).',
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
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-group --group-id 'a1b2c3d4-...' --select 'id,displayName,mail'",
  responseShape: 'single Microsoft Graph `group` resource',
};

export { execute, meta, schema };
