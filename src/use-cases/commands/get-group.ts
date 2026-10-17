import { z } from 'zod';
import { buildSelectableCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { selectExpandOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

const baseSchema = z.object({ groupId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/groups/${p.groupId}`, baseSchema);

const meta: ReadCommandMeta = {
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
  example: "{read-bin} get-group --group-id 'a1b2c3d4-...' --select 'id,displayName,mail'",
  responseShape: 'single Microsoft Graph `group` resource',
  effect: 'read',
  scopesRequired: ['GroupMember.Read.All'],
};

export { execute, meta, schema };
