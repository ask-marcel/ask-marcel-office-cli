import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/memberOf', baseSchema);

const meta: CommandMeta = {
  summary:
    "List the groups, directory roles, and administrative units the signed-in user is a member of. Each entry's `@odata.type` distinguishes #microsoft.graph.group from #microsoft.graph.directoryRole, etc.",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me/memberOf',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-memberof',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-my-memberships',
  responseShape: 'collection of Microsoft Graph `directoryObject` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
