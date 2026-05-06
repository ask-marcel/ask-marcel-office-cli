import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({}).strict();
const { execute } = buildCommand(() => '/me/transitiveMemberOf', schema);

const meta: CommandMeta = {
  summary:
    'List all groups, directory roles, and administrative units the signed-in user is a member of *transitively* — including memberships inherited via nested groups. Sibling to `list-my-memberships` (`/me/memberOf`) which only returns direct memberships.',
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me/transitiveMemberOf',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-transitivememberof',
  options: [],
  example: 'ask-marcel list-my-transitive-memberships',
  responseShape: 'collection of Microsoft Graph `directoryObject` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
