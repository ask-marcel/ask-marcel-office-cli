import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/joinedTeams', baseSchema);

const meta: CommandMeta = {
  summary: 'List the Microsoft Teams the signed-in user is a member of.',
  category: 'teams',
  graphMethod: 'GET',
  graphPathTemplate: '/me/joinedTeams',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-list-joinedteams',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-joined-teams',
  responseShape: 'collection of Microsoft Graph `team` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
