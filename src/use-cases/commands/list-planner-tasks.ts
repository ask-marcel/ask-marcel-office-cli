import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/planner/tasks', baseSchema);

const meta: CommandMeta = {
  summary: 'List every Microsoft Planner task assigned to or owned by the signed-in user, across all plans.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/planner/tasks',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/planneruser-list-tasks',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-planner-tasks',
  responseShape: 'collection of Microsoft Graph `plannerTask` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
