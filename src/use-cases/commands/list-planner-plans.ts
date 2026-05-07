import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/planner/plans', baseSchema);

const meta: CommandMeta = {
  summary:
    'List every Microsoft Planner plan the signed-in user has access to (across every group). Use this to discover plan IDs without needing an existing task as the entry point.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/planner/plans',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/planneruser-list-plans',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-planner-plans',
  responseShape: 'collection of Microsoft Graph `plannerPlan` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
