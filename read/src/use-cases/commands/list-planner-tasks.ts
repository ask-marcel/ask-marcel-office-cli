import { z } from 'zod';
import { buildListCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { odataQueryOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/planner/tasks', baseSchema);

const meta: ReadCommandMeta = {
  summary: 'List every Microsoft Planner task assigned to or owned by the signed-in user, across all plans.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/planner/tasks',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/planneruser-list-tasks',
  options: [...odataQueryOptions],
  example: '{read-bin} list-planner-tasks',
  responseShape: 'collection of Microsoft Graph `plannerTask` resources under `value[]`',
  pagination: true,
  effect: 'read',
  scopesRequired: ['Tasks.Read'],
};

export { execute, meta, schema };
