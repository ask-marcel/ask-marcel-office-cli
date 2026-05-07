import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildListCommand(() => '/me/todo/lists', baseSchema);

const meta: CommandMeta = {
  summary: 'List the signed-in user’s Microsoft To Do task lists (e.g. `Tasks`, `Flagged Emails`, custom lists).',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/todo/lists',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todo-list-lists',
  options: [...odataQueryOptions],
  example: 'ask-marcel list-todo-task-lists',
  responseShape: 'collection of Microsoft Graph `todoTaskList` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
