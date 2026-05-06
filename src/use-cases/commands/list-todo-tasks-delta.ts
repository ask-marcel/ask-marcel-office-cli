import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ todoTaskListId: z.string().min(1) });
const { execute } = buildCommand((p) => `/me/todo/lists/${p.todoTaskListId}/tasks/delta()`, schema);

const meta: CommandMeta = {
  summary:
    'Track incremental task changes (added / updated / completed / deleted) within a single Microsoft To Do list. The first call returns the current snapshot plus `@odata.deltaLink`; subsequent calls with that link return only what has changed since.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/todo/lists/{todo-task-list-id}/tasks/delta()',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotask-delta',
  options: [
    {
      name: 'todo-task-list-id',
      key: 'todoTaskListId',
      required: true,
      description: 'Microsoft To Do task list ID. Returned by `list-todo-task-lists`.',
    },
  ],
  example: "ask-marcel list-todo-tasks-delta --todo-task-list-id 'AAMkAD...'",
  responseShape: 'collection of Microsoft Graph `todoTask` resources plus `@odata.deltaLink` / `@odata.nextLink`',
  pagination: true,
};

export { execute, meta, schema };
