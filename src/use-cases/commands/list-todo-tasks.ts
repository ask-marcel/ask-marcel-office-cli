import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ todoTaskListId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/me/todo/lists/${p.todoTaskListId}/tasks`, baseSchema);

const meta: CommandMeta = {
  summary: 'List every task in a single Microsoft To Do task list, regardless of completion status. Use `list-incomplete-todo-tasks` if you only want the open ones.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/todo/lists/{todo-task-list-id}/tasks',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotasklist-list-tasks',
  options: [
    {
      name: 'todo-task-list-id',
      key: 'todoTaskListId',
      required: true,
      description: 'To Do task list ID. Returned by `ask-marcel list-todo-task-lists`.',
      aliases: [
        { name: 'task-list-id', key: 'taskListId' },
        { name: 'todo-list-id', key: 'todoListId' },
      ],
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-todo-tasks --todo-task-list-id 'AAMkAGI...'",
  responseShape: 'collection of Microsoft Graph `todoTask` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
