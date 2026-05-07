import { z } from 'zod';
import { buildListCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { odataQueryOptions } from './odata-query.ts';

const baseSchema = z.object({ todoTaskListId: z.string().min(1), todoTaskId: z.string().min(1) });
const { execute, schema } = buildListCommand((p) => `/me/todo/lists/${p.todoTaskListId}/tasks/${p.todoTaskId}/linkedResources`, baseSchema);

const meta: CommandMeta = {
  summary: 'List the linked resources (URLs, emails, files) attached to a Microsoft To Do task.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/todo/lists/{todo-task-list-id}/tasks/{todo-task-id}/linkedResources',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotask-list-linkedresources',
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
    {
      name: 'todo-task-id',
      key: 'todoTaskId',
      required: true,
      description: 'To Do task ID. Returned by `ask-marcel list-todo-tasks`.',
      aliases: [{ name: 'task-id', key: 'taskId' }],
    },
    ...odataQueryOptions,
  ],
  example: "ask-marcel list-todo-linked-resources --todo-task-list-id 'AAMkAGI...' --todo-task-id 'AAMkABC...'",
  responseShape: 'collection of Microsoft Graph `linkedResource` resources under `value[]`',
  pagination: true,
};

export { execute, meta, schema };
