import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ todoTaskListId: z.string().min(1), todoTaskId: z.string().min(1) });
const { execute } = buildCommand((p) => `/me/todo/lists/${p.todoTaskListId}/tasks/${p.todoTaskId}`, schema);

const meta: CommandMeta = {
  summary: 'Get a single Microsoft To Do task by its ID and its parent list ID.',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: '/me/todo/lists/{todo-task-list-id}/tasks/{todo-task-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotask-get',
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
  ],
  example: "ask-marcel get-todo-task --todo-task-list-id 'AAMkAGI...' --todo-task-id 'AAMkABC...'",
  responseShape: 'single Microsoft Graph `todoTask` resource',
};

export { execute, meta, schema };
