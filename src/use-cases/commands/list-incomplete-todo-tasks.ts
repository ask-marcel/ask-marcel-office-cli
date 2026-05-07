import { z } from 'zod';
import { buildCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';

const schema = z.object({ todoTaskListId: z.string().min(1) });
const { execute } = buildCommand((p) => `/me/todo/lists/${p.todoTaskListId}/tasks?$filter=status ne 'completed'`, schema);

const meta: CommandMeta = {
  summary: 'List every incomplete Microsoft To Do task in a given list (status not equal to `completed`).',
  category: 'tasks',
  graphMethod: 'GET',
  graphPathTemplate: "/me/todo/lists/{todo-task-list-id}/tasks?$filter=status ne 'completed'",
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotasklist-list-tasks',
  options: [
    {
      name: 'todo-task-list-id',
      key: 'todoTaskListId',
      required: true,
      description:
        'todoTaskList ID. Returned by `ask-marcel list-todo-task-lists`. Well-known names also work, e.g. `tasks` for the default list. ' +
        'There is no Graph endpoint that returns incomplete tasks across every list — call this once per list.',
      aliases: [
        { name: 'task-list-id', key: 'taskListId' },
        { name: 'todo-list-id', key: 'todoListId' },
      ],
    },
  ],
  example: "ask-marcel list-incomplete-todo-tasks --todo-task-list-id 'tasks'",
  responseShape: 'collection of Microsoft Graph `todoTask` resources under `value[]` where `status != "completed"`',
  pagination: true,
};

export { execute, meta, schema };
