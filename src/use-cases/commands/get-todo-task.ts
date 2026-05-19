import { z } from 'zod';
import { buildSelectableCommand } from './build-command.ts';
import type { CommandMeta } from './command-types.ts';
import { selectExpandOptions } from './odata-query.ts';

// Audit v1.0.0 §B9: sibling single-resource GETs (get-my-manager,
// get-user-manager, get-mail-message, etc.) all expose `--select`/`--expand`
// so an LLM can slim a fetched resource. This command was the only Microsoft
// task-list "get" without them — switch to the selectable builder.
const baseSchema = z.object({ todoTaskListId: z.string().min(1), todoTaskId: z.string().min(1) });
const { execute, schema } = buildSelectableCommand((p) => `/me/todo/lists/${p.todoTaskListId}/tasks/${p.todoTaskId}`, baseSchema);

const meta: CommandMeta = {
  summary:
    'Get a single Microsoft To Do task by its ID and its parent list ID. Use `--select` to slim the response (e.g. `--select id,title,status`) or `--expand checklistItems` / `--expand linkedResources` to inline child collections.',
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
      description:
        "To Do task ID. Returned by `ask-marcel list-todo-tasks`. Accepts `--task-id` as a shorter alias (within this command's flag set the To Do context is unambiguous).",
      aliases: [{ name: 'task-id', key: 'taskId' }],
    },
    ...selectExpandOptions,
  ],
  example: "ask-marcel get-todo-task --todo-task-list-id 'AAMkAGI...' --todo-task-id 'AAMkABC...'",
  responseShape: 'single Microsoft Graph `todoTask` resource (slimmed by `--select` when supplied)',
};

export { execute, meta, schema };
