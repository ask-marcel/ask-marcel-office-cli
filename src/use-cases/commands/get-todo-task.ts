import { z } from 'zod';
import { err } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';
import { appendOData, selectExpandOptions, selectExpandSchema } from './odata-query.ts';
import { rewriteTodoTitleQuirk } from './todo-parse-uri-rewrite.ts';

// Audit v1.0.0 §B9: sibling single-resource GETs (get-my-manager,
// get-user-manager, get-mail-message, etc.) all expose `--select`/`--expand`
// so an LLM can slim a fetched resource. This command was the only Microsoft
// task-list "get" without them. It also hits the same `/tasks` endpoint as its
// list sibling, so a `--select` that includes `title` trips the same
// RequestBroker--ParseUri quirk — rewrite it via the shared helper.
const schema = z.object({ todoTaskListId: z.string().min(1), todoTaskId: z.string().min(1) }).extend(selectExpandSchema.shape);

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const path = appendOData(`/me/todo/lists/${parsed.data.todoTaskListId}/tasks/${parsed.data.todoTaskId}`, parsed.data);
  const result = await graph.get(path);
  if (result.ok) return result;
  const rewritten = rewriteTodoTitleQuirk(result.error, parsed.data);
  return rewritten ? err(rewritten) : result;
};

const meta: CommandMeta = {
  summary:
    'Get a single Microsoft To Do task by its ID and its parent list ID. Use `--select` to slim the response (e.g. `--select id,status`) or `--expand checklistItems` / `--expand linkedResources` to inline child collections. Known Graph quirk: any `--select` combo that includes `title` trips `RequestBroker--ParseUri` on this endpoint; the CLI rewrites that opaque error to a hint.',
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
