import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { LIST_OPTION, listAndTask, TASK_ID_OPTION, TODO_RESPONSE } from './todo-command.ts';
import { updateTaskOp } from './todo-task-ops.ts';
import { fieldsOf, TASK_FIELD_OPTIONS, taskFieldShape, taskInputOf } from './todo-task-options.ts';

/*
 * Changes the fields of a To Do task that the options give; the other fields
 * stay. The undo writes the fields of before back, only if each field still
 * has the value that the run set.
 */

const schema = z.object({ list: z.string().min(1), taskId: z.string().min(1), ...taskFieldShape });

const noField: GraphError = {
  type: 'validation_error',
  message: 'update-todo-task: give one field to change or more: --title, --body, --status, --importance, --due, --start, --reminder or --categories.',
};

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const input = taskInputOf(parsed.value);
  if (Object.keys(input).length === 0) return err(noField);
  const found = await listAndTask(reads, parsed.value);
  if (!found.ok) return found;
  const fields = await fieldsOf(reads, input);
  if (!fields.ok) return fields;
  return ok(todoPlan({ ops: [updateTaskOp('update-task', found.value.list, found.value.task, fields.value)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Change the fields of a Microsoft To Do task that the options give: the title, the note, the status, the importance, the dates, the reminder or the categories. The other fields stay.',
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/me/todo/lists/{list}/tasks/{task-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotask-update',
  options: [LIST_OPTION, TASK_ID_OPTION, ...TASK_FIELD_OPTIONS],
  example: '{write-bin} update-todo-task --list "Tasks" --task-id "AAMkAD..." --due none --importance high --dry-run',
  bodyTemplate:
    "{ title?: '{title}', body?: { content: '{body}', contentType: 'text' }, status?: '{status}', importance?: '{importance}', categories?: [{categories}], dueDateTime?: { dateTime: '{due}T00:00:00', timeZone: <time zone of the mailbox> }, startDateTime?: { dateTime: '{start}T00:00:00', timeZone: <time zone of the mailbox> }, reminderDateTime?: { dateTime: '{reminder}:00', timeZone: <time zone of the mailbox> }, isReminderOn?: true }",
  effect: 'task-write',
  undo: 'writes the fields of before back, only if each field still has the value that the run set',
  scopesRequired: ['Tasks.ReadWrite', 'MailboxSettings.Read'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
