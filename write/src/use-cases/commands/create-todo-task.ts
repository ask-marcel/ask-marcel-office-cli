import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { LIST_OPTION, TODO_RESPONSE } from './todo-command.ts';
import { resolveList } from './todo-reader.ts';
import { createTaskOp } from './todo-task-ops.ts';
import { fieldsOf, TASK_FIELD_OPTIONS, taskFieldShape, taskInputOf } from './todo-task-options.ts';

/*
 * Creates a task in a To Do list, with its title and the other fields that
 * the options give. The undo deletes the task if it did not change after the
 * run.
 */

const schema = z.object({ list: z.string().min(1), ...taskFieldShape, title: z.string().trim().min(1) });

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const list = await resolveList(reads, parsed.value.list, '--list');
  if (!list.ok) return list;
  const fields = await fieldsOf(reads, taskInputOf(parsed.value));
  if (!fields.ok) return fields;
  return ok(todoPlan({ ops: [createTaskOp(list.value, fields.value)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Create a task in a Microsoft To Do list: the title, and the note, the status, the importance, the dates, the reminder and the categories that the options give.',
  category: 'tasks',
  graphMethod: 'POST',
  graphPathTemplate: '/me/todo/lists/{list}/tasks',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotasklist-post-tasks',
  options: [LIST_OPTION, ...TASK_FIELD_OPTIONS.map((option) => (option.name === 'title' ? { ...option, required: true } : option))],
  example: '{write-bin} create-todo-task --list "Tasks" --title "Send the report" --due 2026-11-20 --dry-run',
  bodyTemplate:
    "{ title: '{title}', body?: { content: '{body}', contentType: 'text' }, status?: '{status}', importance?: '{importance}', categories?: [{categories}], dueDateTime?: { dateTime: '{due}T00:00:00', timeZone: <time zone of the mailbox> }, startDateTime?: { dateTime: '{start}T00:00:00', timeZone: <time zone of the mailbox> }, reminderDateTime?: { dateTime: '{reminder}:00', timeZone: <time zone of the mailbox> }, isReminderOn?: true }",
  effect: 'task-write',
  undo: 'deletes the task, only if it did not change after the run',
  scopesRequired: ['Tasks.ReadWrite', 'MailboxSettings.Read'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
