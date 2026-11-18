import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { LIST_OPTION, listAndTask, TASK_ID_OPTION, TODO_RESPONSE } from './todo-command.ts';
import { findAttachment, removeFileOp } from './todo-item-ops.ts';

/*
 * Removes a file attachment of a To Do task. Before the delete, the run reads
 * the bytes of the file, and the run journal keeps them, so the undo can
 * attach the file again.
 */

const schema = z.object({ list: z.string().min(1), taskId: z.string().min(1), attachmentId: z.string().min(1) });

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const found = await listAndTask(reads, parsed.value);
  if (!found.ok) return found;
  const attachment = await findAttachment(reads, found.value.task, parsed.value.attachmentId);
  if (!attachment.ok) return attachment;
  return ok(todoPlan({ ops: [removeFileOp(found.value.list, found.value.task, attachment.value)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Remove a file attachment of a Microsoft To Do task. The run journal keeps the bytes of the file (in ~/.ask-marcel/runs, for 30 days), so its undo can attach the file again. Over MCP it runs on run-destructive-command.',
  category: 'tasks',
  graphMethod: 'DELETE',
  graphPathTemplate: '/me/todo/lists/{list}/tasks/{task-id}/attachments/{attachment-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/taskfileattachment-delete',
  options: [
    LIST_OPTION,
    TASK_ID_OPTION,
    {
      name: 'attachment-id',
      key: 'attachmentId',
      required: true,
      description: 'The id of the attachment, as get-todo-task with --expand attachments returns it.',
      argumentHint: { kind: 'idOrName' },
    },
  ],
  example: '{write-bin} remove-todo-task-attachment --list "Tasks" --task-id "AAMkAD..." --attachment-id "AAMkAD..." --dry-run',
  effect: 'delete',
  undo: 'attaches the file again from the bytes in the run journal. It is a new attachment with a new id',
  scopesRequired: ['Tasks.ReadWrite'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
