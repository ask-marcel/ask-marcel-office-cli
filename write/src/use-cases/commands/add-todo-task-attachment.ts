import { z } from 'zod';
import { ATTACHMENT_MAX_BYTES } from '@ask-marcel/office-core/domain/attachment-upload.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { readLocalFile } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { todoPlan } from './todo-batch.ts';
import { LIST_OPTION, listAndTask, TASK_ID_OPTION, TODO_RESPONSE } from './todo-command.ts';
import { addFileOp } from './todo-item-ops.ts';

/*
 * Attaches a local file to a To Do task. The plan reads the file, so a dry-run
 * refuses a file that is missing or too large. A file under 3 MB goes in one
 * POST, a larger one through an upload session (core attachment-upload.ts).
 */

const schema = z.object({ list: z.string().min(1), taskId: z.string().min(1), file: z.string().min(1) });

const plan: PlanWrite = async (reads, params, files) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const file = await readLocalFile(files, parsed.value.file, ATTACHMENT_MAX_BYTES['todo-task']);
  if (!file.ok) return file;
  const found = await listAndTask(reads, parsed.value);
  if (!found.ok) return found;
  return ok(todoPlan({ ops: [addFileOp(found.value.list, found.value.task, file.value)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Attach a local file to a Microsoft To Do task, up to 25 MB. A file of 3 MB or more goes through an upload session.',
  category: 'tasks',
  graphMethod: 'POST',
  graphPathTemplate: '/me/todo/lists/{list}/tasks/{task-id}/attachments (or attachments/createUploadSession for 3 MB or more)',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/todotask-post-attachments',
  options: [
    LIST_OPTION,
    TASK_ID_OPTION,
    {
      name: 'file',
      key: 'file',
      required: true,
      description: 'The local file to attach, at most 25 MB. A relative path starts from the folder the CLI runs in. The attachment has the name of the file.',
    },
  ],
  example: '{write-bin} add-todo-task-attachment --list "Tasks" --task-id "AAMkAD..." --file ./report.pdf --dry-run',
  bodyTemplate: "{ '@odata.type': '#microsoft.graph.taskFileAttachment', name: '<the name of {file}>', contentBytes: '<the bytes of {file}, base64>' }",
  effect: 'task-write',
  undo: 'removes the attachment, only if it is still there with its name',
  scopesRequired: ['Tasks.ReadWrite'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
