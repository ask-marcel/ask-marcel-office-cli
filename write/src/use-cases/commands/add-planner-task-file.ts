import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { LocalFileReader } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import { readLocalFile } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { addFileOp, checkNoFile, filePathOf, FILE_FOLDER, groupPlanRefusal, isGroupPlan, PLANNER_FILE_MAX_BYTES } from './planner-file.ts';
import { readPlannerItem } from './planner-reader.ts';
import { readTask } from './planner-task-reader.ts';
import { PLANNER_RESPONSE, PLANNER_SCOPES, TASK_ID_OPTION } from './planner-texts.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * Attaches a local file to a task of a plan that is not of a group: the file
 * goes into the folder "Planner attachments" of your OneDrive, and the task
 * gets a reference to it. A plan of a group is refused (its files are in the
 * document library of the group, which this CLI cannot write), and so is a
 * name that the folder has already.
 */

const schema = z.object({ taskId: z.string().min(1), file: z.string().min(1), alias: z.string().trim().min(1).max(255).optional() });

const plan: PlanWrite = async (reads, params, files?: LocalFileReader) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const task = await readTask(reads, parsed.value.taskId);
  if (!task.ok) return task;
  const owner = await readPlannerItem(reads, 'plan', task.value.planId);
  if (!owner.ok) return owner;
  if (isGroupPlan(owner.value)) return groupPlanRefusal();
  const file = await readLocalFile(files, parsed.value.file, PLANNER_FILE_MAX_BYTES);
  if (!file.ok) return file;
  const path = filePathOf(file.value.name);
  if (!path.ok) return path;
  const free = await checkNoFile(reads, path.value);
  if (!free.ok) return free;
  const details = await readPlannerItem(reads, 'details', task.value.item.id);
  if (!details.ok) return details;
  const before = { taskId: task.value.item.id, title: task.value.title, file: file.value.name, size: String(file.value.bytes.byteLength) };
  const spec = { path: path.value, name: file.value.name, bytes: file.value.bytes, alias: parsed.value.alias ?? file.value.name, details: details.value, before };
  return ok(todoPlan({ ops: [addFileOp(spec)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: `Attach a local file to a Planner task of a plan that is not of a group: the file goes into the folder "${FILE_FOLDER}" of your OneDrive (shared with nobody), and the task gets a reference to it. A plan of a group is refused. The undo removes the reference and moves the file to the recycle bin.`,
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/planner/tasks/{task-id}/details',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update',
  options: [
    TASK_ID_OPTION,
    {
      name: 'file',
      key: 'file',
      required: true,
      description: `A local file of at most ${PLANNER_FILE_MAX_BYTES} bytes. It goes into the folder "${FILE_FOLDER}" of your OneDrive under its own name; a name that the folder has already is refused. A relative path starts from the folder the CLI runs in.`,
    },
    { name: 'alias', key: 'alias', required: false, description: 'The name of the reference that Planner shows, at most 255 characters. The default is the name of the file.' },
  ],
  example: '{write-bin} add-planner-task-file --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --file ./budget.xlsx --dry-run',
  bodyTemplate: `PUT /me/drive/root:/${FILE_FOLDER}/<name of {file}>:/content, then { references: { <webUrl of the file, encoded>: { alias: '{alias}', type: 'Word' | 'Excel' | 'PowerPoint' | 'Other' } } } with If-Match: <ETag of the details>`,
  effect: 'task-write',
  undo: 'removes the reference, only if its name and its type did not change, and moves the file to the recycle bin of your OneDrive, only if its content did not change',
  scopesRequired: [...PLANNER_SCOPES, 'Files.ReadWrite'],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
