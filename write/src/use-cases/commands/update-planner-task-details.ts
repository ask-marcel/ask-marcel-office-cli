import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { LocalFileReader } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { PREVIEW_TYPES } from '../../domain/planner-fields.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { checklistBodyOf, checklistOf } from './planner-checklist-ops.ts';
import { patchOp } from './planner-patch.ts';
import { readDetails } from './planner-task-reader.ts';
import { PLANNER_RESPONSE, PLANNER_SCOPES, TASK_ID_OPTION } from './planner-texts.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * Changes the details of a task: the plain-text description (Graph v1.0), the
 * HTML notes (Graph beta, the one beta path of this CLI, decision D7), the
 * preview of the card, and the checklist from a file. Graph refuses a change
 * of the description and the notes together, so this command refuses it
 * first. The undo writes the details of before back; the notes come back as
 * plain text, as Graph v1.0 reads no HTML notes.
 */

const schema = z.object({
  taskId: z.string().min(1),
  description: z.string().max(100_000).optional(),
  notesHtml: z.string().min(1).max(100_000).optional(),
  checklistFile: z.string().min(1).optional(),
  preview: z.enum(PREVIEW_TYPES).optional(),
});

type Params = z.infer<typeof schema>;

const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

const badParams = (p: Params): Result<never, GraphError> | undefined => {
  if (p.description !== undefined && p.notesHtml !== undefined)
    return refused('planner_description_and_notes', 'Graph refuses a change of the description and the notes in one request: give --description or --notes-html.');
  if (p.description === undefined && p.notesHtml === undefined && p.checklistFile === undefined && p.preview === undefined)
    return refused('planner_nothing_to_change', 'Give --description, --notes-html, --checklist-file or --preview: there is nothing to change.');
  return undefined;
};

const textFieldsOf = (p: Params): Record<string, unknown> => ({
  ...(p.description === undefined ? {} : { description: p.description }),
  ...(p.notesHtml === undefined ? {} : { notes: { contentType: 'html', content: p.notesHtml } }),
  ...(p.preview === undefined ? {} : { previewType: p.preview }),
});

const plan: PlanWrite = async (reads, params, files?: LocalFileReader) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const bad = badParams(parsed.value);
  if (bad !== undefined) return bad;
  const details = await readDetails(reads, parsed.value.taskId);
  if (!details.ok) return details;
  const checklist = await checklistBodyOf(files, parsed.value.checklistFile, checklistOf(details.value));
  if (!checklist.ok) return checklist;
  const body = { ...textFieldsOf(parsed.value), ...checklist.value.body };
  if (Object.keys(body).length === 0) return refused('planner_nothing_to_change', 'The checklist of the file is the checklist of the task: there is nothing to change.');
  const before = { taskId: details.value.id, fields: Object.keys(body).join(', '), ...checklist.value.shown };
  return ok(todoPlan({ ops: [patchOp({ action: 'update-planner-task-details', resource: 'details', item: details.value, body, before })] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Change the details of a Planner task: the description as plain text (--description) or the notes as HTML (--notes-html, on Graph beta), the preview of the card, and the checklist from a file (--checklist-file). The undo writes the details of before back.',
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/planner/tasks/{task-id}/details',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update',
  options: [
    TASK_ID_OPTION,
    {
      name: 'description',
      key: 'description',
      required: false,
      description: 'The description as plain text. It replaces the description and the notes of before. Not with --notes-html.',
    },
    {
      name: 'notes-html',
      key: 'notesHtml',
      required: false,
      description:
        'The notes as HTML, sent to Graph beta (v1.0 has no notes). Planner keeps their plain text as the description. Not with --description: Graph refuses both in one request.',
    },
    {
      name: 'checklist-file',
      key: 'checklistFile',
      required: false,
      description:
        'A local UTF-8 text file with the checklist: one item on each line, at most 20 items of at most 100 characters, "[x] " before a checked item and "[ ] " before an item that is not checked. An item of the task with the same title stays, the other items of the task are removed, and the new titles are added. At most 64000 bytes. A relative path starts from the folder the CLI runs in.',
    },
    { name: 'preview', key: 'preview', required: false, description: 'What the card of the task shows on the board.', argumentHint: { kind: 'magicValue', values: PREVIEW_TYPES } },
  ],
  example: '{write-bin} update-planner-task-details --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --notes-html "<p>See the <b>budget</b></p>" --checklist-file steps.txt --dry-run',
  bodyTemplate:
    "{ description?: '{description}' | notes?: { contentType: 'html', content: '{notes-html}' }, previewType?: '{preview}', checklist?: { <key>: { title, isChecked } | null } (the items of {checklist-file}) } with If-Match: <ETag of the details>",
  effect: 'task-write',
  undo: 'writes the description, the preview and the checklist of before back (a removed item comes back with a new id), only if the values that the run set did not change. HTML notes come back as their plain text',
  scopesRequired: [...PLANNER_SCOPES],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
