import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { UndoHandlers } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { AppliedChange, GraphReads, ItemIds } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { contentTypeOf } from '../../domain/attachment-content-type.ts';
import { graphIdSchema, isGraphId } from '../../domain/graph-id.ts';
import { referenceKeyOf } from '../../domain/planner-reference.ts';
import { referenceValue } from '../../domain/planner-snapshot.ts';
import { isNotThere } from './mail-folder-reader.ts';
import type { TodoOp } from './todo-batch.ts';
import { outcomeOf } from './todo-batch.ts';
import type { PlannerItem } from './planner-reader.ts';
import { patchOp, sameItem } from './planner-patch.ts';
import { unreadable } from './todo-reader.ts';
import { createTodoUndoHandler, notThere } from './todo-undo-handler.ts';
import { refused } from './planner-reader.ts';

/*
 * A file of a Planner task. Planner keeps a file of a task as a reference to
 * a file in SharePoint or OneDrive: the Planner app puts the file of a group
 * plan in the document library of the group, which the token of this CLI
 * cannot write. So this CLI attaches a file only to a plan that is not of a
 * group: it puts the file in the folder "Planner attachments" of the OneDrive
 * of the user (shared with nobody), then adds a reference to it. The run
 * keeps two changes: the undo removes the reference, then moves the file to
 * the recycle bin of OneDrive if its content did not change.
 */

/** The OneDrive folder of the files of Planner tasks. */
export const FILE_FOLDER = 'Planner attachments';

/** The largest file that this CLI attaches to a task. */
export const PLANNER_FILE_MAX_BYTES = 25_000_000;

// OneDrive refuses these characters in a name; the name is also encoded before it goes into the path.
const BAD_NAME = /["*:<>?/\\|]/;

/** The OneDrive path of the file `name`, after the check of the name. */
export const filePathOf = (name: string): Result<string, GraphError> => {
  if (name.length > 200 || BAD_NAME.test(name) || name.trim() !== name || name.endsWith('.'))
    return refused(
      'planner_file_name_invalid',
      '--file: OneDrive cannot keep a file of this name. Rename the file: no " * : < > ? / \\ |, no space at the start or the end, at most 200 characters.'
    );
  return ok(`/me/drive/root:/${encodeURIComponent(FILE_FOLDER)}/${encodeURIComponent(name)}`);
};

/** Whether the plan is of a group (its container, or the owner of an older plan). */
export const isGroupPlan = (plan: PlannerItem): boolean => {
  const container = plan.fields['container'];
  if (container !== null && typeof container === 'object') return (container as { type?: unknown }).type === 'group';
  return typeof plan.fields['owner'] === 'string';
};

/** The refusal of a file of a group plan. */
export const groupPlanRefusal = (): Result<never, GraphError> =>
  refused(
    'planner_group_plan_file',
    'The plan is of a group: Planner keeps its files in the document library of the group, which this CLI cannot write. Attach the file in Planner, or give a link to a file that is in SharePoint or OneDrive already with add-planner-task-link.'
  );

/** Refuses a file that the OneDrive folder has already: a write would replace it. */
export const checkNoFile = async (reads: GraphReads, path: string): Promise<Result<undefined, GraphError>> => {
  const read = await reads.get(path);
  if (read.ok)
    return refused(
      'planner_file_exists',
      `--file: the folder "${FILE_FOLDER}" of your OneDrive has a file of this name. Rename the file, or give a link to that file with add-planner-task-link.`
    );
  return isNotThere(read.error) ? ok(undefined) : read;
};

const TYPES: Readonly<Record<string, string>> = { doc: 'Word', docx: 'Word', xls: 'Excel', xlsx: 'Excel', ppt: 'PowerPoint', pptx: 'PowerPoint' };

const typeOf = (name: string): string => TYPES[name.slice(name.lastIndexOf('.') + 1).toLowerCase()] ?? 'Other';

const uploadedSchema = z.object({ id: z.string(), webUrl: z.url(), cTag: z.string().optional() }).loose();

/** The file to put and the details of the task to link it from. */
export type FileSpec = {
  readonly path: string;
  readonly name: string;
  readonly bytes: Uint8Array;
  readonly alias: string;
  readonly details: PlannerItem;
  readonly before: ItemIds;
};

/** The op that puts the file in OneDrive, then adds its reference to the task. */
export const addFileOp = (spec: FileSpec): TodoOp => ({
  change: { action: 'add-planner-task-file', target: spec.path, before: spec.before },
  recheck: async (reads) => sameItem(reads, 'details', spec.details),
  write: async (graph) => {
    const put = await graph.put(spec.path, spec.bytes, contentTypeOf(spec.name));
    if (!put.ok) return outcomeOf(put);
    const file = uploadedSchema.safeParse(put.value);
    if (!file.success || !isGraphId(file.data.id)) return outcomeOf(err(unreadable('the file it put in OneDrive')));
    const uploaded: AppliedChange = {
      action: 'upload-planner-task-file',
      target: spec.path,
      before: spec.before,
      after: { itemId: file.data.id, webUrl: file.data.webUrl },
      undo: { kind: 'delete-uploaded-planner-file', data: { itemId: file.data.id, ...(file.data.cTag === undefined ? {} : { cTag: file.data.cTag }) } },
    };
    const body = { references: { [referenceKeyOf(file.data.webUrl)]: referenceValue({ alias: spec.alias, type: typeOf(spec.name) }) } };
    const linked = await patchOp({ action: 'add-planner-task-file', resource: 'details', item: spec.details, body, before: { ...spec.before, url: file.data.webUrl } }).write(
      graph
    );
    return { changes: [uploaded, ...linked.changes], ...(linked.error === undefined ? {} : { error: linked.error }) };
  },
});

const fileSchema = z.object({ itemId: graphIdSchema, cTag: z.string().optional() });
const readSchema = z.object({ cTag: z.string().optional() }).loose();

/** The undo of the file put in OneDrive: to the recycle bin, only if its content did not change. */
export const FILE_UNDO_HANDLERS: UndoHandlers = {
  'delete-uploaded-planner-file': createTodoUndoHandler(
    fileSchema,
    async (reads, data) => {
      const read = await reads.get(`/me/drive/items/${data.itemId}`);
      if (!read.ok) return notThere(read.error);
      const now = readSchema.safeParse(read.value);
      return ok(now.success && now.data.cTag === data.cTag ? 'ready' : 'changed');
    },
    async (graph, data) => graph.delete(`/me/drive/items/${data.itemId}`)
  ),
};
