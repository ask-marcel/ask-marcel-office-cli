import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { UndoHandlers } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { AppliedChange, PlannedChange } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { graphIdSchema, isGraphId } from '../../domain/graph-id.ts';
import type { MailFolder } from './mail-folder-reader.ts';
import type { TodoOp } from './todo-batch.ts';
import { outcomeOf } from './todo-batch.ts';
import { checkRead, createTodoUndoHandler } from './todo-undo-handler.ts';

/*
 * The create of a mail folder (`POST /me/mailFolders/{parent}/childFolders`)
 * and its undo. The run journal keeps the id, the name and the parent of the
 * new folder. The undo deletes the folder only while it is as the run made
 * it: the same name, the same parent, no message and no child folder.
 */

const createdSchema = z.object({ id: z.string(), displayName: z.string() });

const noNewId: GraphError = {
  type: 'validation_error',
  message: 'Graph made the folder but gave no id that this CLI can read, so the run has no undo for it. Run list-mail-folders to find it.',
};

const createFolder = async (graph: WriteGraph, change: PlannedChange, parent: MailFolder, name: string): Promise<Result<AppliedChange, GraphError>> => {
  const created = await graph.post(change.target, { displayName: name });
  if (!created.ok) return created;
  const folder = createdSchema.safeParse(created.value);
  if (!folder.success || !isGraphId(folder.data.id)) return err(noNewId);
  const made = { folderId: folder.data.id, name: folder.data.displayName, parentFolderId: parent.id };
  return ok({ ...change, after: made, undo: { kind: 'delete-created-mail-folder', data: made } });
};

/** The op that makes the folder `name` in `parent`. */
export const createFolderOp = (parent: MailFolder, name: string): TodoOp => {
  const change: PlannedChange = {
    action: 'create-mail-folder',
    target: `/me/mailFolders/${parent.id}/childFolders`,
    before: { parentFolderId: parent.id, parent: parent.displayName, name },
  };
  return { change, write: async (graph) => outcomeOf(await createFolder(graph, change, parent, name)) };
};

const madeSchema = z.object({ folderId: graphIdSchema, name: z.string(), parentFolderId: z.string() });

const nowSchema = z.object({ displayName: z.string(), parentFolderId: z.string(), childFolderCount: z.number(), totalItemCount: z.number() });

// Whether the folder is still as the run made it, and empty.
const isAsMade = (value: unknown, made: z.infer<typeof madeSchema>): boolean => {
  const now = nowSchema.safeParse(value);
  if (!now.success) return false;
  const { displayName, parentFolderId, childFolderCount, totalItemCount } = now.data;
  return displayName === made.name && parentFolderId === made.parentFolderId && childFolderCount === 0 && totalItemCount === 0;
};

const SELECT = '$select=displayName,parentFolderId,childFolderCount,totalItemCount';

/** The undo of a created folder: it deletes the folder while it is empty and as the run made it. */
export const MAIL_FOLDER_UNDO_HANDLERS: UndoHandlers = {
  'delete-created-mail-folder': createTodoUndoHandler(
    madeSchema,
    async (reads, made) => checkRead(await reads.get(`/me/mailFolders/${made.folderId}?${SELECT}`), (value) => isAsMade(value, made)),
    async (graph, made) => graph.delete(`/me/mailFolders/${made.folderId}`)
  ),
};
