import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { UndoCheck, UndoHandlers } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { graphId, graphIdSchema, parseGraphId } from '../../domain/graph-id.ts';
import type { MailMessage, Written } from './mail-batch.ts';
import { createFolderReader, isNotThere } from './mail-folder-reader.ts';
import type { MailFolder } from './mail-folder-reader.ts';
import { refuseFolderInRefused } from './mail-folder-target.ts';

/*
 * The move of one message (`POST /me/messages/{id}/move`) and its undo. A moved
 * message gets a new id, so the run keeps the id after the move, and the undo
 * moves the message with that id back to the folder of before (where it gets a
 * new id again). The undo of a move runs only if the message is still in the
 * folder the run put it in, and never moves mail into a refused folder. The
 * undo of delete-mail-draft moves the draft from Deleted Items back to Drafts.
 */

const movedSchema = z.object({ id: z.string(), parentFolderId: z.string() });

/** The message after a move: its new id, and the id of its folder. */
export type MovedMessage = { readonly id: GraphId; readonly folderId: string };

const noNewId: GraphError = { type: 'validation_error', message: 'Graph moved the message but gave no new id.' };

/** Moves the message to the folder `destinationId` (an id or a well-known name), and answers its new id. */
export const moveMessage = async (graph: WriteGraph, id: GraphId, destinationId: GraphId): Promise<Result<MovedMessage, GraphError>> => {
  const moved = await graph.post(`/me/messages/${id}/move`, { destinationId });
  if (!moved.ok) return moved;
  const parsed = movedSchema.safeParse(moved.value);
  if (!parsed.success) return err(noNewId);
  const newId = parseGraphId(parsed.data.id, 'moved message');
  return newId.ok ? ok({ id: newId.value, folderId: parsed.data.parentFolderId }) : newId;
};

/** The move of one message of a batch to `to`, with the undo that moves it back. */
export const moveTo = async (graph: WriteGraph, message: MailMessage, to: MailFolder): Promise<Result<Written, GraphError>> => {
  const moved = await moveMessage(graph, message.id, to.id);
  if (!moved.ok) return moved;
  return ok({
    after: { messageId: moved.value.id, folderId: to.id, folder: to.displayName },
    undo: { kind: 'move-message-back', data: { messageId: moved.value.id, folderId: to.id, backTo: message.parentFolderId } },
  });
};

const moveBackSchema = z.object({ messageId: graphIdSchema, folderId: graphIdSchema, backTo: graphIdSchema });

const draftBackSchema = z.object({ messageId: graphIdSchema, folderId: graphIdSchema });

const DRAFTS = graphId('drafts');

const notAnId: GraphError = { type: 'validation_error', message: 'The run journal names an id that is not a Graph id, so this undo does not run.' };

const placeSchema = z.object({ parentFolderId: z.string(), isDraft: z.boolean().optional() });

// `ready` when the message is still in the folder the run put it in (and, for a draft, still a draft).
const checkPlace = async (reads: GraphReads, messageId: string, folderId: string, draft: boolean): Promise<Result<UndoCheck, GraphError>> => {
  const read = await reads.get(`/me/messages/${messageId}?$select=parentFolderId,isDraft`);
  if (!read.ok) return isNotThere(read.error) ? ok('missing') : read;
  const place = placeSchema.safeParse(read.value);
  const here = place.success && place.data.parentFolderId === folderId && (!draft || place.data.isDraft === true);
  return ok(here ? 'ready' : 'changed');
};

// The new ids, as the undo report shows them (undo-run.ts).
const movedTo = async (graph: WriteGraph, messageId: GraphId, destinationId: GraphId): Promise<Result<unknown, GraphError>> => {
  const moved = await moveMessage(graph, messageId, destinationId);
  return moved.ok ? ok({ now: { messageId: moved.value.id, folderId: moved.value.folderId } }) : moved;
};

// The folder of before can be a refused folder (a message moved out of Junk Email): the undo never moves mail there.
const moveBack = async (graph: WriteGraph, messageId: GraphId, backTo: GraphId): Promise<Result<unknown, GraphError>> => {
  const folders = createFolderReader(graph);
  const folder = await folders.folder(backTo);
  const allowed = folder.ok ? await refuseFolderInRefused(folders, folder.value, 'undo') : folder;
  return allowed.ok ? movedTo(graph, messageId, allowed.value.id) : allowed;
};

/** The undo handlers of the moves, by kind. */
export const MOVE_UNDO_HANDLERS: UndoHandlers = {
  'move-message-back': {
    check: async (reads, data) => {
      const step = moveBackSchema.safeParse(data);
      return step.success ? checkPlace(reads, step.data.messageId, step.data.folderId, false) : err(notAnId);
    },
    apply: async (graph, data) => {
      const step = moveBackSchema.safeParse(data);
      return step.success ? moveBack(graph, step.data.messageId, step.data.backTo) : err(notAnId);
    },
  },
  'move-draft-to-drafts': {
    check: async (reads, data) => {
      const step = draftBackSchema.safeParse(data);
      return step.success ? checkPlace(reads, step.data.messageId, step.data.folderId, true) : err(notAnId);
    },
    apply: async (graph, data) => {
      const step = draftBackSchema.safeParse(data);
      return step.success ? movedTo(graph, step.data.messageId, DRAFTS) : err(notAnId);
    },
  },
};
