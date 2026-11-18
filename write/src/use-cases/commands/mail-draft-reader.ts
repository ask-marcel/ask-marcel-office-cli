import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { graphIdSchema } from '../../domain/graph-id.ts';
import { unreadable } from './todo-reader.ts';

/*
 * The reads of the attachment commands of a mail draft. A command changes a
 * message only while it is an unsent draft that is not in the Outbox (a draft
 * in the Outbox is on its way out). The guard reads the message before any
 * write, again just before the write, and in each undo check.
 */

/** Whether the message is an unsent draft, a message that was sent or received, or a draft in the Outbox. */
export type DraftState = 'draft' | 'not-draft' | 'outbox';

/** An attachment as the list of the attachments of a message gives it: no bytes, and a size larger than the file. */
export type DraftAttachment = { readonly id: GraphId; readonly name: string; readonly size: number };

const messageSchema = z.object({ isDraft: z.boolean().optional(), parentFolderId: z.string() });
const folderSchema = z.object({ id: z.string() });
const listSchema = z.object({ value: z.array(z.object({ id: graphIdSchema, name: z.string(), size: z.number() })) });

/** The path of the attachments of the message. */
export const attachmentsOf = (messageId: GraphId): string => `/me/messages/${messageId}/attachments`;

/** The state of the message: a 404 when it is not there. */
export const readDraftState = async (reads: GraphReads, messageId: GraphId): Promise<Result<DraftState, GraphError>> => {
  const read = await reads.get(`/me/messages/${messageId}?$select=isDraft,parentFolderId`);
  if (!read.ok) return read;
  const message = messageSchema.safeParse(read.value);
  if (!message.success) return err(unreadable('the message'));
  if (message.data.isDraft !== true) return ok('not-draft');
  const outbox = await reads.get('/me/mailFolders/outbox?$select=id');
  if (!outbox.ok) return outbox;
  const folder = folderSchema.safeParse(outbox.value);
  if (!folder.success) return err(unreadable('the Outbox folder'));
  return ok(folder.data.id === message.data.parentFolderId ? 'outbox' : 'draft');
};

const REFUSALS: Readonly<Record<Exclude<DraftState, 'draft'>, GraphError>> = {
  'not-draft': {
    type: 'validation_error',
    code: 'not_a_draft',
    message: '--message-id: the message is not a draft (it was sent or received). This command changes an unsent draft only.',
  },
  outbox: {
    type: 'validation_error',
    code: 'draft_in_outbox',
    message: '--message-id: the draft is in the Outbox, so Outlook sends it now. This command changes an unsent draft only.',
  },
};

/** Reads the message and refuses it when it is not an unsent draft out of the Outbox. */
export const requireDraft = async (reads: GraphReads, messageId: GraphId): Promise<Result<undefined, GraphError>> => {
  const state = await readDraftState(reads, messageId);
  if (!state.ok) return state;
  return state.value === 'draft' ? ok(undefined) : err(REFUSALS[state.value]);
};

/** Whether the message is still an unsent draft out of the Outbox: the check before a write and in an undo. */
export const isDraftNow = async (reads: GraphReads, messageId: GraphId): Promise<Result<boolean, GraphError>> => {
  const state = await readDraftState(reads, messageId);
  return state.ok ? ok(state.value === 'draft') : state;
};

/** The attachments of the message, with no bytes. */
export const readDraftAttachments = async (reads: GraphReads, messageId: GraphId): Promise<Result<ReadonlyArray<DraftAttachment>, GraphError>> => {
  const read = await reads.get(`${attachmentsOf(messageId)}?$select=id,name,size`);
  if (!read.ok) return read;
  const page = listSchema.safeParse(read.value);
  return page.success ? ok(page.data.value) : err(unreadable('the attachments of a mail draft'));
};
