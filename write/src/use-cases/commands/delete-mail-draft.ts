import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { graphId, parseGraphId } from '../../domain/graph-id.ts';
import { planMailBatch } from './mail-batch.ts';
import { createFolderReader } from './mail-folder-reader.ts';
import { moveMessage } from './mail-move-write.ts';
import { BATCH_RESPONSE, createExecute, parseParams } from './mail-state-command.ts';

/*
 * Moves one draft to Deleted Items: the one command that puts a message there.
 * It refuses a message that is not a draft. It moves the draft
 * (`POST /me/messages/{id}/move` to `deleteditems`) and does not DELETE it: the
 * answer of a move gives the new id of the draft, which its undo needs to move
 * it back to Drafts.
 */

const schema = z.object({ messageId: z.string().min(1) });

const notADraft = (id: string): GraphError => ({
  type: 'validation_error',
  code: 'not_a_draft',
  message: `The message ${id} is not a draft, so delete-mail-draft does not delete it. This CLI deletes only unsent drafts. Check the id with get-mail-message.`,
});

const DELETED_ITEMS = graphId('deleteditems');

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const id = parseGraphId(parsed.value.messageId, '--message-id');
  if (!id.ok) return id;
  const folders = createFolderReader(reads);
  const deleted = await folders.folder(DELETED_ITEMS);
  if (!deleted.ok) return deleted;
  const bin = deleted.value;
  return planMailBatch(reads, folders, [id.value], {
    action: 'delete-draft',
    select: ['isDraft'],
    target: (messageId) => `/me/messages/${messageId}/move`,
    decide: (message) => {
      if (message.fields['isDraft'] !== true) return { refuse: notADraft(message.id) };
      return message.parentFolderId === bin.id ? { skip: 'The draft is in Deleted Items already.' } : 'change';
    },
    shown: () => ({ isDraft: 'true' }),
    write: async (graph, message) => {
      const moved = await moveMessage(graph, message.id, bin.id);
      if (!moved.ok) return moved;
      return ok({
        after: { messageId: moved.value.id, folderId: bin.id, folder: bin.displayName },
        undo: { kind: 'move-draft-to-drafts', data: { messageId: moved.value.id, folderId: bin.id } },
      });
    },
  });
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Delete one unsent mail draft: move it to Deleted Items, where its undo or Outlook can restore it. Refuses a message that is not a draft. It is the one command of this CLI that puts a message in Deleted Items. Over MCP it runs on run-destructive-command.',
  category: 'mail',
  graphMethod: 'POST',
  graphPathTemplate: '/me/messages/{message-id}/move, to the well-known folder deleteditems',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-move',
  options: [
    {
      name: 'message-id',
      key: 'messageId',
      required: true,
      description: 'The id of the draft, as create-mail-draft or find-mail-drafts return it. A message that is not a draft is refused.',
      argumentHint: { kind: 'idOrName' },
    },
  ],
  example: '{write-bin} delete-mail-draft --message-id "AAMkAD..." --dry-run',
  bodyTemplate: "{ destinationId: 'deleteditems' }",
  effect: 'delete',
  undo: 'moves the draft back to Drafts (it gets a new id again), only if it is still a draft in Deleted Items',
  scopesRequired: ['Mail.ReadWrite'],
  responseShape: BATCH_RESPONSE,
};

export { execute, meta, plan, schema };
