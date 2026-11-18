import type { CommandOptionMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';

/*
 * The words that the two attachment commands of a mail draft share: the
 * option of the draft and the shape of the answer.
 */

export const MESSAGE_ID_OF_DRAFT: CommandOptionMeta = {
  name: 'message-id',
  key: 'messageId',
  required: true,
  description:
    'The id of the unsent draft, as create-mail-draft, find-mail-drafts or list-mail-folder-messages --mail-folder-id drafts return it. A message that is not a draft, or a draft in the Outbox, is refused.',
  argumentHint: { kind: 'idOrName' },
};

export const DRAFT_ATTACHMENT_RESPONSE =
  '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the id of the draft and the name and the size of the file (and a `warning` when there is one), and `after` the ids of the draft and of the attachment. A draft that was sent, or an attachment that changed or is gone, between the plan and the write is in `skipped`; a write that Graph refused is in `failed`. With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.';
