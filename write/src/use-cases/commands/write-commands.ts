import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { writeCommandOf } from '@ask-marcel/office-core/use-cases/commands/command-graph.ts';
import type { Command, WriteCommand } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { withUnknownParamRejection } from '@ask-marcel/office-core/use-cases/commands/reject-unknown-params.ts';
import type { UndoHandlers } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import * as convertCalendarEventAttachmentToPdf from './convert-calendar-event-attachment-to-pdf.ts';
import * as convertGroupPostAttachmentToPdf from './convert-group-post-attachment-to-pdf.ts';
import * as convertMailAttachmentToPdf from './convert-mail-attachment-to-pdf.ts';
import * as createForwardDraft from './create-forward-draft.ts';
import * as createMailDraft from './create-mail-draft.ts';
import * as createReplyDraft from './create-reply-draft.ts';
import * as deleteMailDraft from './delete-mail-draft.ts';
import { DRAFT_UNDO_HANDLERS } from './draft-runs.ts';
import { FIELD_UNDO_HANDLERS } from './mail-field-write.ts';
import { MOVE_UNDO_HANDLERS } from './mail-move-write.ts';
import * as markMailMessage from './mark-mail-message.ts';
import * as moveMailMessage from './move-mail-message.ts';
import * as moveMailThread from './move-mail-thread.ts';
import * as setMailCategories from './set-mail-categories.ts';
import * as setMailFlag from './set-mail-flag.ts';
import * as updateMailDraft from './update-mail-draft.ts';

/*
 * The registry of the write package `@ask-marcel/office-write` (package split,
 * phase 5): its commands. Each is wrapped once, here, so that an unknown
 * parameter is refused before it runs and the command runs on the write view of
 * the graph it gets (`writeCommandOf`). The CLI, the MCP server and a library
 * caller all pass through this one point.
 */
const modules: Readonly<Record<string, WriteCommand>> = {
  'create-mail-draft': createMailDraft,
  'create-reply-draft': createReplyDraft,
  'create-forward-draft': createForwardDraft,
  'update-mail-draft': updateMailDraft,
  'convert-mail-attachment-to-pdf': convertMailAttachmentToPdf,
  'convert-calendar-event-attachment-to-pdf': convertCalendarEventAttachmentToPdf,
  'convert-group-post-attachment-to-pdf': convertGroupPostAttachmentToPdf,
  'move-mail-message': moveMailMessage,
  'move-mail-thread': moveMailThread,
  'mark-mail-message': markMailMessage,
  'set-mail-flag': setMailFlag,
  'set-mail-categories': setMailCategories,
  'delete-mail-draft': deleteMailDraft,
};

const commands: Readonly<Record<string, Command<WriteGraph>>> = Object.fromEntries(
  Object.entries(modules).map(([name, command]) => [name, withUnknownParamRejection(writeCommandOf(command))])
);

/**
 * The undo handler of each kind of change that a write of this package
 * records (write-plan.ts). A new command with an undo adds its kinds here.
 */
const undoHandlers: UndoHandlers = { ...DRAFT_UNDO_HANDLERS, ...MOVE_UNDO_HANDLERS, ...FIELD_UNDO_HANDLERS };

export { commands, undoHandlers };
