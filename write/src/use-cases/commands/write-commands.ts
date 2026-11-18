import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { writeCommandOf } from '@ask-marcel/office-core/use-cases/commands/command-graph.ts';
import type { Command, WriteCommand } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { withUnknownParamRejection } from '@ask-marcel/office-core/use-cases/commands/reject-unknown-params.ts';
import type { UndoHandlers } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import * as addMailDraftAttachment from './add-mail-draft-attachment.ts';
import * as addTodoLinkedResource from './add-todo-linked-resource.ts';
import * as addTodoTaskAttachment from './add-todo-task-attachment.ts';
import * as addTodoTaskStep from './add-todo-task-step.ts';
import * as applyTodoOps from './apply-todo-ops.ts';
import * as completeTodoTask from './complete-todo-task.ts';
import * as createTodoList from './create-todo-list.ts';
import * as createTodoTask from './create-todo-task.ts';
import * as deleteTodoTask from './delete-todo-task.ts';
import * as deleteTodoTaskStep from './delete-todo-task-step.ts';
import * as moveTodoTask from './move-todo-task.ts';
import * as removeMailDraftAttachment from './remove-mail-draft-attachment.ts';
import * as removeTodoTaskAttachment from './remove-todo-task-attachment.ts';
import * as updateTodoList from './update-todo-list.ts';
import * as updateTodoTask from './update-todo-task.ts';
import * as updateTodoTaskStep from './update-todo-task-step.ts';
import * as convertCalendarEventAttachmentToPdf from './convert-calendar-event-attachment-to-pdf.ts';
import * as convertGroupPostAttachmentToPdf from './convert-group-post-attachment-to-pdf.ts';
import * as convertMailAttachmentToPdf from './convert-mail-attachment-to-pdf.ts';
import * as createForwardDraft from './create-forward-draft.ts';
import * as createMailDraft from './create-mail-draft.ts';
import * as createReplyDraft from './create-reply-draft.ts';
import * as deleteMailDraft from './delete-mail-draft.ts';
import { DRAFT_UNDO_HANDLERS } from './draft-runs.ts';
import { DRAFT_ATTACHMENT_UNDO_HANDLERS } from './mail-attachment-undo.ts';
import { FIELD_UNDO_HANDLERS } from './mail-field-write.ts';
import { MOVE_UNDO_HANDLERS } from './mail-move-write.ts';
import { ITEM_UNDO_HANDLERS } from './todo-item-undo.ts';
import { TASK_UNDO_HANDLERS } from './todo-task-undo.ts';
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
  'add-mail-draft-attachment': addMailDraftAttachment,
  'remove-mail-draft-attachment': removeMailDraftAttachment,
  'create-todo-task': createTodoTask,
  'update-todo-task': updateTodoTask,
  'complete-todo-task': completeTodoTask,
  'move-todo-task': moveTodoTask,
  'delete-todo-task': deleteTodoTask,
  'add-todo-linked-resource': addTodoLinkedResource,
  'apply-todo-ops': applyTodoOps,
  'add-todo-task-step': addTodoTaskStep,
  'update-todo-task-step': updateTodoTaskStep,
  'delete-todo-task-step': deleteTodoTaskStep,
  'add-todo-task-attachment': addTodoTaskAttachment,
  'remove-todo-task-attachment': removeTodoTaskAttachment,
  'create-todo-list': createTodoList,
  'update-todo-list': updateTodoList,
};

const commands: Readonly<Record<string, Command<WriteGraph>>> = Object.fromEntries(
  Object.entries(modules).map(([name, command]) => [name, withUnknownParamRejection(writeCommandOf(command))])
);

/**
 * The undo handler of each kind of change that a write of this package
 * records (write-plan.ts). A new command with an undo adds its kinds here.
 */
const undoHandlers: UndoHandlers = {
  ...DRAFT_UNDO_HANDLERS,
  ...DRAFT_ATTACHMENT_UNDO_HANDLERS,
  ...MOVE_UNDO_HANDLERS,
  ...FIELD_UNDO_HANDLERS,
  ...TASK_UNDO_HANDLERS,
  ...ITEM_UNDO_HANDLERS,
};

export { commands, undoHandlers };
