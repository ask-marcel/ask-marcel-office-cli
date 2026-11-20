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
import * as createCalendarDraft from './create-calendar-draft.ts';
import * as deleteCalendarDraft from './delete-calendar-draft.ts';
import * as updateCalendarDraft from './update-calendar-draft.ts';
import { CALENDAR_UNDO_HANDLERS } from './calendar-draft-undo.ts';
import * as createPlannerBucket from './create-planner-bucket.ts';
import * as deletePlannerBucket from './delete-planner-bucket.ts';
import * as updatePlannerBucket from './update-planner-bucket.ts';
import * as createPlannerTask from './create-planner-task.ts';
import * as updatePlannerTask from './update-planner-task.ts';
import * as deletePlannerTask from './delete-planner-task.ts';
import * as updatePlannerTaskDetails from './update-planner-task-details.ts';
import * as addPlannerChecklistItem from './add-planner-checklist-item.ts';
import * as updatePlannerChecklistItem from './update-planner-checklist-item.ts';
import * as removePlannerChecklistItem from './remove-planner-checklist-item.ts';
import * as addPlannerTaskFile from './add-planner-task-file.ts';
import * as addPlannerTaskLink from './add-planner-task-link.ts';
import * as removePlannerTaskLink from './remove-planner-task-link.ts';
import * as updatePlannerPlan from './update-planner-plan.ts';
import { FILE_UNDO_HANDLERS } from './planner-file.ts';
import { PLANNER_UNDO_HANDLERS } from './planner-undo.ts';
import { TASK_COPY_UNDO_HANDLERS } from './planner-task-copy.ts';
import * as convertGroupPostAttachmentToPdf from './convert-group-post-attachment-to-pdf.ts';
import * as convertMailAttachmentToPdf from './convert-mail-attachment-to-pdf.ts';
import * as createForwardDraft from './create-forward-draft.ts';
import * as createMailDraft from './create-mail-draft.ts';
import * as createMailFolder from './create-mail-folder.ts';
import * as createReplyDraft from './create-reply-draft.ts';
import * as deleteMailDraft from './delete-mail-draft.ts';
import { DRAFT_UNDO_HANDLERS } from './draft-runs.ts';
import { DRAFT_ATTACHMENT_UNDO_HANDLERS } from './mail-attachment-undo.ts';
import { FIELD_UNDO_HANDLERS } from './mail-field-write.ts';
import { MAIL_FOLDER_UNDO_HANDLERS } from './mail-folder-create.ts';
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
  'create-mail-folder': createMailFolder,
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
  'create-calendar-draft': createCalendarDraft,
  'update-calendar-draft': updateCalendarDraft,
  'delete-calendar-draft': deleteCalendarDraft,
  'create-planner-bucket': createPlannerBucket,
  'update-planner-bucket': updatePlannerBucket,
  'delete-planner-bucket': deletePlannerBucket,
  'create-planner-task': createPlannerTask,
  'update-planner-task': updatePlannerTask,
  'delete-planner-task': deletePlannerTask,
  'update-planner-task-details': updatePlannerTaskDetails,
  'add-planner-checklist-item': addPlannerChecklistItem,
  'update-planner-checklist-item': updatePlannerChecklistItem,
  'remove-planner-checklist-item': removePlannerChecklistItem,
  'add-planner-task-file': addPlannerTaskFile,
  'add-planner-task-link': addPlannerTaskLink,
  'remove-planner-task-link': removePlannerTaskLink,
  'update-planner-plan': updatePlannerPlan,
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
  ...MAIL_FOLDER_UNDO_HANDLERS,
  ...TASK_UNDO_HANDLERS,
  ...ITEM_UNDO_HANDLERS,
  ...CALENDAR_UNDO_HANDLERS,
  ...PLANNER_UNDO_HANDLERS,
  ...TASK_COPY_UNDO_HANDLERS,
  ...FILE_UNDO_HANDLERS,
};

export { commands, undoHandlers };
