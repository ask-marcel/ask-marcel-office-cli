import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { deleteTaskOp } from './planner-delete.ts';
import { readMe } from './planner-people.ts';
import { readPlannerItem } from './planner-reader.ts';
import { copyOf } from './planner-task-copy.ts';
import { readTask } from './planner-task-reader.ts';
import { NOTIFY_OTHERS_OPTION, PLANNER_RESPONSE, PLANNER_SCOPES, TASK_ID_OPTION } from './planner-texts.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * Deletes a task. The run journal keeps a copy of the task and its details,
 * so the undo makes it again with a new id. The undo assigns the user again;
 * it assigns another person again only when the delete had --notify-others
 * true, as Planner notifies that person. The dry-run says who the undo
 * assigns.
 */

const schema = z.object({ taskId: z.string().min(1), notifyOthers: z.enum(['true', 'false']).optional() });

/** The words of the dry-run on the people that the undo assigns again. */
export const undoPeopleNote = (assignedMe: boolean, others: number, notify: boolean): string => {
  const me = assignedMe ? 'you' : 'not you (you are not assigned)';
  if (others === 0) return me;
  return notify
    ? `${me}, and the ${others} other person(s) (Planner notifies them)`
    : `${me}, but not the ${others} other person(s): give --notify-others true to assign them again`;
};

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const task = await readTask(reads, parsed.value.taskId);
  if (!task.ok) return task;
  const { item, planId, title } = task.value;
  const details = await readPlannerItem(reads, 'details', item.id);
  if (!details.ok) return details;
  const me = await readMe(reads);
  if (!me.ok) return me;
  const notify = parsed.value.notifyOthers === 'true';
  const copy = copyOf(item, details.value, planId, me.value.id, notify);
  const before = {
    taskId: item.id,
    title,
    planId,
    ...(copy.task.bucketId === undefined ? {} : { bucketId: copy.task.bucketId }),
    undoAssigns: undoPeopleNote(copy.assignedMe, copy.others.length, notify),
  };
  return ok(todoPlan({ ops: [deleteTaskOp(item, details.value, copy, before)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Delete a Planner task. The run journal keeps a copy, so the undo makes the task again with a new id. The undo assigns another person again only with --notify-others true on the delete. Over MCP it runs on run-destructive-command.',
  category: 'tasks',
  graphMethod: 'DELETE',
  graphPathTemplate: '/planner/tasks/{task-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertask-delete',
  options: [
    TASK_ID_OPTION,
    {
      ...NOTIFY_OTHERS_OPTION,
      description:
        '`true` lets the undo assign the other people of the task again (Planner notifies them). Without it, the undo assigns only you. The dry-run says who the undo assigns.',
    },
  ],
  example: '{write-bin} delete-planner-task --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --dry-run',
  bodyTemplate: 'no body, with If-Match: <ETag of the task>. The run journal keeps {notify-others} for the undo',
  effect: 'delete',
  undo: 'makes the task again from the copy in the run journal, with a new id: its fields, labels, bucket, description, preview, checklist and references, and your assignment (another person only with --notify-others true on the delete). Lost: the comments, the dates of creation and completion, the place in the bucket and the HTML formatting of the notes',
  scopesRequired: [...PLANNER_SCOPES],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
