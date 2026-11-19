import { z } from 'zod';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { assignmentChangeOf } from './planner-assignments.ts';
import { createOp, NEW_TASK_COUNTERS } from './planner-create.ts';
import { resolveBucket, resolveLabels, resolvePlan } from './planner-resolve.ts';
import { TASK_FIELD_OPTIONS, taskFieldShape, taskFieldsOf } from './planner-task-options.ts';
import { NOTIFY_OTHERS_OPTION, PLAN_OPTION, PLANNER_RESPONSE, PLANNER_SCOPES } from './planner-texts.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * Adds a task to a bucket of a plan, as the add op of the user's script. An
 * assignment to the user needs nothing; an assignment to another person needs
 * --notify-others true, as Planner sends that person a notification. The
 * dry-run shows each person. The undo deletes the task, only if the fields
 * that the run set did not change and nobody added a checklist item, a
 * reference or a description to it.
 */

const schema = z.object({
  plan: z.string().trim().min(1),
  bucket: z.string().trim().min(1),
  ...taskFieldShape,
  title: z.string().trim().min(1).max(255),
  labels: z.string().optional(),
  assign: z.string().optional(),
  notifyOthers: z.enum(['true', 'false']).optional(),
});

// The labels and the assignments of the POST, when the params give them.
const labelsAndPeople = (labels: ReadonlyArray<string>, people: Readonly<Record<string, unknown>>): Record<string, unknown> => ({
  ...(labels.length === 0 ? {} : { appliedCategories: Object.fromEntries(labels.map((key) => [key, true])) }),
  ...(Object.keys(people).length === 0 ? {} : { assignments: people }),
});

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const fields = taskFieldsOf(parsed.value);
  if (!fields.ok) return fields;
  const found = await resolvePlan(reads, parsed.value.plan);
  if (!found.ok) return found;
  const planId = found.value.id;
  const bucket = await resolveBucket(reads, planId, parsed.value.bucket, '--bucket');
  if (!bucket.ok) return bucket;
  const labels = await resolveLabels(reads, planId, parsed.value.labels ?? '');
  if (!labels.ok) return labels;
  const people = await assignmentChangeOf(reads, parsed.value, {});
  if (!people.ok) return people;
  const body = { planId, bucketId: bucket.value.id, ...fields.value, ...labelsAndPeople(labels.value, people.value.entries) };
  const before = {
    planId,
    plan: String(found.value.fields['title'] ?? ''),
    bucketId: bucket.value.id,
    bucket: bucket.value.name,
    title: parsed.value.title,
    ...people.value.shown,
  };
  const after = (id: string): Record<string, string> => ({ taskId: id, title: parsed.value.title, planId, bucketId: bucket.value.id });
  return ok(todoPlan({ ops: [createOp('task', body, { ...body, ...NEW_TASK_COUNTERS }, before, after)] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Add a task to a bucket of a Planner plan, with its dates, progress, priority, labels and people. An assignment to another person needs --notify-others true (Planner notifies that person). The undo deletes the task if nobody changed it.',
  category: 'tasks',
  graphMethod: 'POST',
  graphPathTemplate: '/planner/tasks',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/planner-post-tasks',
  options: [
    PLAN_OPTION,
    { name: 'bucket', key: 'bucket', required: true, description: 'The bucket of the task: its name (no case) or its id, in the plan.', argumentHint: { kind: 'idOrName' } },
    ...TASK_FIELD_OPTIONS.map((option) => (option.name === 'title' ? { ...option, required: true } : option)),
    {
      name: 'assign',
      key: 'assign',
      required: false,
      description: 'The people of the task, separated by commas: `me`, a mail address or a user id. Another person than you needs --notify-others true.',
    },
    NOTIFY_OTHERS_OPTION,
  ],
  example: '{write-bin} create-planner-task --plan "Launch" --bucket "To do" --title "Book the room" --due 2026-11-25 --priority important --assign me --dry-run',
  bodyTemplate:
    "{ planId: <id of {plan}>, bucketId: <id of {bucket}>, title: '{title}', dueDateTime?: '{due}T10:00:00Z', startDateTime?: '{start}T10:00:00Z', percentComplete?: <{status} or {percent}>, priority?: <{priority}>, appliedCategories?: { categoryN: true } (from {labels}), assignments?: { <user id of {assign}>: { orderHint: ' !' } } }. Another person than you needs {notify-others} true",
  effect: 'task-write',
  undo: 'deletes the task, only if the fields that the run set did not change and nobody added a checklist item, a reference or a description to it',
  scopesRequired: [...PLANNER_SCOPES, 'User.ReadBasic.All'],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
