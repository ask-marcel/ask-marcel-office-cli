import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { GraphReads, PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { labelPatch } from '../../domain/planner-fields.ts';
import { recordOf } from '../../domain/planner-snapshot.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { assignmentChangeOf } from './planner-assignments.ts';
import { patchOp } from './planner-patch.ts';
import { readTask } from './planner-task-reader.ts';
import { resolveBucket, resolveLabels } from './planner-resolve.ts';
import { TASK_FIELD_OPTIONS, taskFieldShape, taskFieldsOf } from './planner-task-options.ts';
import { NOTIFY_OTHERS_OPTION, PLANNER_RESPONSE, PLANNER_SCOPES, TASK_ID_OPTION } from './planner-texts.ts';
import { todoPlan } from './todo-batch.ts';
import { refused } from './planner-reader.ts';

/*
 * Changes the fields of a task, as the update op of the user's script: the
 * title, the days, the progress, the priority, the labels (they replace the
 * labels of before), the bucket (in the same plan) and the people. A change
 * of the assignment of another person needs --notify-others true: Planner
 * notifies that person, and the undo of a removal assigns the person again.
 * The undo writes the fields of before back.
 */

const schema = z.object({
  taskId: z.string().min(1),
  ...taskFieldShape,
  labels: z.string().optional(),
  bucket: z.string().trim().min(1).optional(),
  assign: z.string().optional(),
  unassign: z.string().optional(),
  notifyOthers: z.enum(['true', 'false']).optional(),
});

type Params = z.infer<typeof schema>;

const labelsOf = async (reads: GraphReads, planId: GraphId, current: unknown, raw: string | undefined): Promise<Result<Record<string, unknown>, GraphError>> => {
  if (raw === undefined) return ok({});
  const keys = await resolveLabels(reads, planId, raw);
  if (!keys.ok) return keys;
  const patch = labelPatch(recordOf(current) as Record<string, boolean>, keys.value);
  return ok(Object.keys(patch).length === 0 ? {} : { appliedCategories: patch });
};

const bucketOf = async (reads: GraphReads, planId: GraphId, raw: string | undefined): Promise<Result<Record<string, unknown>, GraphError>> => {
  if (raw === undefined) return ok({});
  const bucket = await resolveBucket(reads, planId, raw, '--bucket');
  return bucket.ok ? ok({ bucketId: bucket.value.id }) : bucket;
};

const nothing = (): Result<never, GraphError> =>
  refused('planner_nothing_to_change', 'The params change nothing of the task: give a field, --labels, --bucket, --assign or --unassign.');

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const p: Params = parsed.value;
  const fields = taskFieldsOf(p);
  if (!fields.ok) return fields;
  const task = await readTask(reads, p.taskId);
  if (!task.ok) return task;
  const { item, planId, title } = task.value;
  const labels = await labelsOf(reads, planId, item.fields['appliedCategories'], p.labels);
  if (!labels.ok) return labels;
  const bucket = await bucketOf(reads, planId, p.bucket);
  if (!bucket.ok) return bucket;
  const people = await assignmentChangeOf(reads, p, recordOf(item.fields['assignments']));
  if (!people.ok) return people;
  const assignments = Object.keys(people.value.entries).length === 0 ? {} : { assignments: people.value.entries };
  const body = { ...fields.value, ...labels.value, ...bucket.value, ...assignments };
  if (Object.keys(body).length === 0) return nothing();
  const before = { taskId: item.id, title, planId, ...people.value.shown };
  return ok(todoPlan({ ops: [patchOp({ action: 'update-planner-task', resource: 'task', item, body, before })] }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary:
    'Change a Planner task: its title, days, progress, priority, labels, bucket (in the same plan) and people. A change of the assignment of another person needs --notify-others true (Planner notifies that person). The undo writes the fields of before back.',
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/planner/tasks/{task-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertask-update',
  options: [
    TASK_ID_OPTION,
    ...TASK_FIELD_OPTIONS,
    {
      name: 'bucket',
      key: 'bucket',
      required: false,
      description: 'Move the task to this bucket of the same plan: its name (no case) or its id.',
      argumentHint: { kind: 'idOrName' },
    },
    {
      name: 'assign',
      key: 'assign',
      required: false,
      description: 'Add people to the task, separated by commas: `me`, a mail address or a user id. Another person than you needs --notify-others true.',
    },
    {
      name: 'unassign',
      key: 'unassign',
      required: false,
      description: 'Remove people from the task, in the form of --assign. Another person than you needs --notify-others true, as the undo assigns the person again.',
    },
    NOTIFY_OTHERS_OPTION,
  ],
  example: '{write-bin} update-planner-task --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --status completed --labels "Red" --dry-run',
  bodyTemplate:
    "{ title?: '{title}', dueDateTime?: '{due}T10:00:00Z', startDateTime?: '{start}T10:00:00Z', percentComplete?: <{status} or {percent}>, priority?: <{priority}>, bucketId?: <id of {bucket}>, appliedCategories?: { categoryN: true | false } (from {labels}), assignments?: { <user id of {assign}>: { orderHint: ' !' }, <user id of {unassign}>: null } } with If-Match: <ETag of the task>. Another person than you needs {notify-others} true",
  effect: 'task-write',
  undo: 'writes the fields, the labels, the bucket and the people of before back, only if the values that the run set did not change',
  scopesRequired: [...PLANNER_SCOPES, 'User.ReadBasic.All'],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
