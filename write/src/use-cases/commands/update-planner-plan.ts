import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { GraphReads, ItemIds, PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { isLabelKey } from '../../domain/planner-fields.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import { patchOp } from './planner-patch.ts';
import { readPlannerItem } from './planner-reader.ts';
import { resolvePlan } from './planner-resolve.ts';
import { PLAN_OPTION, PLANNER_RESPONSE, PLANNER_SCOPES } from './planner-texts.ts';
import type { TodoOp } from './todo-batch.ts';
import { todoPlan } from './todo-batch.ts';

/*
 * Renames a plan and the labels of a plan. No command creates a plan. The
 * name and the label names are two Planner items (the plan and its details),
 * so the run can have two changes, each with its ETag and its undo.
 */

const schema = z.object({ plan: z.string().trim().min(1), name: z.string().trim().min(1).max(255).optional(), labelNames: z.string().optional() });

const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

/** The label names of `--label-names`: `category1=Red,category3=` (an empty name removes it). The refusal names the item, not its text. */
export const labelNamesOf = (raw: string): Result<Record<string, string | null>, GraphError> => {
  const names: Record<string, string | null> = {};
  const items = raw.split(',').filter((item) => item.trim() !== '');
  for (const [index, item] of items.entries()) {
    const at = item.indexOf('=');
    const key = item.slice(0, Math.max(at, 0)).trim();
    if (at < 0 || !isLabelKey(key) || Object.hasOwn(names, key))
      return refused('planner_label_names_invalid', `--label-names: item ${index + 1} is not "categoryN=name" with N from 1 to 25, or names a label twice.`);
    const name = item.slice(at + 1).trim();
    names[key] = name === '' ? null : name;
  }
  return Object.keys(names).length === 0 ? refused('planner_label_names_invalid', '--label-names: give at least one "categoryN=name".') : ok(names);
};

// The op of the label names, read with the ETag of the details of the plan; none when --label-names is not given.
const labelOpsOf = async (
  reads: GraphReads,
  planId: GraphId,
  labels: Record<string, string | null> | undefined,
  before: ItemIds
): Promise<Result<ReadonlyArray<TodoOp>, GraphError>> => {
  if (labels === undefined) return ok([]);
  const details = await readPlannerItem(reads, 'planDetails', planId);
  if (!details.ok) return details;
  return ok([patchOp({ action: 'rename-planner-labels', resource: 'planDetails', item: details.value, body: { categoryDescriptions: labels }, before })]);
};

const plan: PlanWrite = async (reads, params) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const { name, labelNames } = parsed.value;
  if (name === undefined && labelNames === undefined) return refused('planner_nothing_to_change', 'Give --name or --label-names: there is nothing to change.');
  const labels = labelNames === undefined ? ok(undefined) : labelNamesOf(labelNames);
  if (!labels.ok) return labels;
  const found = await resolvePlan(reads, parsed.value.plan);
  if (!found.ok) return found;
  const before = { planId: found.value.id, plan: String(found.value.fields['title'] ?? '') };
  const ops: TodoOp[] = name === undefined ? [] : [patchOp({ action: 'rename-planner-plan', resource: 'plan', item: found.value, body: { title: name }, before })];
  const labelOps = await labelOpsOf(reads, found.value.id, labels.value, before);
  return labelOps.ok ? ok(todoPlan({ ops: [...ops, ...labelOps.value] })) : labelOps;
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: 'Rename a Planner plan (--name) and its labels (--label-names). No command creates a plan. The undo writes the name and the label names of before back.',
  category: 'tasks',
  graphMethod: 'PATCH',
  graphPathTemplate: '/planner/plans/<id of {plan}> and /planner/plans/<id of {plan}>/details',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannerplan-update',
  options: [
    PLAN_OPTION,
    { name: 'name', key: 'name', required: false, description: 'The new name of the plan, at most 255 characters.' },
    {
      name: 'label-names',
      key: 'labelNames',
      required: false,
      description:
        'The names of labels, separated by commas, in the form category1=Red,category5=Blocked (N from 1 to 25). An empty name removes the name: category2=. The other labels keep their names.',
    },
  ],
  example: '{write-bin} update-planner-plan --plan "Launch" --name "Launch 2027" --label-names "category1=Urgent,category2=" --dry-run',
  bodyTemplate: "{ title: '{name}' } on the plan, and { categoryDescriptions: { categoryN: '<name>' | null } } (from {label-names}) on its details, each with If-Match: <its ETag>",
  effect: 'task-write',
  undo: 'writes the name and the label names of before back, only if the values that the run set did not change',
  scopesRequired: [...PLANNER_SCOPES],
  responseShape: PLANNER_RESPONSE,
};

export { execute, meta, plan, schema };
