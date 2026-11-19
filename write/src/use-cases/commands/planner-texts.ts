import type { CommandOptionMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';

/*
 * What the Planner commands share: the options that name an item, the option
 * of the consent to notify other people, and the words of their output.
 */

export const TASK_ID_OPTION: CommandOptionMeta = {
  name: 'task-id',
  key: 'taskId',
  required: true,
  description: 'The id of the Planner task, as list-planner-tasks, get-planner-task or create-planner-task return it.',
  argumentHint: { kind: 'idOrName' },
};

export const BUCKET_ID_OPTION: CommandOptionMeta = {
  name: 'bucket-id',
  key: 'bucketId',
  required: true,
  description: 'The id of the bucket, as list-plan-buckets or create-planner-bucket return it.',
  argumentHint: { kind: 'idOrName' },
};

export const PLAN_OPTION: CommandOptionMeta = {
  name: 'plan',
  key: 'plan',
  required: true,
  description: 'The plan: its name (no case) or its id, as list-planner-plans returns them.',
  argumentHint: { kind: 'idOrName' },
};

export const NOTIFY_OTHERS_OPTION: CommandOptionMeta = {
  name: 'notify-others',
  key: 'notifyOthers',
  required: false,
  description:
    '`true` permits a change of the assignment of a person other than you. Planner sends that person a notification, so without it such a change is refused. The dry-run shows each person.',
  argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
};

/** The scope of each Planner command. */
export const PLANNER_SCOPES = ['Tasks.ReadWrite'] as const;

/** The words of the output that each Planner command shares. */
export const PLANNER_RESPONSE =
  '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.';
