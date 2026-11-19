import { describe, expect, it } from 'bun:test';
import { meta as createPlannerTask } from './create-planner-task.ts';
import { meta as updatePlannerTask } from './update-planner-task.ts';
import { meta as deletePlannerTask } from './delete-planner-task.ts';

/*
 * The docs of the task commands of Planner, as the help, the docs pages and
 * the MCP tools show them: each option, the Graph path and body, the effect,
 * what the undo does and the scopes.
 */

describe('the docs of the Planner task commands', () => {
  it('pins the meta of create-planner-task', () => {
    expect(createPlannerTask).toStrictEqual({
      summary:
        'Add a task to a bucket of a Planner plan, with its dates, progress, priority, labels and people. An assignment to another person needs --notify-others true (Planner notifies that person). The undo deletes the task if nobody changed it.',
      category: 'tasks',
      graphMethod: 'POST',
      graphPathTemplate: '/planner/tasks',
      graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/planner-post-tasks',
      options: [
        {
          name: 'plan',
          key: 'plan',
          required: true,
          description: 'The plan: its name (no case) or its id, as list-planner-plans returns them.',
          argumentHint: {
            kind: 'idOrName',
          },
        },
        {
          name: 'bucket',
          key: 'bucket',
          required: true,
          description: 'The bucket of the task: its name (no case) or its id, in the plan.',
          argumentHint: {
            kind: 'idOrName',
          },
        },
        {
          name: 'title',
          key: 'title',
          required: true,
          description: 'The title of the task, at most 255 characters.',
        },
        {
          name: 'due',
          key: 'due',
          required: false,
          description: 'The due day, YYYY-MM-DD (kept at 10:00 UTC, as the Planner app does). `none` removes it.',
          argumentHint: {
            kind: 'iso8601',
          },
        },
        {
          name: 'start',
          key: 'start',
          required: false,
          description: 'The start day, YYYY-MM-DD. `none` removes it.',
          argumentHint: {
            kind: 'iso8601',
          },
        },
        {
          name: 'status',
          key: 'status',
          required: false,
          description: 'The progress: notStarted (0 %), inProgress (50 %) or completed (100 %). Not with --percent.',
          argumentHint: {
            kind: 'magicValue',
            values: ['notStarted', 'inProgress', 'completed'],
          },
        },
        {
          name: 'percent',
          key: 'percent',
          required: false,
          description: 'The progress in percent, a whole number from 0 to 100. Not with --status.',
        },
        {
          name: 'priority',
          key: 'priority',
          required: false,
          description: 'The priority: urgent, important, medium, low, or a number from 0 to 10.',
        },
        {
          name: 'labels',
          key: 'labels',
          required: false,
          description:
            'The labels, separated by commas: category1 to category25, or the label names of the plan (no case). They replace the labels of before. An empty value removes all labels.',
        },
        {
          name: 'assign',
          key: 'assign',
          required: false,
          description: 'The people of the task, separated by commas: `me`, a mail address or a user id. Another person than you needs --notify-others true.',
        },
        {
          name: 'notify-others',
          key: 'notifyOthers',
          required: false,
          description:
            '`true` permits a change of the assignment of a person other than you. Planner sends that person a notification, so without it such a change is refused. The dry-run shows each person.',
          argumentHint: {
            kind: 'magicValue',
            values: ['true', 'false'],
          },
        },
      ],
      example: '{write-bin} create-planner-task --plan "Launch" --bucket "To do" --title "Book the room" --due 2026-11-25 --priority important --assign me --dry-run',
      bodyTemplate:
        "{ planId: <id of {plan}>, bucketId: <id of {bucket}>, title: '{title}', dueDateTime?: '{due}T10:00:00Z', startDateTime?: '{start}T10:00:00Z', percentComplete?: <{status} or {percent}>, priority?: <{priority}>, appliedCategories?: { categoryN: true } (from {labels}), assignments?: { <user id of {assign}>: { orderHint: ' !' } } }. Another person than you needs {notify-others} true",
      effect: 'task-write',
      undo: 'deletes the task, only if the fields that the run set did not change and nobody added a checklist item, a reference or a description to it',
      scopesRequired: ['Tasks.ReadWrite', 'User.ReadBasic.All'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });

  it('pins the meta of update-planner-task', () => {
    expect(updatePlannerTask).toStrictEqual({
      summary:
        'Change a Planner task: its title, days, progress, priority, labels, bucket (in the same plan) and people. A change of the assignment of another person needs --notify-others true (Planner notifies that person). The undo writes the fields of before back.',
      category: 'tasks',
      graphMethod: 'PATCH',
      graphPathTemplate: '/planner/tasks/{task-id}',
      graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertask-update',
      options: [
        {
          name: 'task-id',
          key: 'taskId',
          required: true,
          description: 'The id of the Planner task, as list-planner-tasks, get-planner-task or create-planner-task return it.',
          argumentHint: {
            kind: 'idOrName',
          },
        },
        {
          name: 'title',
          key: 'title',
          required: false,
          description: 'The title of the task, at most 255 characters.',
        },
        {
          name: 'due',
          key: 'due',
          required: false,
          description: 'The due day, YYYY-MM-DD (kept at 10:00 UTC, as the Planner app does). `none` removes it.',
          argumentHint: {
            kind: 'iso8601',
          },
        },
        {
          name: 'start',
          key: 'start',
          required: false,
          description: 'The start day, YYYY-MM-DD. `none` removes it.',
          argumentHint: {
            kind: 'iso8601',
          },
        },
        {
          name: 'status',
          key: 'status',
          required: false,
          description: 'The progress: notStarted (0 %), inProgress (50 %) or completed (100 %). Not with --percent.',
          argumentHint: {
            kind: 'magicValue',
            values: ['notStarted', 'inProgress', 'completed'],
          },
        },
        {
          name: 'percent',
          key: 'percent',
          required: false,
          description: 'The progress in percent, a whole number from 0 to 100. Not with --status.',
        },
        {
          name: 'priority',
          key: 'priority',
          required: false,
          description: 'The priority: urgent, important, medium, low, or a number from 0 to 10.',
        },
        {
          name: 'labels',
          key: 'labels',
          required: false,
          description:
            'The labels, separated by commas: category1 to category25, or the label names of the plan (no case). They replace the labels of before. An empty value removes all labels.',
        },
        {
          name: 'bucket',
          key: 'bucket',
          required: false,
          description: 'Move the task to this bucket of the same plan: its name (no case) or its id.',
          argumentHint: {
            kind: 'idOrName',
          },
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
        {
          name: 'notify-others',
          key: 'notifyOthers',
          required: false,
          description:
            '`true` permits a change of the assignment of a person other than you. Planner sends that person a notification, so without it such a change is refused. The dry-run shows each person.',
          argumentHint: {
            kind: 'magicValue',
            values: ['true', 'false'],
          },
        },
      ],
      example: '{write-bin} update-planner-task --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --status completed --labels "Red" --dry-run',
      bodyTemplate:
        "{ title?: '{title}', dueDateTime?: '{due}T10:00:00Z', startDateTime?: '{start}T10:00:00Z', percentComplete?: <{status} or {percent}>, priority?: <{priority}>, bucketId?: <id of {bucket}>, appliedCategories?: { categoryN: true | false } (from {labels}), assignments?: { <user id of {assign}>: { orderHint: ' !' }, <user id of {unassign}>: null } } with If-Match: <ETag of the task>. Another person than you needs {notify-others} true",
      effect: 'task-write',
      undo: 'writes the fields, the labels, the bucket and the people of before back, only if the values that the run set did not change',
      scopesRequired: ['Tasks.ReadWrite', 'User.ReadBasic.All'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });

  it('pins the meta of delete-planner-task', () => {
    expect(deletePlannerTask).toStrictEqual({
      summary:
        'Delete a Planner task. The run journal keeps a copy, so the undo makes the task again with a new id. The undo assigns another person again only with --notify-others true on the delete. Over MCP it runs on run-destructive-command.',
      category: 'tasks',
      graphMethod: 'DELETE',
      graphPathTemplate: '/planner/tasks/{task-id}',
      graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannertask-delete',
      options: [
        {
          name: 'task-id',
          key: 'taskId',
          required: true,
          description: 'The id of the Planner task, as list-planner-tasks, get-planner-task or create-planner-task return it.',
          argumentHint: {
            kind: 'idOrName',
          },
        },
        {
          name: 'notify-others',
          key: 'notifyOthers',
          required: false,
          description:
            '`true` lets the undo assign the other people of the task again (Planner notifies them). Without it, the undo assigns only you. The dry-run says who the undo assigns.',
          argumentHint: {
            kind: 'magicValue',
            values: ['true', 'false'],
          },
        },
      ],
      example: '{write-bin} delete-planner-task --task-id "01gzSlKkIUSUl6DF_EilrmQAKDhh" --dry-run',
      bodyTemplate: 'no body, with If-Match: <ETag of the task>. The run journal keeps {notify-others} for the undo',
      effect: 'delete',
      undo: 'makes the task again from the copy in the run journal, with a new id: its fields, labels, bucket, description, preview, checklist and references, and your assignment (another person only with --notify-others true on the delete). Lost: the comments, the dates of creation and completion, the place in the bucket and the HTML formatting of the notes',
      scopesRequired: ['Tasks.ReadWrite'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });
});
