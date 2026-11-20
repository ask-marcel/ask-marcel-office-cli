import { describe, expect, it } from 'bun:test';
import { meta as createPlannerBucket } from './create-planner-bucket.ts';
import { meta as updatePlannerBucket } from './update-planner-bucket.ts';
import { meta as deletePlannerBucket } from './delete-planner-bucket.ts';

/*
 * The docs of the bucket commands of Planner, as the help, the docs pages and the
 * MCP tools show them.
 */

describe('the docs of the Planner bucket commands', () => {
  it('pins the meta of create-planner-bucket', () => {
    expect(createPlannerBucket).toStrictEqual({
      summary: 'Add a bucket to a Planner plan. A plan that has a bucket of the same name is refused. The undo deletes the bucket if it has no task.',
      category: 'tasks',
      graphMethod: 'POST',
      graphPathTemplate: '/planner/buckets',
      graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/planner-post-buckets',
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
          name: 'name',
          key: 'name',
          required: true,
          description: 'The name of the new bucket, at most 255 characters.',
        },
      ],
      example: '{write-bin} create-planner-bucket --plan "Launch" --name "Waiting" --dry-run',
      bodyTemplate: "{ name: '{name}', planId: <the id of {plan}> }",
      effect: 'task-write',
      undo: 'deletes the bucket, only if its name did not change and it has no task',
      scopesRequired: ['Tasks.ReadWrite'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });

  it('pins the meta of update-planner-bucket', () => {
    expect(updatePlannerBucket).toStrictEqual({
      summary:
        'Rename a Planner bucket, or move it: --first true puts it first in the plan, --after puts it just after another bucket. The undo writes the name back and puts the bucket back between the buckets that were next to it.',
      category: 'tasks',
      graphMethod: 'PATCH',
      graphPathTemplate: '/planner/buckets/{bucket-id}',
      graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannerbucket-update',
      options: [
        {
          name: 'bucket-id',
          key: 'bucketId',
          required: true,
          description: 'The id of the bucket, as list-plan-buckets or create-planner-bucket return it.',
          argumentHint: {
            kind: 'idOrName',
          },
        },
        {
          name: 'name',
          key: 'name',
          required: false,
          description: 'The new name of the bucket, at most 255 characters.',
        },
        {
          name: 'after',
          key: 'after',
          required: false,
          description: 'Put the bucket just after this bucket of the same plan: its name (no case) or its id. Not with --first.',
        },
        {
          name: 'first',
          key: 'first',
          required: false,
          description: '`true` puts the bucket first in the plan. Not with --after.',
          argumentHint: {
            kind: 'magicValue',
            values: ['true', 'false'],
          },
        },
      ],
      example: '{write-bin} update-planner-bucket --bucket-id "hsOf2dhOJkqyYYZEtdzDe2QAIUCR" --name "Waiting" --after "To do" --dry-run',
      bodyTemplate:
        "{ name?: '{name}', orderHint?: '<hint of {after}> <hint of the bucket after it>!', or ' <hint of the first bucket>!' with {first} true } with If-Match: <ETag of the bucket>",
      effect: 'task-write',
      undo: 'writes the name back and puts the bucket back between the buckets that were next to it, only if the name and the order hint that the run set did not change',
      scopesRequired: ['Tasks.ReadWrite'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });

  it('pins the meta of delete-planner-bucket', () => {
    expect(deletePlannerBucket).toStrictEqual({
      summary:
        'Delete an empty Planner bucket. A bucket with tasks is refused before any write: delete or move its tasks first. The undo makes the bucket again with a new id. Over MCP it runs on run-destructive-command.',
      category: 'tasks',
      graphMethod: 'DELETE',
      graphPathTemplate: '/planner/buckets/{bucket-id}',
      graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/plannerbucket-delete',
      options: [
        {
          name: 'bucket-id',
          key: 'bucketId',
          required: true,
          description: 'The id of the bucket, as list-plan-buckets or create-planner-bucket return it.',
          argumentHint: {
            kind: 'idOrName',
          },
        },
      ],
      example: '{write-bin} delete-planner-bucket --bucket-id "hsOf2dhOJkqyYYZEtdzDe2QAIUCR" --dry-run',
      bodyTemplate: 'no body, with If-Match: <ETag of the bucket>',
      effect: 'delete',
      undo: 'makes the bucket again (new id) between the buckets that were next to it, if the plan is still there',
      scopesRequired: ['Tasks.ReadWrite'],
      responseShape:
        '`{ done, skipped: [{ target, before, reason }], failed: [{ target, before, reason }], runId, changes: [{ action, target, before, after }] }`. `before` has the ids and the name of the item before the change, and `after` the ids after it: a new item shows its new id there. An item that changed between the plan and the write is in `skipped`; a write that Planner refused is in `failed` (an item changed after the read gives the code `planner_item_changed`: run the command again). With --dry-run: `{ dryRun: true, changes, skipped }`, and nothing changes.',
    });
  });
});
