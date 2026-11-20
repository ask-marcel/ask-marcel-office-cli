import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createPlannerStoreFake, seedPlanner } from '../../test-helpers/planner-store-fake.ts';
import type { Item } from '../../test-helpers/planner-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The bucket commands: create-planner-bucket, update-planner-bucket (name and
 * place) and delete-planner-bucket, each with its dry-run and its undo, on a
 * fake Planner that refuses an old ETag as Planner does.
 */

const withDone = (): Array<[string, Item]> => [...seedPlanner(), ['/planner/buckets/bucket-c', { id: 'bucket-c', name: 'Done', planId: 'plan-1', orderHint: 'c' }]];

const UNDO_CHANGED = 'changed: The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.';

const CHANGED_SINCE_READ =
  'Planner refused the change: the item changed after this command read it (its ETag is not the same). Nothing changed. Run the command again to read the item as it is now.';

describe('create-planner-bucket', () => {
  it('adds the bucket to the plan of the name, and the undo deletes it with its ETag', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    const run = await runWrite(planner.graph, 'create-planner-bucket', { plan: 'LAUNCH', name: ' Waiting ' }, { journal });
    expect(run).toEqual({
      done: 1,
      skipped: [],
      failed: [],
      runId: RUN_ID,
      changes: [
        {
          action: 'create-planner-bucket',
          target: '/planner/buckets',
          before: { planId: 'plan-1', plan: 'Launch', name: 'Waiting' },
          after: { bucketId: 'bucket-new-1', name: 'Waiting', planId: 'plan-1' },
        },
      ],
    });
    expect(planner.items.get('/planner/buckets/bucket-new-1')).toMatchObject({ name: 'Waiting', planId: 'plan-1' });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(planner.items.has('/planner/buckets/bucket-new-1')).toBe(false);
  });

  it('shows the bucket on a dry-run and writes nothing', async () => {
    const planner = createPlannerStoreFake();
    const dry = await runWrite(planner.graph, 'create-planner-bucket', { plan: 'plan-1', name: 'Waiting' }, { dryRun: true });
    expect(dry).toEqual({
      dryRun: true,
      command: 'create-planner-bucket',
      changes: [{ action: 'create-planner-bucket', target: '/planner/buckets', before: { planId: 'plan-1', plan: 'Launch', name: 'Waiting' } }],
    });
    expect(planner.calls).toEqual(['GET /me/planner/plans', 'GET /planner/plans/plan-1', 'GET /planner/plans/plan-1/buckets']);
  });

  it('refuses a name that the plan has (no case)', async () => {
    const planner = createPlannerStoreFake();
    expect(await runWrite(planner.graph, 'create-planner-bucket', { plan: 'plan-1', name: 'TO DO' })).toMatchObject({
      error: { code: 'planner_bucket_exists', message: '--name: the plan has a bucket of this name. Give another name.' },
    });
  });

  it('keeps the bucket when the undo finds a task in it or a new name', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'create-planner-bucket', { plan: 'plan-1', name: 'Waiting' }, { journal });
    planner.items.set('/planner/tasks/task-x', { id: 'task-x', planId: 'plan-1', bucketId: 'bucket-new-1', title: 'New', '@odata.etag': 'W/"x"' });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual([UNDO_CHANGED]);
    planner.items.delete('/planner/tasks/task-x');
    planner.edit('/planner/buckets/bucket-new-1', { name: 'Renamed' });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual([UNDO_CHANGED]);
    expect(planner.items.has('/planner/buckets/bucket-new-1')).toBe(true);
  });
});

describe('the plan of a name or an id', () => {
  it('finds a plan by its id when the list of the plans of the user does not give it', async () => {
    const planner = createPlannerStoreFake();
    planner.answers.set('GET /me/planner/plans', { value: [] });
    const dry = await runWrite(planner.graph, 'create-planner-bucket', { plan: 'plan-1', name: 'Waiting' }, { dryRun: true });
    expect(dry).toMatchObject({ dryRun: true });
  });

  it('refuses a plan that no list and no read gives, and a name of two plans', async () => {
    const planner = createPlannerStoreFake();
    const notFound = {
      error: {
        code: 'planner_plan_not_found',
        message: '--plan: no Planner plan of yours has this name or id. list-planner-plans gives the name and the id of each plan.',
      },
    };
    expect(await runWrite(planner.graph, 'create-planner-bucket', { plan: 'plan-zz', name: 'Waiting' })).toMatchObject(notFound);
    expect(await runWrite(planner.graph, 'create-planner-bucket', { plan: 'No such plan', name: 'Waiting' })).toMatchObject(notFound);
    expect(planner.calls.filter((call) => call.includes('No such'))).toEqual([]);
    planner.items.set('/planner/plans/plan-2', { id: 'plan-2', title: 'launch', '@odata.etag': 'W/"p2"' });
    expect(await runWrite(planner.graph, 'create-planner-bucket', { plan: 'Launch', name: 'Waiting' })).toMatchObject({
      error: {
        code: 'planner_plan_ambiguous',
        message: '--plan: 2 plans have this name. Give the id of the plan, as list-planner-plans returns it.',
      },
    });
  });
});

describe('update-planner-bucket', () => {
  it('renames the bucket with its ETag, and the undo writes the name back', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    const run = await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-a', name: 'Backlog' }, { journal });
    expect(run).toMatchObject({
      done: 1,
      changes: [
        {
          action: 'update-planner-bucket',
          target: '/planner/buckets/bucket-a',
          before: { bucketId: 'bucket-a', name: 'To do', planId: 'plan-1' },
          after: { bucketId: 'bucket-a', planId: 'plan-1', fields: 'name' },
        },
      ],
    });
    expect(planner.items.get('/planner/buckets/bucket-a')).toMatchObject({ name: 'Backlog' });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(planner.items.get('/planner/buckets/bucket-a')).toMatchObject({ name: 'To do' });
  });

  it('moves the bucket first or after another one, and the undo puts it back between its old neighbours', async () => {
    const planner = createPlannerStoreFake(withDone());
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-c', first: 'true' }, { journal });
    expect(planner.items.get('/planner/buckets/bucket-c')).toMatchObject({ orderHint: ' a!' });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(planner.items.get('/planner/buckets/bucket-c')).toMatchObject({ orderHint: 'b !' });
    await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-a', after: 'doing' });
    expect(planner.items.get('/planner/buckets/bucket-a')).toMatchObject({ orderHint: 'b b !!' });
  });

  it('refuses two places, no change, and the bucket after itself', async () => {
    const planner = createPlannerStoreFake();
    const refusal = (code: string, message: string): object => ({ error: { code, message } });
    expect(await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-a', first: 'true', after: 'Doing' })).toMatchObject(
      refusal('planner_order_conflict', '--first true and --after both give the place of the bucket: give one of them.')
    );
    expect(await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-a', first: 'false' })).toMatchObject(
      refusal('planner_nothing_to_change', 'Give --name, --after or --first true: there is nothing to change.')
    );
    expect(await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-a', after: 'To Do' })).toMatchObject(
      refusal('planner_bucket_after_itself', '--after names the bucket itself. Name another bucket of the plan.')
    );
    expect(planner.calls.filter((call) => !call.startsWith('GET'))).toEqual([]);
  });

  it('passes over a bucket that changed after the plan, and reports a change made between the last read and the write', async () => {
    const planner = createPlannerStoreFake();
    planner.hooks.afterRead = (count) => {
      if (count === 2) planner.edit('/planner/buckets/bucket-a', { name: 'Other' });
    };
    const skipped = await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-a', name: 'Backlog' });
    expect(skipped).toMatchObject({ done: 0, skipped: [{ target: '/planner/buckets/bucket-a', reason: expect.stringContaining('changed after the plan') }] });
    const late = createPlannerStoreFake();
    late.hooks.afterRead = (count) => {
      if (count === 3) late.edit('/planner/buckets/bucket-a', { name: 'Other' });
    };
    expect(await runWrite(late.graph, 'update-planner-bucket', { bucketId: 'bucket-a', name: 'Backlog' })).toMatchObject({
      error: { code: 'planner_item_changed', message: CHANGED_SINCE_READ },
    });
    expect(late.items.get('/planner/buckets/bucket-a')).toMatchObject({ name: 'Other' });
  });
});

describe('delete-planner-bucket', () => {
  it('deletes an empty bucket, and the undo makes it again at its place with a new id', async () => {
    const planner = createPlannerStoreFake(withDone());
    const journal = createRunJournalFake();
    const run = await runWrite(planner.graph, 'delete-planner-bucket', { bucketId: 'bucket-b' }, { journal });
    expect(run).toMatchObject({
      done: 1,
      changes: [{ action: 'delete-planner-bucket', before: { bucketId: 'bucket-b', name: 'Doing', planId: 'plan-1' }, after: {} }],
    });
    expect(planner.calls).toContain('DELETE /planner/buckets/bucket-b');
    const undo = await undoWrite(planner.graph, journal, true);
    expect(undo.ok && undo.value.changes.map((change) => [change.status, change.now])).toEqual([['undone', { bucketId: 'bucket-new-1' }]]);
    expect(planner.items.get('/planner/buckets/bucket-new-1')).toMatchObject({ name: 'Doing', planId: 'plan-1', orderHint: 'a c!' });
  });

  it('refuses a bucket with tasks before any write, on a dry-run and on a run, and asks to delete or move them first', async () => {
    const planner = createPlannerStoreFake();
    const refusal = {
      error: {
        code: 'planner_bucket_has_tasks',
        source: 'validation',
        message:
          '--bucket-id: the bucket has 1 task(s) or more. Delete them first with delete-planner-task, or move them to another bucket with update-planner-task --bucket, then delete the bucket.',
      },
    };
    expect(await runWrite(planner.graph, 'delete-planner-bucket', { bucketId: 'bucket-a' }, { dryRun: true })).toMatchObject(refusal);
    expect(await runWrite(planner.graph, 'delete-planner-bucket', { bucketId: 'bucket-a' })).toMatchObject(refusal);
    expect(planner.calls.filter((call) => !call.startsWith('GET'))).toEqual([]);
    expect([planner.items.has('/planner/buckets/bucket-a'), planner.items.has('/planner/tasks/task-1')]).toEqual([true, true]);
  });

  it('passes over a bucket that got a task after the plan', async () => {
    const planner = createPlannerStoreFake();
    planner.hooks.afterRead = (count) => {
      if (count === 3) planner.items.set('/planner/tasks/task-x', { id: 'task-x', planId: 'plan-1', bucketId: 'bucket-b', title: 'New' });
    };
    const run = await runWrite(planner.graph, 'delete-planner-bucket', { bucketId: 'bucket-b' });
    expect(run).toMatchObject({ done: 0, skipped: [{ target: '/planner/buckets/bucket-b' }] });
    expect(planner.items.has('/planner/buckets/bucket-b')).toBe(true);
  });

  it('refuses an id that is not a Graph id before any request', async () => {
    const planner = createPlannerStoreFake();
    expect(await runWrite(planner.graph, 'delete-planner-bucket', { bucketId: '../plans/plan-1' })).toMatchObject({ error: { code: 'invalid_id' } });
    expect(planner.calls).toEqual([]);
  });
});
