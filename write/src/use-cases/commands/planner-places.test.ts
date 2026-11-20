import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createPlannerStoreFake, ME, seedPlanner } from '../../test-helpers/planner-store-fake.ts';
import type { Item } from '../../test-helpers/planner-store-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The places of the buckets (first, last, after another, in the order of their
 * hints), the passes over an item that changed after the plan, and what the
 * dry-run of a delete says.
 */

const bucket = (id: string, name: string, orderHint: string): [string, Item] => [`/planner/buckets/${id}`, { id, name, planId: 'plan-1', orderHint }];
const withDone = (): Array<[string, Item]> => [...seedPlanner(), bucket('bucket-c', 'Done', 'c')];
const hintOf = (planner: ReturnType<typeof createPlannerStoreFake>, id: string): unknown => planner.items.get(`/planner/buckets/${id}`)?.['orderHint'];

describe('the place of a bucket', () => {
  it('puts the first bucket first again after the other first one, and a bucket after the last one at the end', async () => {
    const planner = createPlannerStoreFake(withDone());
    await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-a', first: 'true' });
    expect(hintOf(planner, 'bucket-a')).toBe(' b!');
    await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-b', after: 'Done' });
    expect(hintOf(planner, 'bucket-b')).toBe('c !');
  });

  it('reads the buckets in the order of their hints, not of the list', async () => {
    const planner = createPlannerStoreFake([...seedPlanner(), bucket('bucket-0', 'Inbox', '0')]);
    await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-b', first: 'true' });
    expect(hintOf(planner, 'bucket-b')).toBe(' 0!');
  });

  it('puts the one bucket of a plan first with an empty hint around it', async () => {
    const planner = createPlannerStoreFake(seedPlanner().filter(([path]) => path !== '/planner/buckets/bucket-b'));
    await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-a', first: 'true' });
    expect(hintOf(planner, 'bucket-a')).toBe(' !');
  });

  it('keeps the place of a bucket when the undo writes back only its name', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-b', name: 'Now' }, { journal });
    await undoWrite(planner.graph, journal, true);
    expect(planner.items.get('/planner/buckets/bucket-b')).toMatchObject({ name: 'Doing', orderHint: 'b' });
  });

  it('makes a deleted first or last bucket again at the start or at the end', async () => {
    const first = createPlannerStoreFake();
    first.edit('/planner/tasks/task-1', { bucketId: 'bucket-b' });
    const journal = createRunJournalFake();
    await runWrite(first.graph, 'delete-planner-bucket', { bucketId: 'bucket-a' }, { journal });
    await undoWrite(first.graph, journal, true);
    expect(first.items.get('/planner/buckets/bucket-new-1')).toMatchObject({ name: 'To do', orderHint: ' b!' });
    const last = createPlannerStoreFake();
    const lastJournal = createRunJournalFake();
    await runWrite(last.graph, 'delete-planner-bucket', { bucketId: 'bucket-b' }, { journal: lastJournal });
    await undoWrite(last.graph, lastJournal, true);
    expect(last.items.get('/planner/buckets/bucket-new-1')).toMatchObject({ name: 'Doing', orderHint: 'a !' });
  });

  it('gives the ids and the name of an empty bucket on the dry-run of its delete', async () => {
    const planner = createPlannerStoreFake();
    const dry = await runWrite(planner.graph, 'delete-planner-bucket', { bucketId: 'bucket-b' }, { dryRun: true });
    expect((dry['changes'] as ReadonlyArray<{ before: unknown }>)[0]?.before).toEqual({ bucketId: 'bucket-b', name: 'Doing', planId: 'plan-1' });
  });
});

describe('an item that changed after the plan', () => {
  it('passes over a bucket renamed after the plan, with the same tasks', async () => {
    const planner = createPlannerStoreFake();
    planner.hooks.afterRead = (count) => (count === 3 ? planner.edit('/planner/buckets/bucket-b', { name: 'Other' }) : undefined);
    expect(await runWrite(planner.graph, 'delete-planner-bucket', { bucketId: 'bucket-b' })).toMatchObject({ done: 0, skipped: [{ target: '/planner/buckets/bucket-b' }] });
    expect(planner.items.has('/planner/buckets/bucket-b')).toBe(true);
  });

  it('passes over a task changed after the plan, with the same details', async () => {
    const planner = createPlannerStoreFake();
    planner.hooks.afterRead = (count) => (count === 3 ? planner.edit('/planner/tasks/task-1', { title: 'Other' }) : undefined);
    expect(await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1' })).toMatchObject({ done: 0, skipped: [{ target: '/planner/tasks/task-1' }] });
    expect(planner.calls.filter((call) => call.startsWith('DELETE'))).toEqual([]);
  });
});

describe('the dry-run of the delete of a task', () => {
  it('says that the undo does not assign you when you are not assigned, and gives no bucket of a task with none', async () => {
    const planner = createPlannerStoreFake();
    planner.items.set('/planner/tasks/task-1', { ...planner.items.get('/planner/tasks/task-1'), bucketId: undefined, assignments: { [ME.id]: null } });
    const dry = await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1' }, { dryRun: true });
    expect((dry['changes'] as ReadonlyArray<{ before: unknown }>)[0]?.before).toEqual({
      taskId: 'task-1',
      title: 'Write the report',
      planId: 'plan-1',
      undoAssigns: 'not you (you are not assigned)',
    });
  });

  it('marks the undo of a delete missing when the plan is gone', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1' }, { journal });
    planner.items.delete('/planner/plans/plan-1');
    expect(statusesOf(await undoWrite(planner.graph, journal, false))).toEqual(['missing: The item is no longer there, so there is nothing to undo.']);
  });
});
