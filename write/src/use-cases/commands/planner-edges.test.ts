import { describe, expect, it } from 'bun:test';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err } from '@ask-marcel/office-core/domain/result.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createLocalBytesFake, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { createPlannerStoreFake } from '../../test-helpers/planner-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';
import { PLANNER_UNDO_HANDLERS } from './planner-undo.ts';

/*
 * The edges of the Planner commands: an older plan of a group (an owner and no
 * container), a PATCH whose answer has no item, a bucket name of two buckets,
 * a task with no details, and an undo whose journal this CLI did not write.
 */

const BUSY: GraphError = { type: 'api_error', status: 503, code: 'ServiceUnavailable', message: 'Try again later.' };
const UNDO_CHANGED = 'changed: The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.';
const BAD_JOURNAL = 'failed: The run journal names an id that is not a Graph id, or data that this CLI did not write, so this undo does not run.';

describe('the edges of the Planner commands', () => {
  it('refuses a file for an older plan of a group, which has an owner and no container', async () => {
    const planner = createPlannerStoreFake();
    planner.items.set('/planner/plans/plan-1', { id: 'plan-1', title: 'Launch', owner: 'group-1', '@odata.etag': 'W/"o"' });
    const files = createLocalBytesFake({ 'a.txt': utf8Bytes('x') });
    expect(await runWrite(planner.graph, 'add-planner-task-file', { taskId: 'task-1', file: 'a.txt' }, { files })).toMatchObject({ error: { code: 'planner_group_plan_file' } });
  });

  it('reads the item when the answer of a PATCH has none, and keeps its values for the undo', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    planner.answers.set('PATCH /planner/buckets/bucket-a', undefined);
    await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-a', name: 'Backlog' }, { journal });
    planner.answers.delete('PATCH /planner/buckets/bucket-a');
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(planner.items.get('/planner/buckets/bucket-a')).toMatchObject({ name: 'To do' });
  });

  it('keeps no values when that read fails too, so the undo says changed and writes nothing', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    planner.answers.set('PATCH /planner/buckets/bucket-a', undefined);
    planner.hooks.afterRead = (count) => (count === 3 ? planner.failures.set('GET /planner/buckets/bucket-a', BUSY) : undefined);
    expect(await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-a', name: 'Backlog' }, { journal })).toMatchObject({ done: 1 });
    planner.failures.clear();
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual([UNDO_CHANGED]);
    expect(planner.items.get('/planner/buckets/bucket-a')).toMatchObject({ name: 'Backlog' });
  });

  it('gives planner_item_changed for a 409 of Planner too, which a PATCH with an old ETag got live', async () => {
    const planner = createPlannerStoreFake();
    planner.failures.set('PATCH /planner/buckets/bucket-a', { type: 'api_error', status: 409, message: 'The attempted changes conflicted with already accepted changes.' });
    expect(await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-a', name: 'Backlog' })).toMatchObject({
      error: {
        code: 'planner_item_changed',
        message:
          'Planner refused the change: the item changed after this command read it (its ETag is not the same). Nothing changed. Run the command again to read the item as it is now.',
      },
    });
  });

  it('puts new checklist items after the last item, in the order of the file', async () => {
    const planner = createPlannerStoreFake();
    const files = createLocalBytesFake({ 'steps.txt': utf8Bytes('One\n[x] Draft\nTwo\nSend') });
    await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', checklistFile: 'steps.txt' }, { files });
    const hints = Object.values(planner.items.get('/planner/tasks/task-1/details')?.['checklist'] as Record<string, { title: string; orderHint: string }>).map((item) => [
      item.title,
      item.orderHint,
    ]);
    expect(hints.toSorted((a, b) => String(a[0]).localeCompare(String(b[0])))).toEqual([
      ['Draft', '2'],
      ['One', '2 !'],
      ['Send', '1'],
      ['Two', '2 ! !'],
    ]);
    await runWrite(planner.graph, 'add-planner-checklist-item', { taskId: 'task-1', title: 'Three' });
    const added = Object.values(planner.items.get('/planner/tasks/task-1/details')?.['checklist'] as Record<string, { title: string; orderHint: string }>).find(
      (item) => item.title === 'Three'
    );
    expect(added?.orderHint).toBe('2 ! ! !');
  });

  it('gives the ids of the item and the fields of the change after an update, not the values of before', async () => {
    const planner = createPlannerStoreFake();
    const run = await runWrite(planner.graph, 'update-planner-task', { taskId: 'task-1', title: 'Report' });
    expect((run['changes'] as ReadonlyArray<{ after: unknown }>)[0]?.after).toEqual({ taskId: 'task-1', planId: 'plan-1', fields: 'title' });
  });

  it('refuses a bucket name of two buckets and asks for the id', async () => {
    const planner = createPlannerStoreFake();
    planner.edit('/planner/buckets/bucket-b', { name: 'TO DO' });
    expect(await runWrite(planner.graph, 'create-planner-task', { plan: 'plan-1', bucket: 'to do', title: 'x' })).toMatchObject({
      error: { code: 'planner_bucket_ambiguous', message: '--bucket: 2 buckets of the plan have this name. Give the id of the bucket, as list-plan-buckets returns it.' },
    });
  });

  it('makes a task with no details again with no write of details', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    planner.edit('/planner/tasks/task-1/details', { description: '', checklist: { c1: null, c2: null }, references: { 'https%3A//example%2Ecom/spec': null } });
    await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1' }, { journal });
    const undo = await undoWrite(planner.graph, journal, true);
    expect(undo.ok && undo.value.changes.map((change) => change.now)).toEqual([{ taskId: 'task-new-1' }]);
    expect(planner.calls.filter((call) => call.includes('task-new-1/details'))).toEqual([]);
  });

  it('makes the description of a task with no checklist and no reference again, and nothing else', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    planner.edit('/planner/tasks/task-1/details', { checklist: { c1: null, c2: null }, references: { 'https%3A//example%2Ecom/spec': null } });
    await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1' }, { journal });
    await undoWrite(planner.graph, journal, true);
    expect(planner.items.get('/planner/tasks/task-new-1/details')).toMatchObject({ description: 'The text', previewType: 'automatic', checklist: {}, references: {} });
  });

  it('refuses the undo of a bucket deleted with its tasks, in the check and in the write', async () => {
    const handler = PLANNER_UNDO_HANDLERS['planner-bucket-deleted-with-tasks'];
    const planner = createPlannerStoreFake();
    const refusal: Result<never, GraphError> = err({
      type: 'validation_error',
      code: 'planner_bucket_no_undo',
      message:
        'The bucket had 2 task(s) when the run deleted it. This CLI cannot make a bucket and its tasks again, so this change has no undo. Look for the tasks in Planner: they can be gone, or have no bucket.',
    });
    expect(await handler?.check(planner.graph, { bucketId: 'bucket-a', tasks: 2 })).toEqual(refusal);
    expect(await handler?.apply(planner.graph, { bucketId: 'bucket-a', tasks: 2 })).toEqual(refusal);
    expect(planner.calls).toEqual([]);
  });

  it('refuses a journal whose undo names a kind of item or an id that this CLI did not write', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'update-planner-bucket', { bucketId: 'bucket-a', name: 'Backlog' }, { journal });
    const record = journal.records.get(RUN_ID) as { changes: Array<{ undo: { data: Record<string, unknown> } }> };
    const data = record.changes[0]?.undo.data ?? {};
    data['resource'] = 'mailbox';
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual([BAD_JOURNAL]);
    data['resource'] = 'bucket';
    data['id'] = '../plans/plan-1';
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual([BAD_JOURNAL]);
    expect(planner.calls.filter((call) => call.startsWith('PATCH'))).toEqual(['PATCH /planner/buckets/bucket-a']);
  });
});
