import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createLocalBytesFake, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { createPlannerStoreFake } from '../../test-helpers/planner-store-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The reads of the Planner commands: each answer that this CLI cannot read
 * stops the command with words that name what Graph gave, and a failed read
 * stops a dry-run too.
 */

const files = createLocalBytesFake({ 'a.txt': utf8Bytes('x') });
const shape = (what: string): string => `Graph gave ${what} in a shape that this CLI cannot read.`;

type Unreadable = readonly [string, Record<string, string>, string, unknown, string];

const UNREADABLE: ReadonlyArray<Unreadable> = [
  ['update-planner-task', { taskId: 'task-1', title: 'x' }, 'GET /planner/tasks/task-1', {}, 'a Planner task'],
  ['update-planner-task', { taskId: 'task-1', title: 'x' }, 'GET /planner/tasks/task-1', { '@odata.etag': 'W/"x"', planId: 'bad id' }, 'a Planner task'],
  ['update-planner-task', { taskId: 'task-1', title: 'x' }, 'GET /planner/tasks/task-1', { '@odata.etag': 'W/"x"' }, 'a Planner task'],
  ['update-planner-task-details', { taskId: 'task-1', description: 'x' }, 'GET /planner/tasks/task-1/details', { id: 'task-1' }, 'the details of a Planner task'],
  ['update-planner-bucket', { bucketId: 'bucket-a', name: 'x' }, 'GET /planner/buckets/bucket-a', { '@odata.etag': '' }, 'a Planner bucket'],
  ['update-planner-bucket', { bucketId: 'bucket-a', name: 'x' }, 'GET /planner/buckets/bucket-a', { '@odata.etag': 'W/"x"', name: 'x' }, 'a Planner bucket'],
  ['update-planner-bucket', { bucketId: 'bucket-a', name: 'x' }, 'GET /planner/buckets/bucket-a', { '@odata.etag': 'W/"x"', planId: 'plan-1', name: 5 }, 'a Planner bucket'],
  [
    'update-planner-bucket',
    { bucketId: 'bucket-a', name: 'x' },
    'GET /planner/plans/plan-1/buckets',
    { value: [{ id: 'bad id', name: 'x', orderHint: 'a' }] },
    'the buckets of a Planner plan',
  ],
  ['update-planner-bucket', { bucketId: 'bucket-a', name: 'x' }, 'GET /planner/plans/plan-1/buckets', { value: [{ id: 'b', name: 'x' }] }, 'the buckets of a Planner plan'],
  ['delete-planner-bucket', { bucketId: 'bucket-a' }, 'GET /planner/buckets/bucket-a/tasks', { value: 'x' }, 'the tasks of a Planner bucket'],
  ['create-planner-bucket', { plan: 'Launch', name: 'x' }, 'GET /me/planner/plans', { value: [{ id: 'plan-1' }] }, 'the Planner plans'],
  ['create-planner-bucket', { plan: 'plan-1', name: 'x' }, 'GET /planner/plans/plan-1', { id: 'plan-1' }, 'a Planner plan'],
  ['create-planner-task', { plan: 'plan-1', bucket: 'bucket-a', title: 'x', labels: 'Red' }, 'GET /planner/plans/plan-1/details', {}, 'the details of a Planner plan'],
  ['create-planner-task', { plan: 'plan-1', bucket: 'bucket-a', title: 'x', assign: 'me' }, 'GET /me', { displayName: 'x' }, 'a user'],
  ['create-planner-task', { plan: 'plan-1', bucket: 'bucket-a', title: 'x' }, 'POST /planner/tasks', { id: 'bad id' }, 'the Planner task it created'],
  ['create-planner-bucket', { plan: 'plan-1', name: 'x' }, 'POST /planner/buckets', {}, 'the Planner bucket it created'],
  ['add-planner-task-file', { taskId: 'task-1', file: 'a.txt' }, 'PUT /me/drive/root:/Planner%20attachments/a.txt', { id: 'drive-9' }, 'the file it put in OneDrive'],
  [
    'add-planner-task-file',
    { taskId: 'task-1', file: 'a.txt' },
    'PUT /me/drive/root:/Planner%20attachments/a.txt',
    { id: 'bad id', webUrl: 'https://x.example/a' },
    'the file it put in OneDrive',
  ],
];

describe('a Planner answer that this CLI cannot read', () => {
  it.each(UNREADABLE)('stops %s at %s', async (name, params, call, answer, what) => {
    const planner = createPlannerStoreFake();
    planner.answers.set(call, answer);
    const run = await runWrite(planner.graph, name, params, { files });
    expect(run['error'] ?? (run['failed'] as ReadonlyArray<{ reason: string }>)[0]?.reason).toEqual(expect.objectContaining({ message: shape(what) }));
  });

  it('marks the undo of a deleted task failed when Graph gives the task made again with no id', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1' }, { journal });
    planner.answers.set('POST /planner/tasks', {});
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual([`failed: ${shape('the Planner task it made again')}`]);
  });

  it('gives no new id to a bucket made again when Graph gives none', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'delete-planner-bucket', { bucketId: 'bucket-b' }, { journal });
    planner.answers.set('POST /planner/buckets', {});
    const undo = await undoWrite(planner.graph, journal, true);
    expect(undo.ok && undo.value.changes.map((change) => [change.status, change.now])).toEqual([['undone', undefined]]);
  });
});

const SCENARIOS: ReadonlyArray<readonly [string, Record<string, string>]> = [
  ['create-planner-bucket', { plan: 'Launch', name: 'Waiting' }],
  ['update-planner-bucket', { bucketId: 'bucket-b', after: 'To do' }],
  ['delete-planner-bucket', { bucketId: 'bucket-b' }],
  ['create-planner-task', { plan: 'plan-1', bucket: 'To do', title: 'Book', labels: 'Red', assign: 'me,ann@example.com', notifyOthers: 'true' }],
  ['update-planner-task', { taskId: 'task-1', labels: 'Blue', bucket: 'Doing', unassign: 'me' }],
  ['delete-planner-task', { taskId: 'task-1' }],
  ['update-planner-task-details', { taskId: 'task-1', preview: 'noPreview' }],
  ['add-planner-checklist-item', { taskId: 'task-1', title: 'Print' }],
  ['update-planner-checklist-item', { taskId: 'task-1', item: 'c2', checked: 'true' }],
  ['remove-planner-checklist-item', { taskId: 'task-1', item: 'c1' }],
  ['add-planner-task-file', { taskId: 'task-1', file: 'a.txt' }],
  ['add-planner-task-link', { taskId: 'task-1', url: 'https://example.com/a' }],
  ['remove-planner-task-link', { taskId: 'task-1', url: 'https://example.com/spec' }],
  ['update-planner-plan', { plan: 'Launch', name: 'Next', labelNames: 'category1=Urgent' }],
];

describe('a failed read of a dry-run', () => {
  it.each(SCENARIOS)('stops the dry-run of %s at each read', async (name, params) => {
    const clean = createPlannerStoreFake();
    await runWrite(clean.graph, name, params, { files, dryRun: true });
    for (const call of new Set(clean.calls)) {
      const planner = createPlannerStoreFake();
      planner.failures.set(call, { type: 'api_error', status: 503, message: 'Try again later.' });
      expect([call, await runWrite(planner.graph, name, params, { files, dryRun: true })]).toEqual([
        call,
        expect.objectContaining({ error: expect.objectContaining({ message: 'Try again later.' }) }),
      ]);
    }
  });
});
