import { describe, expect, it } from 'bun:test';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createLocalBytesFake, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { createPlannerStoreFake } from '../../test-helpers/planner-store-fake.ts';
import type { PlannerStoreFake } from '../../test-helpers/planner-store-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * Each Graph call of the Planner commands that fails, and each read answer
 * that this CLI cannot read, stops the command with an error and makes no
 * later write. Each failed call of an undo marks its change failed and
 * changes nothing after it. The sweep takes the calls of a clean run of each
 * command, then fails them one by one.
 */

const BUSY: GraphError = { type: 'api_error', status: 503, code: 'ServiceUnavailable', message: 'Try again later.' };

const files = createLocalBytesFake({ 'steps.txt': utf8Bytes('[x] Draft\nNew item'), 'C:/work/budget.xlsx': utf8Bytes('figures') });

type Scenario = readonly [string, Record<string, string>];

const SCENARIOS: ReadonlyArray<Scenario> = [
  ['create-planner-bucket', { plan: 'Launch', name: 'Waiting' }],
  ['update-planner-bucket', { bucketId: 'bucket-b', name: 'Now', first: 'true' }],
  ['update-planner-bucket', { bucketId: 'bucket-a', after: 'Doing' }],
  ['delete-planner-bucket', { bucketId: 'bucket-b' }],
  ['create-planner-task', { plan: 'plan-1', bucket: 'To do', title: 'Book', labels: 'Red', assign: 'me,ann@example.com', notifyOthers: 'true' }],
  ['update-planner-task', { taskId: 'task-1', title: 'New', labels: 'Blue', bucket: 'Doing', unassign: 'me' }],
  ['delete-planner-task', { taskId: 'task-1' }],
  ['update-planner-task-details', { taskId: 'task-1', notesHtml: '<p>x</p>', checklistFile: 'steps.txt' }],
  ['add-planner-checklist-item', { taskId: 'task-1', title: 'Print' }],
  ['update-planner-checklist-item', { taskId: 'task-1', item: 'c2', checked: 'true' }],
  ['remove-planner-checklist-item', { taskId: 'task-1', item: 'Draft' }],
  ['add-planner-task-file', { taskId: 'task-1', file: 'C:/work/budget.xlsx' }],
  ['add-planner-task-link', { taskId: 'task-1', url: 'https://example.com/a' }],
  ['remove-planner-task-link', { taskId: 'task-1', url: 'https://example.com/spec' }],
  ['update-planner-plan', { plan: 'plan-1', name: 'Next', labelNames: 'category1=Urgent' }],
];

const keyOf = (call: string): string => call.replace(' [beta]', '');
const writesOf = (planner: PlannerStoreFake): ReadonlyArray<string> => planner.calls.filter((call) => !call.startsWith('GET'));

const callsOfRun = async ([name, params]: Scenario): Promise<ReadonlyArray<string>> => {
  const planner = createPlannerStoreFake();
  await runWrite(planner.graph, name, params, { files });
  return [...new Set(planner.calls.map(keyOf))];
};

const FAILED: ReadonlyArray<readonly [string, Record<string, string>, string]> = (
  await Promise.all(SCENARIOS.map(async (scenario) => (await callsOfRun(scenario)).map((call) => [scenario[0], scenario[1], call] as const)))
).flat();

describe('a failed Graph call of a Planner command', () => {
  it.each(FAILED)('stops %s at %s with no later write', async (name, params, call) => {
    const planner = createPlannerStoreFake();
    planner.failures.set(call, BUSY);
    const run = await runWrite(planner.graph, name, params, { files });
    const failed = (run['failed'] as ReadonlyArray<{ reason: string }> | undefined) ?? [];
    expect(run['error'] === undefined ? failed.map((item) => item.reason) : [(run['error'] as { message: string }).message]).toEqual(['Try again later.']);
    // The rename of a plan and of its labels are two ops: the second still runs when the first fails.
    const last = name === 'update-planner-plan' && call === 'PATCH /planner/plans/plan-1' ? 'PATCH /planner/plans/plan-1/details' : call;
    expect(keyOf(planner.calls.at(-1) ?? '')).toBe(last);
  });
});

const READS: ReadonlyArray<readonly [string, Record<string, string>, string]> = FAILED.filter(([, , call]) => call.startsWith('GET') && !call.startsWith('GET /me/drive/root:'));

describe('a read answer that this CLI cannot read', () => {
  it.each(READS)('stops %s at %s before any write', async (name, params, call) => {
    const planner = createPlannerStoreFake();
    planner.answers.set(call, null);
    const run = await runWrite(planner.graph, name, params, { files });
    expect(run['error'] ?? run['skipped']).toBeDefined();
    expect(writesOf(planner)).toEqual([]);
  });
});

const undoCallsOf = async ([name, params]: Scenario): Promise<ReadonlyArray<string>> => {
  const planner = createPlannerStoreFake();
  const journal = createRunJournalFake();
  await runWrite(planner.graph, name, params, { journal, files });
  const before = planner.calls.length;
  await undoWrite(planner.graph, journal, true);
  return [...new Set(planner.calls.slice(before).map(keyOf))];
};

const UNDO_SCENARIOS = SCENARIOS.filter(([name, params]) => name !== 'delete-planner-bucket' || params['bucketId'] === 'bucket-b');

// The details of a task made again are written after the task: a failure there is in the answer of the undo (a test below).
const UNDO_FAILED: ReadonlyArray<readonly [string, Record<string, string>, string]> = (
  await Promise.all(UNDO_SCENARIOS.map(async (scenario) => (await undoCallsOf(scenario)).map((call) => [scenario[0], scenario[1], call] as const)))
)
  .flat()
  .filter(([, , call]) => !call.includes('task-new-'));

describe('a failed Graph call of the undo of a Planner command', () => {
  it.each(UNDO_FAILED)('marks the undo of %s failed at %s', async (name, params, call) => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, name, params, { journal, files });
    planner.failures.set(call, BUSY);
    const statuses = statusesOf(await undoWrite(planner.graph, journal, true)) as ReadonlyArray<string>;
    expect(statuses).toContain('failed: Try again later.');
  });
});

describe('a failed write of the details of a task that the undo makes again', () => {
  it('keeps the task made again, and says in the answer of the undo that its details are not written', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1' }, { journal });
    planner.failures.set('GET /planner/tasks/task-new-1/details', BUSY);
    const undo = await undoWrite(planner.graph, journal, true);
    expect(undo.ok && undo.value.changes.map((change) => [change.status, change.now])).toEqual([['undone', { taskId: 'task-new-1', details: 'not written: Try again later.' }]]);
    expect(planner.calls.filter((call) => call === 'GET /planner/tasks/task-new-1/details')).toHaveLength(3);
    expect(planner.items.has('/planner/tasks/task-new-1')).toBe(true);
  });

  it('says so too when the PATCH of the details fails', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1' }, { journal });
    planner.failures.set('PATCH /planner/tasks/task-new-1/details', BUSY);
    const undo = await undoWrite(planner.graph, journal, true);
    expect(undo.ok && undo.value.changes.map((change) => change.now)).toEqual([{ taskId: 'task-new-1', details: 'not written: Try again later.' }]);
  });
});
