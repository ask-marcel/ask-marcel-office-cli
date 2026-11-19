import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { ANN, createPlannerStoreFake, ME } from '../../test-helpers/planner-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The task commands: create-planner-task, update-planner-task and
 * delete-planner-task, with the rule of the people: an assignment of another
 * person needs --notify-others true, in a run and in its undo.
 */

const ASSIGNED = { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' };
const UNDO_CHANGED = 'changed: The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.';

const posts = (calls: ReadonlyArray<string>): ReadonlyArray<string> => calls.filter((call) => !call.startsWith('GET'));

describe('create-planner-task', () => {
  it('adds the task with its fields, labels and the user, and the undo deletes it', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    const params = {
      plan: 'launch',
      bucket: 'doing',
      title: 'Book the room',
      due: '2026-11-25',
      status: 'inProgress',
      priority: 'important',
      labels: 'blue, category3',
      assign: 'me',
    };
    const run = await runWrite(planner.graph, 'create-planner-task', params, { journal });
    expect(run).toEqual({
      done: 1,
      skipped: [],
      failed: [],
      runId: RUN_ID,
      changes: [
        {
          action: 'create-planner-task',
          target: '/planner/tasks',
          before: { planId: 'plan-1', plan: 'Launch', bucketId: 'bucket-b', bucket: 'Doing', title: 'Book the room', assign: 'Robin Doe (you)' },
          after: { taskId: 'task-new-1', title: 'Book the room', planId: 'plan-1', bucketId: 'bucket-b' },
        },
      ],
    });
    expect(planner.items.get('/planner/tasks/task-new-1')).toMatchObject({
      planId: 'plan-1',
      bucketId: 'bucket-b',
      title: 'Book the room',
      dueDateTime: '2026-11-25T10:00:00Z',
      percentComplete: 50,
      priority: 3,
      appliedCategories: { category2: true, category3: true },
      assignments: { [ME.id]: ASSIGNED },
    });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(planner.items.has('/planner/tasks/task-new-1')).toBe(false);
  });

  it('refuses another person without --notify-others true, and shows each person with it', async () => {
    const planner = createPlannerStoreFake();
    const params = { plan: 'plan-1', bucket: 'bucket-a', title: 'Book the room', assign: 'me, ann@example.com' };
    expect(await runWrite(planner.graph, 'create-planner-task', params)).toMatchObject({
      error: {
        code: 'planner_notify_others_needed',
        message: '--assign or --unassign names 1 person(s) other than you. Planner sends each of them a notification, so this command needs --notify-others true for it.',
      },
    });
    const dry = await runWrite(planner.graph, 'create-planner-task', { ...params, notifyOthers: 'true' }, { dryRun: true });
    expect(dry).toMatchObject({ changes: [{ before: { assign: 'Robin Doe (you), Ann Lee' } }] });
    await runWrite(planner.graph, 'create-planner-task', { ...params, assign: ANN.id, notifyOthers: 'true' });
    expect(planner.items.get('/planner/tasks/task-new-1')).toMatchObject({ assignments: { [ANN.id]: ASSIGNED } });
  });

  it('refuses each bad field before any write', async () => {
    const planner = createPlannerStoreFake();
    const base = { plan: 'plan-1', bucket: 'bucket-a', title: 'Book the room' };
    const codeOf = async (params: Record<string, string>): Promise<unknown> =>
      ((await runWrite(planner.graph, 'create-planner-task', params)) as { error: { code?: string } }).error.code;
    expect(await codeOf({ ...base, status: 'completed', percent: '20' })).toBe('planner_status_and_percent');
    expect(await codeOf({ ...base, labels: 'Green' })).toBe('planner_label_not_found');
    expect(await codeOf({ ...base, assign: 'not a person' })).toBe('planner_person_invalid');
    expect(await codeOf({ ...base, assign: 'zed@example.com' })).toBe('planner_person_not_found');
    expect(await codeOf({ ...base, bucket: 'Nowhere' })).toBe('planner_bucket_not_found');
    expect(await runWrite(planner.graph, 'create-planner-task', { ...base, due: '2026-02-30' })).toMatchObject({ error: { message: expect.stringContaining('due') } });
    expect(posts(planner.calls)).toEqual([]);
  });

  it('keeps a new task when a person adds a checklist item to it before the undo', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'create-planner-task', { plan: 'plan-1', bucket: 'bucket-a', title: 'Book the room' }, { journal });
    planner.edit('/planner/tasks/task-new-1', { checklistItemCount: 1 });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual([UNDO_CHANGED]);
    expect(planner.items.has('/planner/tasks/task-new-1')).toBe(true);
  });
});

describe('update-planner-task', () => {
  it('changes the fields, the labels and the bucket with the ETag, and the undo writes them back', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    const run = await runWrite(planner.graph, 'update-planner-task', { taskId: 'task-1', title: 'Report', due: 'none', labels: 'Blue', bucket: 'Doing' }, { journal });
    expect(run).toMatchObject({
      done: 1,
      changes: [
        {
          action: 'update-planner-task',
          target: '/planner/tasks/task-1',
          before: { taskId: 'task-1', title: 'Write the report', planId: 'plan-1' },
          after: { fields: 'title, dueDateTime, appliedCategories, bucketId' },
        },
      ],
    });
    expect(planner.items.get('/planner/tasks/task-1')).toMatchObject({ title: 'Report', dueDateTime: null, appliedCategories: { category2: true }, bucketId: 'bucket-b' });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(planner.items.get('/planner/tasks/task-1')).toMatchObject({
      title: 'Write the report',
      dueDateTime: '2026-11-20T10:00:00Z',
      appliedCategories: { category1: true, category2: false },
      bucketId: 'bucket-a',
    });
  });

  it('removes the user with no consent, and refuses another person, the same person twice, or no change', async () => {
    const planner = createPlannerStoreFake();
    const codeOf = async (params: Record<string, string>): Promise<unknown> =>
      ((await runWrite(planner.graph, 'update-planner-task', params)) as { error: { code?: string } }).error.code;
    expect(await codeOf({ taskId: 'task-1', assign: 'ann@example.com' })).toBe('planner_notify_others_needed');
    expect(await codeOf({ taskId: 'task-1', assign: 'me', unassign: 'ME' })).toBe('planner_assign_and_unassign');
    expect(await codeOf({ taskId: 'task-1', unassign: 'ann@example.com', notifyOthers: 'true' })).toBe('planner_nothing_to_change');
    expect(await codeOf({ taskId: 'task-1', labels: 'Red' })).toBe('planner_nothing_to_change');
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'update-planner-task', { taskId: 'task-1', unassign: 'me' }, { journal });
    expect(planner.items.get('/planner/tasks/task-1')?.['assignments']).toEqual({});
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(planner.items.get('/planner/tasks/task-1')?.['assignments']).toEqual({ [ME.id]: ASSIGNED });
  });

  it('leaves the task when it changed after the run', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'update-planner-task', { taskId: 'task-1', percent: '30' }, { journal });
    planner.edit('/planner/tasks/task-1', { percentComplete: 40 });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual([UNDO_CHANGED]);
    expect(planner.items.get('/planner/tasks/task-1')).toMatchObject({ percentComplete: 40 });
  });
});

describe('delete-planner-task', () => {
  it('deletes the task, and the undo makes it again with its fields, details and the user', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    const run = await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1' }, { journal });
    expect(run).toMatchObject({
      done: 1,
      changes: [{ action: 'delete-planner-task', before: { taskId: 'task-1', title: 'Write the report', planId: 'plan-1', bucketId: 'bucket-a', undoAssigns: 'you' }, after: {} }],
    });
    expect(planner.items.has('/planner/tasks/task-1')).toBe(false);
    const undo = await undoWrite(planner.graph, journal, true);
    expect(undo.ok && undo.value.changes.map((change) => [change.status, change.now])).toEqual([['undone', { taskId: 'task-new-1' }]]);
    expect(planner.items.get('/planner/tasks/task-new-1')).toMatchObject({
      planId: 'plan-1',
      bucketId: 'bucket-a',
      title: 'Write the report',
      priority: 5,
      percentComplete: 0,
      dueDateTime: '2026-11-20T10:00:00Z',
      appliedCategories: { category1: true },
      assignments: { [ME.id]: ASSIGNED },
    });
    const details = planner.items.get('/planner/tasks/task-new-1/details') ?? {};
    expect(details).toMatchObject({ description: 'The text', previewType: 'automatic', references: { 'https%3A//example%2Ecom/spec': { alias: 'Spec', type: 'Other' } } });
    expect(
      Object.values(details['checklist'] as object).map((item: { title: string; isChecked: boolean; orderHint: string }) => [item.title, item.isChecked, item.orderHint])
    ).toEqual([
      ['Send', false, ' !'],
      ['Draft', true, ' ! !'],
    ]);
  });

  it('assigns another person again only when the delete had --notify-others true, and says so on the dry-run', async () => {
    const planner = createPlannerStoreFake();
    planner.edit('/planner/tasks/task-1', { assignments: { [ANN.id]: { orderHint: 'y' } } });
    const dry = await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1' }, { dryRun: true });
    expect(dry).toMatchObject({ changes: [{ before: { undoAssigns: 'you, but not the 1 other person(s): give --notify-others true to assign them again' } }] });
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1' }, { journal });
    await undoWrite(planner.graph, journal, true);
    expect(planner.items.get('/planner/tasks/task-new-1')?.['assignments']).toEqual({ [ME.id]: ASSIGNED });
    const consent = createRunJournalFake();
    planner.edit('/planner/tasks/task-new-1', { assignments: { [ANN.id]: { orderHint: 'y' } } });
    const kept = await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-new-1', notifyOthers: 'true' }, { journal: consent });
    expect(kept).toMatchObject({ changes: [{ before: { undoAssigns: 'you, and the 1 other person(s) (Planner notifies them)' } }] });
    await undoWrite(planner.graph, consent, true);
    expect(planner.items.get('/planner/tasks/task-new-2')?.['assignments']).toEqual({ [ME.id]: ASSIGNED, [ANN.id]: ASSIGNED });
  });

  it('passes over a task whose details changed after the plan', async () => {
    const planner = createPlannerStoreFake();
    planner.hooks.afterRead = (count) => {
      if (count === 4) planner.edit('/planner/tasks/task-1/details', { description: 'Other' });
    };
    const run = await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1' });
    expect(run).toMatchObject({ done: 0, skipped: [{ target: '/planner/tasks/task-1' }] });
    expect(planner.items.has('/planner/tasks/task-1')).toBe(true);
  });
});
