import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createLocalBytesFake, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { createPlannerStoreFake } from '../../test-helpers/planner-store-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The references of a task: add-planner-task-link and remove-planner-task-link
 * (the key is the URL in the encoded form of Graph), add-planner-task-file (a
 * file in the OneDrive of the user and a reference to it), and the names of a
 * plan: update-planner-plan.
 */

const DETAILS = '/planner/tasks/task-1/details';
const SPEC = 'https%3A//example%2Ecom/spec';
const UNDO_CHANGED = 'changed: The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.';

const referencesOf = (planner: ReturnType<typeof createPlannerStoreFake>): unknown => planner.items.get(DETAILS)?.['references'];

describe('the link commands', () => {
  it('adds a link under its encoded key, and the undo removes it', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    const run = await runWrite(planner.graph, 'add-planner-task-link', { taskId: 'task-1', url: 'https://example.com/a.b#c', alias: 'Budget' }, { journal });
    expect(run).toMatchObject({ done: 1, changes: [{ action: 'add-planner-task-link', target: DETAILS, before: { taskId: 'task-1', url: 'https://example.com/a.b#c' } }] });
    expect(referencesOf(planner)).toMatchObject({ 'https%3A//example%2Ecom/a%2Eb%23c': { alias: 'Budget', type: 'Other' } });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(Object.keys(referencesOf(planner) as object)).toEqual([SPEC]);
  });

  it('refuses a link that the task has, and a URL that is not http or https before any read', async () => {
    const planner = createPlannerStoreFake();
    expect(await runWrite(planner.graph, 'add-planner-task-link', { taskId: 'task-1', url: 'https://example.com/spec' })).toMatchObject({
      error: { code: 'planner_link_exists', message: '--url: the task has a link to this URL.' },
    });
    const before = planner.calls.length;
    expect(await runWrite(planner.graph, 'add-planner-task-link', { taskId: 'task-1', url: 'file:///c:/x.txt' })).toMatchObject({ error: { code: 'invalid_reference_url' } });
    expect(planner.calls.length).toBe(before);
  });

  it('removes a link, and the undo adds it again with its name and its type', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'remove-planner-task-link', { taskId: 'task-1', url: 'https://example.com/spec' }, { journal });
    expect(referencesOf(planner)).toEqual({});
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(referencesOf(planner)).toMatchObject({ [SPEC]: { alias: 'Spec', type: 'Other' } });
    expect(await runWrite(planner.graph, 'remove-planner-task-link', { taskId: 'task-1', url: 'https://example.com/other' })).toMatchObject({
      error: { code: 'planner_link_not_found' },
    });
  });
});

describe('add-planner-task-file', () => {
  const files = createLocalBytesFake({ 'C:/work/budget.xlsx': utf8Bytes('figures'), 'C:/work/a?.txt': utf8Bytes('x') });

  it('puts the file in the OneDrive folder, links it to the task, and the undo removes both', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    const run = await runWrite(planner.graph, 'add-planner-task-file', { taskId: 'task-1', file: 'C:/work/budget.xlsx' }, { journal, files });
    const path = '/me/drive/root:/Planner%20attachments/budget.xlsx';
    expect(run).toMatchObject({
      done: 1,
      changes: [
        { action: 'upload-planner-task-file', target: path, after: { itemId: 'drive-1' } },
        { action: 'add-planner-task-file', target: DETAILS },
      ],
    });
    expect(planner.calls.filter((call) => !call.startsWith('GET'))).toEqual([`PUT ${path}`, `PATCH ${DETAILS}`]);
    expect(referencesOf(planner)).toMatchObject({
      'https%3A//contoso-my%2Esharepoint%2Ecom/personal/robin/Documents/Planner%2520attachments/budget%2Exlsx': { alias: 'budget.xlsx', type: 'Excel' },
    });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone', 'undone']);
    expect(Object.keys(referencesOf(planner) as object)).toEqual([SPEC]);
    expect(planner.drive.size).toBe(0);
  });

  it('keeps the file in OneDrive when its content changed after the run', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'add-planner-task-file', { taskId: 'task-1', file: 'C:/work/budget.xlsx', alias: 'Figures' }, { journal, files });
    const item = planner.drive.get('/me/drive/items/drive-1') ?? {};
    planner.drive.set('/me/drive/items/drive-1', { ...item, cTag: '"c:{drive-1},2"' });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual([UNDO_CHANGED, 'undone']);
    expect(planner.drive.size).toBe(1);
  });

  it('refuses a group plan, a name that OneDrive cannot keep, and a name that the folder has', async () => {
    const planner = createPlannerStoreFake();
    planner.items.set('/planner/tasks/task-g', { id: 'task-g', planId: 'plan-g', title: 'Team task', '@odata.etag': 'W/"g"' });
    const codeOf = async (params: Record<string, string>): Promise<unknown> =>
      ((await runWrite(planner.graph, 'add-planner-task-file', params, { files })) as { error: { code?: string } }).error.code;
    expect(await codeOf({ taskId: 'task-g', file: 'C:/work/budget.xlsx' })).toBe('planner_group_plan_file');
    expect(await codeOf({ taskId: 'task-1', file: 'C:/work/a?.txt' })).toBe('planner_file_name_invalid');
    planner.drive.set('/me/drive/items/old', { id: 'old', path: '/me/drive/root:/Planner%20attachments/budget.xlsx' });
    expect(await codeOf({ taskId: 'task-1', file: 'C:/work/budget.xlsx' })).toBe('planner_file_exists');
    expect(planner.calls.filter((call) => !call.startsWith('GET'))).toEqual([]);
  });
});

describe('update-planner-plan', () => {
  it('renames the plan and its labels in two changes, and the undo writes both back', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    const run = await runWrite(
      planner.graph,
      'update-planner-plan',
      { plan: 'Launch', name: 'Launch 2027', labelNames: 'category1=Urgent, category3=New,category2=' },
      { journal }
    );
    expect(run).toMatchObject({
      done: 2,
      changes: [
        { action: 'rename-planner-plan', target: '/planner/plans/plan-1', before: { planId: 'plan-1', plan: 'Launch' } },
        { action: 'rename-planner-labels', target: '/planner/plans/plan-1/details' },
      ],
    });
    expect(planner.items.get('/planner/plans/plan-1')).toMatchObject({ title: 'Launch 2027' });
    expect(planner.items.get('/planner/plans/plan-1/details')?.['categoryDescriptions']).toEqual({ category1: 'Urgent', category3: 'New' });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone', 'undone']);
    expect(planner.items.get('/planner/plans/plan-1')).toMatchObject({ title: 'Launch' });
    expect(planner.items.get('/planner/plans/plan-1/details')?.['categoryDescriptions']).toEqual({ category1: 'Red', category2: 'Blue' });
  });

  it('refuses label names that are not categoryN=name, a label twice, and no change', async () => {
    const planner = createPlannerStoreFake();
    const message = '--label-names: item 2 is not "categoryN=name" with N from 1 to 25, or names a label twice.';
    for (const labelNames of ['category1=A,Red', 'category1=A,category26=B', 'category1=A,category1=B', 'category1=A,=B']) {
      expect(await runWrite(planner.graph, 'update-planner-plan', { plan: 'plan-1', labelNames })).toMatchObject({ error: { code: 'planner_label_names_invalid', message } });
    }
    expect(await runWrite(planner.graph, 'update-planner-plan', { plan: 'plan-1', labelNames: ' , ' })).toMatchObject({
      error: { code: 'planner_label_names_invalid', message: '--label-names: give at least one "categoryN=name".' },
    });
    expect(await runWrite(planner.graph, 'update-planner-plan', { plan: 'plan-1' })).toMatchObject({ error: { code: 'planner_nothing_to_change' } });
    expect(planner.calls).toEqual([]);
  });
});
