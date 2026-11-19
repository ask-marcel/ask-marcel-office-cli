import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { ANN, createPlannerStoreFake, ME } from '../../test-helpers/planner-store-fake.ts';
import type { Item } from '../../test-helpers/planner-store-fake.ts';
import { runWrite } from '../../test-helpers/run-write.ts';
import { undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The copy of a deleted task: what it keeps of fields, people and details that
 * Graph gives in odd shapes, and when the undo writes the details again.
 */

const TASK = '/planner/tasks/task-1';
const DETAILS = '/planner/tasks/task-1/details';

const deleteAndUndo = async (task: Item, details: Item): Promise<ReturnType<typeof createPlannerStoreFake>> => {
  const planner = createPlannerStoreFake();
  planner.items.set(TASK, { ...task, '@odata.etag': 'W/"t"' });
  planner.items.set(DETAILS, { ...details, '@odata.etag': 'W/"d"' });
  const journal = createRunJournalFake();
  await runWrite(planner.graph, 'delete-planner-task', { taskId: 'task-1' }, { journal });
  await undoWrite(planner.graph, journal, true);
  return planner;
};

const base: Item = { id: 'task-1', planId: 'plan-1', title: 'Write the report' };

describe('the copy of a deleted task', () => {
  it('keeps the fields that have a value, the assignments that are not removed, and the labels that are on', async () => {
    const task = {
      ...base,
      bucketId: 'bad id',
      dueDateTime: null,
      priority: 3,
      assignments: { [ME.id]: {}, [ANN.id]: null },
      appliedCategories: { category1: true, category2: false },
    };
    const planner = await deleteAndUndo(task, { id: 'task-1' });
    const made = planner.items.get('/planner/tasks/task-new-1') ?? {};
    expect(made).toMatchObject({ planId: 'plan-1', title: 'Write the report', priority: 3, appliedCategories: { category1: true }, assignments: { [ME.id]: {} } });
    expect(Object.keys(made).filter((key) => ['bucketId', 'dueDateTime', 'startDateTime', 'percentComplete'].includes(key))).toEqual([]);
    expect(Object.keys(made['assignments'] as object)).toEqual([ME.id]);
  });

  it('leaves out a checklist item and a reference that Graph gives in another shape', async () => {
    const details = {
      id: 'task-1',
      description: 'Text',
      checklist: { c1: { title: 'Kept', isChecked: true }, c2: null, c3: { title: 5 }, c4: 'x' },
      references: { 'https%3A//a%2Eb': { alias: 'A' }, 'https%3A//c%2Ed': null },
    };
    const planner = await deleteAndUndo(base, details);
    const made = planner.items.get('/planner/tasks/task-new-1/details') ?? {};
    expect(Object.values(made['checklist'] as object).map((item: { title: string }) => item.title)).toEqual(['Kept']);
    expect(Object.keys(made['references'] as object)).toEqual(['https%3A//a%2Eb']);
  });

  it('writes the details again when only one part of them has a value', async () => {
    const parts: ReadonlyArray<Item> = [
      { previewType: 'noPreview' },
      { checklist: { c1: { title: 'One', isChecked: false } } },
      { references: { 'https%3A//a%2Eb': { alias: 'A', type: 'Other' } } },
      { description: 'Text' },
    ];
    for (const details of parts) {
      const planner = await deleteAndUndo(base, { id: 'task-1', ...details });
      expect(planner.calls).toContain('PATCH /planner/tasks/task-new-1/details');
      expect(Object.keys(planner.items.get('/planner/tasks/task-new-1/details') ?? {})).toEqual(expect.arrayContaining(Object.keys(details)));
    }
  });

  it('writes no details when Graph gives them with no value at all', async () => {
    const planner = await deleteAndUndo(base, { id: 'task-1', description: 5, previewType: null });
    expect(planner.calls.filter((call) => call.includes('task-new-1/details'))).toEqual([]);
  });
});
