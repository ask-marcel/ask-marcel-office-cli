import { describe, expect, it } from 'bun:test';
import { createLocalBytesFake, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { createPlannerStoreFake } from '../../test-helpers/planner-store-fake.ts';
import type { Item } from '../../test-helpers/planner-store-fake.ts';
import { runWrite } from '../../test-helpers/run-write.ts';

/*
 * The checklist of a task as the commands read it: an entry that is not an
 * item does not count, a hint that is not a string is not the last hint, a
 * changed item keeps its hint, and a checklist file that is not good is
 * refused with its words.
 */

const DETAILS = '/planner/tasks/task-1/details';
const checklistOf = (planner: ReturnType<typeof createPlannerStoreFake>): Record<string, Item> => planner.items.get(DETAILS)?.['checklist'] as Record<string, Item>;

describe('the checklist of a task', () => {
  it('counts only the entries that are items, so a task with 19 items and one odd entry takes one more', async () => {
    const planner = createPlannerStoreFake();
    const items = Object.fromEntries(Array.from({ length: 19 }, (_, index) => [`k${index}`, { title: `Item ${index}`, isChecked: false }]));
    planner.items.set(DETAILS, { id: 'task-1', '@odata.etag': 'W/"d"', checklist: { ...items, odd: { title: 5 } } });
    expect(await runWrite(planner.graph, 'add-planner-checklist-item', { taskId: 'task-1', title: 'Twenty' })).toMatchObject({ done: 1 });
  });

  it('puts a new item after the greatest hint that is a string', async () => {
    const planner = createPlannerStoreFake();
    planner.items.set(DETAILS, {
      id: 'task-1',
      '@odata.etag': 'W/"d"',
      checklist: { c2: { title: 'B', isChecked: false, orderHint: 'b' }, c1: { title: 'A', isChecked: false, orderHint: 'a' }, c3: { title: 'C', isChecked: false, orderHint: 7 } },
    });
    await runWrite(planner.graph, 'add-planner-checklist-item', { taskId: 'task-1', title: 'D' });
    expect(Object.values(checklistOf(planner)).find((item) => item['title'] === 'D')?.['orderHint']).toBe('b !');
  });

  it('keeps the hint of an item that the file changes', async () => {
    const planner = createPlannerStoreFake();
    const files = createLocalBytesFake({ 'a.txt': utf8Bytes('[ ] Draft\nSend') });
    await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', checklistFile: 'a.txt' }, { files });
    expect(checklistOf(planner)['c1']).toMatchObject({ title: 'Draft', isChecked: false, orderHint: '2' });
  });

  it('counts the items that a file adds, removes and changes', async () => {
    const planner = createPlannerStoreFake();
    const files = createLocalBytesFake({ 'a.txt': utf8Bytes('[ ] Draft\n[x] Send\nNew') });
    const dry = await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', checklistFile: 'a.txt' }, { files, dryRun: true });
    expect(dry).toMatchObject({ changes: [{ before: { checklist: '1 added, 0 removed, 2 changed' } }] });
  });

  it('refuses a file that is not UTF-8, and a file with no item, with their words', async () => {
    const planner = createPlannerStoreFake();
    const files = createLocalBytesFake({ 'bad.txt': Uint8Array.from([0xff, 0xfe]), 'empty.txt': utf8Bytes('\n \n') });
    expect(await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', checklistFile: 'bad.txt' }, { files })).toMatchObject({
      error: { code: 'checklist_file_not_utf8', source: 'validation', message: '--checklist-file is not UTF-8 text. Save it as UTF-8 (a byte order mark is removed).' },
    });
    expect(await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', checklistFile: 'empty.txt' }, { files })).toMatchObject({
      error: { code: 'checklist_file_empty' },
    });
  });
});
