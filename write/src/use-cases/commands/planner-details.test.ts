import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createLocalBytesFake, UTF8_BOM, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { createPlannerStoreFake } from '../../test-helpers/planner-store-fake.ts';
import type { Item } from '../../test-helpers/planner-store-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * update-planner-task-details: the plain description on Graph v1.0, the HTML
 * notes on Graph beta (the one beta path), the preview, and the checklist of
 * a file. The undo writes the details of before back.
 */

const DETAILS = '/planner/tasks/task-1/details';

const checklistOf = (details: Item | undefined): ReadonlyArray<[string, boolean]> =>
  Object.values((details?.['checklist'] ?? {}) as Record<string, { title: string; isChecked: boolean }>)
    .map((item): [string, boolean] => [item.title, item.isChecked])
    .toSorted((a, b) => a[0].localeCompare(b[0]));

describe('update-planner-task-details', () => {
  it('sends a description to Graph v1.0 with the ETag, and the undo writes the description of before back', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    const run = await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', description: 'New text', preview: 'noPreview' }, { journal });
    expect(run).toMatchObject({ done: 1, changes: [{ action: 'update-planner-task-details', target: DETAILS, before: { taskId: 'task-1', fields: 'description, previewType' } }] });
    expect(planner.calls).toContain(`PATCH ${DETAILS}`);
    expect(planner.items.get(DETAILS)).toMatchObject({ description: 'New text', previewType: 'noPreview' });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(planner.items.get(DETAILS)).toMatchObject({ description: 'The text', previewType: 'automatic' });
  });

  it('sends HTML notes to Graph beta, and the undo writes the plain text of before back on v1.0', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', notesHtml: '<p>See the <b>budget</b></p>' }, { journal });
    expect(planner.calls).toContain(`PATCH ${DETAILS} [beta]`);
    expect(planner.items.get(DETAILS)).toMatchObject({ notes: { contentType: 'html', content: '<p>See the <b>budget</b></p>' }, description: 'See the budget' });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(planner.calls.at(-1)).toBe(`PATCH ${DETAILS}`);
    expect(planner.items.get(DETAILS)).toMatchObject({ description: 'The text' });
  });

  it('refuses a description with notes, and no change, before any read', async () => {
    const planner = createPlannerStoreFake();
    expect(await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', description: 'a', notesHtml: '<p>b</p>' })).toMatchObject({
      error: { code: 'planner_description_and_notes', message: 'Graph refuses a change of the description and the notes in one request: give --description or --notes-html.' },
    });
    expect(await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1' })).toMatchObject({
      error: { code: 'planner_nothing_to_change', message: 'Give --description, --notes-html, --checklist-file or --preview: there is nothing to change.' },
    });
    expect(planner.calls).toEqual([]);
  });

  it('makes the checklist the list of the file, and the undo puts the checklist of before back', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    const files = createLocalBytesFake({ 'steps.txt': utf8Bytes('[ ] Draft\nCall the team\n', UTF8_BOM) });
    const run = await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', checklistFile: 'steps.txt' }, { journal, files });
    expect(run).toMatchObject({ changes: [{ before: { taskId: 'task-1', fields: 'checklist', checklist: '1 added, 1 removed, 1 changed' } }] });
    expect(checklistOf(planner.items.get(DETAILS))).toEqual([
      ['Call the team', false],
      ['Draft', false],
    ]);
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(checklistOf(planner.items.get(DETAILS))).toEqual([
      ['Draft', true],
      ['Send', false],
    ]);
  });

  it('refuses a checklist file that is not UTF-8, that is missing, or that changes nothing', async () => {
    const planner = createPlannerStoreFake();
    const files = createLocalBytesFake({ 'bad.txt': Uint8Array.from([0xff, 0xfe, 0x00]), 'same.txt': utf8Bytes('[x] Draft\nSend') });
    const codeOf = async (path: string): Promise<unknown> =>
      ((await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', checklistFile: path }, { files })) as { error: { code?: string } }).error.code;
    expect(await codeOf('bad.txt')).toBe('checklist_file_not_utf8');
    expect(await codeOf('none.txt')).toBe('local_file_not_found');
    expect(await codeOf('same.txt')).toBe('planner_nothing_to_change');
    expect(planner.calls.filter((call) => call.startsWith('PATCH'))).toEqual([]);
  });
});

describe('the checklist commands', () => {
  it('adds an item, and the undo removes it', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    const run = await runWrite(planner.graph, 'add-planner-checklist-item', { taskId: 'task-1', title: 'Book the room', checked: 'true' }, { journal });
    expect(run).toMatchObject({ done: 1, changes: [{ action: 'add-planner-checklist-item', before: { taskId: 'task-1', title: 'Book the room' } }] });
    expect(checklistOf(planner.items.get(DETAILS))).toContainEqual(['Book the room', true]);
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(checklistOf(planner.items.get(DETAILS))).toEqual([
      ['Draft', true],
      ['Send', false],
    ]);
  });

  it('refuses a title that the checklist has, and a checklist of 20 items', async () => {
    const planner = createPlannerStoreFake();
    expect(await runWrite(planner.graph, 'add-planner-checklist-item', { taskId: 'task-1', title: 'Send' })).toMatchObject({ error: { code: 'planner_checklist_item_exists' } });
    planner.edit(DETAILS, { checklist: Object.fromEntries(Array.from({ length: 18 }, (_, index) => [`k${index}`, { title: `Item ${index}`, isChecked: false }])) });
    expect(await runWrite(planner.graph, 'add-planner-checklist-item', { taskId: 'task-1', title: 'One more' })).toMatchObject({
      error: { code: 'planner_checklist_full', message: 'The checklist of the task has 20 items, the most that Planner keeps.' },
    });
  });

  it('checks an item named by its title, and the undo unchecks it', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'update-planner-checklist-item', { taskId: 'task-1', item: 'Send', checked: 'true', title: 'Send it' }, { journal });
    expect(planner.items.get(DETAILS)?.['checklist']).toMatchObject({ c2: { title: 'Send it', isChecked: true } });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    expect(planner.items.get(DETAILS)?.['checklist']).toMatchObject({ c2: { title: 'Send', isChecked: false } });
  });

  it('refuses an item that is not there, a title of two items, and no change', async () => {
    const planner = createPlannerStoreFake();
    const codeOf = async (params: Record<string, string>): Promise<unknown> =>
      ((await runWrite(planner.graph, 'update-planner-checklist-item', { taskId: 'task-1', ...params })) as { error: { code?: string } }).error.code;
    expect(await codeOf({ item: 'send', checked: 'true' })).toBe('planner_checklist_item_not_found');
    expect(await codeOf({ item: 'Send' })).toBe('planner_nothing_to_change');
    planner.edit(DETAILS, { checklist: { c3: { title: 'Send', isChecked: true } } });
    expect(await codeOf({ item: 'Send', checked: 'true' })).toBe('planner_checklist_item_ambiguous');
  });

  it('removes an item named by its id, and the undo adds it again with a new id', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'remove-planner-checklist-item', { taskId: 'task-1', item: 'c1' }, { journal });
    expect(checklistOf(planner.items.get(DETAILS))).toEqual([['Send', false]]);
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone']);
    const checklist = planner.items.get(DETAILS)?.['checklist'] as Record<string, unknown>;
    expect(checklistOf(planner.items.get(DETAILS))).toEqual([
      ['Draft', true],
      ['Send', false],
    ]);
    expect(Object.hasOwn(checklist, 'c1')).toBe(false);
  });
});
