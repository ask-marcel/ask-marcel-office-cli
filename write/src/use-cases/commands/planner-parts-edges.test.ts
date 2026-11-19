import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createLocalBytesFake, utf8Bytes } from '../../test-helpers/local-bytes-fake.ts';
import { createPlannerStoreFake } from '../../test-helpers/planner-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The edges of the parts of a task (checklist, links, files) and of a plan:
 * the words of each refusal, what the dry-run says, the type of a file, and
 * the labels of a plan in odd shapes.
 */

const DETAILS = '/planner/tasks/task-1/details';
const UUID = /^[\da-f]{8}-[\da-f]{4}-[\da-f]{4}-[\da-f]{4}-[\da-f]{12}$/;

const beforeOf = (run: Record<string, unknown>): unknown => (run['changes'] as ReadonlyArray<{ before: unknown }>)[0]?.before;

describe('the checklist commands', () => {
  it('give the words of their refusals', async () => {
    const planner = createPlannerStoreFake();
    const run = async (name: string, params: Record<string, string>): Promise<unknown> => (await runWrite(planner.graph, name, { taskId: 'task-1', ...params }))['error'];
    expect(await run('add-planner-checklist-item', { title: 'Draft' })).toMatchObject({ message: '--title: the checklist of the task has an item with this title.' });
    expect(await run('update-planner-checklist-item', { item: 'c1' })).toMatchObject({ message: 'Give --title or --checked: there is nothing to change.' });
    expect(await run('remove-planner-checklist-item', { item: 'draft' })).toMatchObject({
      message: '--item: the checklist of the task has no item with this id or title (the title must be exact).',
    });
    planner.edit(DETAILS, { checklist: { c3: { title: 'Draft', isChecked: false } } });
    expect(await run('remove-planner-checklist-item', { item: 'Draft' })).toMatchObject({
      message: '--item: 2 items of the checklist have this title. Give the id of the item, as get-planner-task-details returns it.',
    });
  });

  it('add an item not checked by default, and say what each one changes', async () => {
    const planner = createPlannerStoreFake();
    const add = await runWrite(planner.graph, 'add-planner-checklist-item', { taskId: 'task-1', title: 'Print' });
    const key = String((beforeOf(add) as { itemId: string }).itemId);
    expect(key).toMatch(UUID);
    expect(beforeOf(add)).toEqual({ taskId: 'task-1', itemId: key, title: 'Print' });
    expect(planner.items.get(DETAILS)?.['checklist']).toMatchObject({ [key]: { title: 'Print', isChecked: false } });
    const update = await runWrite(planner.graph, 'update-planner-checklist-item', { taskId: 'task-1', item: 'Draft', title: 'Draft 2' }, { dryRun: true });
    expect(beforeOf(update)).toEqual({ taskId: 'task-1', itemId: 'c1', title: 'Draft' });
    const remove = await runWrite(planner.graph, 'remove-planner-checklist-item', { taskId: 'task-1', item: 'c2' }, { dryRun: true });
    expect(beforeOf(remove)).toEqual({ taskId: 'task-1', itemId: 'c2', title: 'Send' });
  });

  it('give an item that the undo adds again a new id of the UUID form', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'remove-planner-checklist-item', { taskId: 'task-1', item: 'c1' }, { journal });
    await undoWrite(planner.graph, journal, true);
    const keys = Object.keys(planner.items.get(DETAILS)?.['checklist'] as object);
    expect(keys.filter((key) => key !== 'c2')).toEqual([expect.stringMatching(UUID)]);
  });

  it('count the items added, removed and changed of a checklist file, and refuse a file that changes nothing', async () => {
    const planner = createPlannerStoreFake();
    const files = createLocalBytesFake({ 'a.txt': utf8Bytes('[ ] Draft\nNew'), 'b.txt': utf8Bytes('[x] Draft\n[ ] Send') });
    const dry = await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', checklistFile: 'a.txt' }, { files, dryRun: true });
    expect(beforeOf(dry)).toEqual({ taskId: 'task-1', fields: 'checklist', checklist: '1 added, 1 removed, 1 changed' });
    expect(await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', checklistFile: 'b.txt' }, { files })).toMatchObject({
      error: { message: 'The checklist of the file is the checklist of the task: there is nothing to change.' },
    });
  });

  it('change only the preview, and keep the notes as HTML', async () => {
    const planner = createPlannerStoreFake();
    expect(await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', preview: 'description' })).toMatchObject({ done: 1 });
    await runWrite(planner.graph, 'update-planner-task-details', { taskId: 'task-1', notesHtml: '<p>x</p>' });
    expect(planner.items.get(DETAILS)).toMatchObject({ previewType: 'description', notes: { contentType: 'html', content: '<p>x</p>' } });
  });

  it('mark the undo failed when the read of the item fails after its check', async () => {
    const planner = createPlannerStoreFake();
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'add-planner-checklist-item', { taskId: 'task-1', title: 'Print' }, { journal });
    let reads = 0;
    planner.hooks.afterRead = (_count, path) => {
      if (path === DETAILS && ++reads === 1) planner.failures.set(`GET ${DETAILS}`, { type: 'api_error', status: 503, message: 'Try again later.' });
    };
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['failed: Try again later.']);
  });
});

describe('the link and file commands', () => {
  it('give the URL of a link in the dry-run of its remove, and refuse a bad URL with the words of the option', async () => {
    const planner = createPlannerStoreFake();
    const dry = await runWrite(planner.graph, 'remove-planner-task-link', { taskId: 'task-1', url: 'https://example.com/spec' }, { dryRun: true });
    expect(beforeOf(dry)).toEqual({ taskId: 'task-1', url: 'https://example.com/spec' });
    expect(await runWrite(planner.graph, 'remove-planner-task-link', { taskId: 'task-1', url: 'mailto:a@b.c' })).toMatchObject({
      error: { message: '--url: give a full http or https URL, for example https://example.com/page.' },
    });
    expect(await runWrite(planner.graph, 'remove-planner-task-link', { taskId: 'task-1', url: 'https://example.com/none' })).toMatchObject({
      error: { message: '--url: the task has no link or file to this URL. get-planner-task-details gives its references.' },
    });
  });

  it('give a Word, Excel or PowerPoint file its type, and another file the type Other', async () => {
    const names = ['a.doc', 'b.DOCX', 'c.xls', 'd.xlsx', 'e.ppt', 'f.pptx', 'g.pdf', 'h'];
    const files = createLocalBytesFake(Object.fromEntries(names.map((name) => [name, utf8Bytes('x')])));
    const types: string[] = [];
    for (const name of names) {
      const planner = createPlannerStoreFake();
      await runWrite(planner.graph, 'add-planner-task-file', { taskId: 'task-1', file: name }, { files });
      const references = planner.items.get(DETAILS)?.['references'] as Record<string, { type: string; alias: string }>;
      types.push(Object.values(references).find((reference) => reference.alias === name)?.type ?? '');
    }
    expect(types).toEqual(['Word', 'Word', 'Excel', 'Excel', 'PowerPoint', 'PowerPoint', 'Other', 'Other']);
  });

  it('refuse a file name that OneDrive cannot keep, with the words of the rule', async () => {
    const long = `${'a'.repeat(196)}.txt`;
    const files = createLocalBytesFake({ [long]: utf8Bytes('x'), [`a${long}`]: utf8Bytes('x'), 'dir/ lead.txt': utf8Bytes('x'), 'dot.': utf8Bytes('x') });
    const codeOf = async (file: string): Promise<unknown> =>
      (await runWrite(createPlannerStoreFake().graph, 'add-planner-task-file', { taskId: 'task-1', file }, { files }))['error'];
    expect(await codeOf(long)).toBeUndefined();
    const refusal = {
      code: 'planner_file_name_invalid',
      message: '--file: OneDrive cannot keep a file of this name. Rename the file: no " * : < > ? / \\ |, no space at the start or the end, at most 200 characters.',
    };
    for (const file of [`a${long}`, 'dir/ lead.txt', 'dot.']) expect(await codeOf(file)).toMatchObject(refusal);
  });

  it('attach a file to a plan of the user that has a container and an owner, and give the words of the group refusal', async () => {
    const files = createLocalBytesFake({ 'a.txt': utf8Bytes('x') });
    const planner = createPlannerStoreFake();
    planner.edit('/planner/plans/plan-1', { container: { type: 'unknownFutureValue' }, owner: 'user-1' });
    expect(await runWrite(planner.graph, 'add-planner-task-file', { taskId: 'task-1', file: 'a.txt' }, { files })).toMatchObject({ done: 1 });
    planner.edit('/planner/plans/plan-1', { container: { type: 'group' } });
    expect(await runWrite(planner.graph, 'add-planner-task-file', { taskId: 'task-1', file: 'a.txt' }, { files })).toMatchObject({
      error: {
        message:
          'The plan is of a group: Planner keeps its files in the document library of the group, which this CLI cannot write. Attach the file in Planner, or give a link to a file that is in SharePoint or OneDrive already with add-planner-task-link.',
      },
    });
  });

  it('undo a file that OneDrive gave with no content tag, and refuse a file name that the folder has', async () => {
    const files = createLocalBytesFake({ 'a.txt': utf8Bytes('x') });
    const planner = createPlannerStoreFake();
    planner.answers.set('PUT /me/drive/root:/Planner%20attachments/a.txt', { id: 'drive-9', webUrl: 'https://x.example/a.txt' });
    planner.drive.set('/me/drive/items/drive-9', { id: 'drive-9' });
    const journal = createRunJournalFake();
    await runWrite(planner.graph, 'add-planner-task-file', { taskId: 'task-1', file: 'a.txt' }, { files, journal });
    expect(statusesOf(await undoWrite(planner.graph, journal, true))).toEqual(['undone', 'undone']);
    planner.drive.set('/me/drive/items/old', { id: 'old', path: '/me/drive/root:/Planner%20attachments/a.txt' });
    expect(await runWrite(planner.graph, 'add-planner-task-file', { taskId: 'task-1', file: 'a.txt' }, { files })).toMatchObject({
      error: {
        message: '--file: the folder "Planner attachments" of your OneDrive has a file of this name. Rename the file, or give a link to that file with add-planner-task-link.',
      },
    });
  });
});

describe('the labels and the names of a plan', () => {
  it('know only the named labels of the plan, and give the item of a label that is not one', async () => {
    const planner = createPlannerStoreFake();
    planner.edit('/planner/plans/plan-1/details', { categoryDescriptions: { category1: null, category2: '', category3: 'Green', note: 'Red' } });
    const create = async (labels: string): Promise<Record<string, unknown>> =>
      runWrite(planner.graph, 'create-planner-task', { plan: 'plan-1', bucket: 'bucket-a', title: 'x', labels }, { dryRun: true });
    expect(await create('GREEN, category2')).toMatchObject({ dryRun: true });
    expect(await create('green,red')).toMatchObject({
      error: { message: '--labels: item 2 is not a label of the plan: give category1 to category25, or a label name of the plan.' },
    });
  });

  it('read no label names for no label, and give one key once', async () => {
    const planner = createPlannerStoreFake();
    await runWrite(planner.graph, 'update-planner-task', { taskId: 'task-1', labels: ' , ' });
    expect(planner.calls.filter((call) => call.endsWith('/details'))).toEqual([]);
    expect(planner.items.get('/planner/tasks/task-1')?.['appliedCategories']).toEqual({ category1: false });
    await runWrite(planner.graph, 'update-planner-task', { taskId: 'task-1', labels: 'Red,category1' });
    expect(planner.items.get('/planner/tasks/task-1')?.['appliedCategories']).toEqual({ category1: true });
  });

  it('rename only the plan, or only its labels, in one change each, with the label name trimmed', async () => {
    const planner = createPlannerStoreFake();
    const name = await runWrite(planner.graph, 'update-planner-plan', { plan: 'plan-1', name: 'Next' }, { dryRun: true });
    expect((name['changes'] as ReadonlyArray<{ action: string }>).map((change) => change.action)).toEqual(['rename-planner-plan']);
    expect(planner.calls.filter((call) => call.endsWith('/details'))).toEqual([]);
    const labels = await runWrite(planner.graph, 'update-planner-plan', { plan: 'plan-1', labelNames: 'category1= Urgent ' });
    expect((labels['changes'] as ReadonlyArray<{ action: string }>).map((change) => change.action)).toEqual(['rename-planner-labels']);
    expect(planner.items.get('/planner/plans/plan-1/details')?.['categoryDescriptions']).toMatchObject({ category1: 'Urgent' });
    expect(await runWrite(planner.graph, 'update-planner-plan', { plan: 'plan-1' })).toMatchObject({
      error: { message: 'Give --name or --label-names: there is nothing to change.' },
    });
    expect(await runWrite(planner.graph, 'update-planner-plan', { plan: 'plan-1', labelNames: 'category1' })).toMatchObject({ error: { code: 'planner_label_names_invalid' } });
  });

  it('give the words of a bucket that the plan does not have', async () => {
    const planner = createPlannerStoreFake();
    expect(await runWrite(planner.graph, 'update-planner-task', { taskId: 'task-1', bucket: 'Nowhere' })).toMatchObject({
      error: { message: '--bucket: the plan has no bucket with this name or id. list-plan-buckets gives the name and the id of each bucket.' },
    });
    expect(RUN_ID).toBeDefined();
  });
});
