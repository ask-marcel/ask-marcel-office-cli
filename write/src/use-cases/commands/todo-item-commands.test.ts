import { describe, expect, it } from 'bun:test';
import { createLocalFilesFake } from '@ask-marcel/office-core/test-helpers/local-files-fake.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createSeededTodoStore } from '../../test-helpers/todo-store-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The commands on the items of a task (steps, linked resources, attachments)
 * and on the lists, with the undo of each.
 */

const TASK_IDS = { listId: 'list-a', list: 'Tasks', taskId: 'task-a', title: 'Send the report' };
const STEP_BEFORE = { ...TASK_IDS, stepId: 'step-a', step: 'Read the draft', checked: 'false' };
const STEP = { list: 'Tasks', taskId: 'task-a', stepId: 'step-a' };

const stepsOf = (todo: ReturnType<typeof createSeededTodoStore>): ReadonlyArray<unknown> => [...todo.steps.values()].map((step) => [step['displayName'], step['isChecked']]);

describe('add-todo-task-step', () => {
  it('adds a step, not checked by default, and the undo deletes it', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'add-todo-task-step', { list: 'Tasks', taskId: 'task-a', title: 'Print it' }, { journal });
    const id = [...todo.steps.keys()][1];
    expect(run['changes']).toEqual([
      { action: 'add-step', target: '/me/todo/lists/list-a/tasks/task-a/checklistItems', before: TASK_IDS, after: { listId: 'list-a', taskId: 'task-a', stepId: id } },
    ]);
    expect(stepsOf(todo)).toEqual([
      ['Read the draft', false],
      ['Print it', false],
    ]);
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual(['undone']);
    expect(stepsOf(todo)).toEqual([['Read the draft', false]]);
  });

  it('adds a checked step with --checked true, and keeps a step that changed after the run', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'add-todo-task-step', { list: 'Tasks', taskId: 'task-a', title: 'Done before', checked: 'true' }, { journal });
    const id = [...todo.steps.keys()][1] ?? '';
    expect(todo.steps.get(id)?.['isChecked']).toBe(true);
    await todo.graph.patch(`/me/todo/lists/list-a/tasks/task-a/checklistItems/${id}`, { isChecked: false });
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual([expect.stringMatching(/^changed: /)]);
    expect(todo.steps.has(id)).toBe(true);
  });

  it('refuses a --checked that is not true or false', async () => {
    const todo = createSeededTodoStore();
    expect((await runWrite(todo.graph, 'add-todo-task-step', { list: 'Tasks', taskId: 'task-a', title: 'x', checked: 'yes' }))['error']).toMatchObject({
      message: expect.stringMatching(/^--checked /),
    });
  });
});

describe('update-todo-task-step', () => {
  it('checks the step off, and the undo clears the check mark again', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'update-todo-task-step', { ...STEP, checked: 'true' }, { journal });
    expect(run['changes']).toEqual([
      {
        action: 'update-step',
        target: '/me/todo/lists/list-a/tasks/task-a/checklistItems/step-a',
        before: STEP_BEFORE,
        after: { listId: 'list-a', taskId: 'task-a', stepId: 'step-a', checked: 'true' },
      },
    ]);
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual(['undone']);
    expect(stepsOf(todo)).toEqual([['Read the draft', false]]);
  });

  it('changes the text, and does not undo it after the text changed again', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'update-todo-task-step', { ...STEP, title: 'Read the final draft' }, { journal });
    expect(stepsOf(todo)).toEqual([['Read the final draft', false]]);
    await todo.graph.patch('/me/todo/lists/list-a/tasks/task-a/checklistItems/step-a', { displayName: 'By hand' });
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual([expect.stringMatching(/^changed: /)]);
  });

  it('passes over a step that has the text and the check mark already', async () => {
    const todo = createSeededTodoStore();
    const run = await runWrite(todo.graph, 'update-todo-task-step', { ...STEP, title: 'Read the draft', checked: 'false' });
    expect(run).toEqual({
      done: 0,
      skipped: [
        {
          target: '/me/todo/lists/list-a/tasks/task-a/checklistItems/step-a',
          before: { listId: 'list-a', taskId: 'task-a', stepId: 'step-a' },
          reason: 'The step has this text and this check mark already.',
        },
      ],
      failed: [],
      changes: [],
    });
  });

  it('refuses an update with no --title and no --checked, and an id that is not a Graph id', async () => {
    const todo = createSeededTodoStore();
    expect((await runWrite(todo.graph, 'update-todo-task-step', STEP))['error']).toEqual({
      source: 'validation',
      message: 'update-todo-task-step: give --title, --checked, or both.',
    });
    expect((await runWrite(todo.graph, 'update-todo-task-step', { ...STEP, stepId: '../x', checked: 'true' }))['error']).toMatchObject({
      code: 'invalid_id',
      message: expect.stringMatching(/^--step-id: "\.\.\/x" is not a Graph id/),
    });
  });
});

describe('delete-todo-task-step', () => {
  it('deletes the step, and the undo adds it again with a new id', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'delete-todo-task-step', STEP, { journal });
    expect(run['changes']).toEqual([{ action: 'delete-step', target: '/me/todo/lists/list-a/tasks/task-a/checklistItems/step-a', before: STEP_BEFORE, after: {} }]);
    expect(todo.steps.size).toBe(0);
    const undo = await undoWrite(todo.graph, journal, true);
    const id = [...todo.steps.keys()][0];
    expect(undo.ok && undo.value.changes.map((change) => [change.status, change.now])).toEqual([['undone', { listId: 'list-a', taskId: 'task-a', stepId: id }]]);
    expect(stepsOf(todo)).toEqual([['Read the draft', false]]);
  });

  it('says missing when the task of the step is gone', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'delete-todo-task-step', STEP, { journal });
    todo.tasks.delete('task-a');
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual([expect.stringMatching(/^missing: /)]);
  });
});

describe('add-todo-linked-resource', () => {
  it('links a resource, and the undo deletes it', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    const params = { list: 'Tasks', taskId: 'task-a', displayName: 'Spec', applicationName: 'Browser', webUrl: 'https://example.com/spec', externalId: 'spec-1' };
    const run = await runWrite(todo.graph, 'add-todo-linked-resource', params, { journal });
    const [id, link] = [...todo.links.entries()][0] ?? [];
    expect(link).toMatchObject({ displayName: 'Spec', applicationName: 'Browser', webUrl: 'https://example.com/spec', externalId: 'spec-1' });
    expect(run['changes']).toEqual([
      {
        action: 'add-linked-resource',
        target: '/me/todo/lists/list-a/tasks/task-a/linkedResources',
        before: TASK_IDS,
        after: { listId: 'list-a', taskId: 'task-a', linkedResourceId: id },
      },
    ]);
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual(['undone']);
    expect(todo.links.size).toBe(0);
  });

  it('refuses a --web-url that is not http or https', async () => {
    const todo = createSeededTodoStore();
    const params = { list: 'Tasks', taskId: 'task-a', displayName: 'x', applicationName: 'y', webUrl: 'javascript:alert(1)' };
    expect((await runWrite(todo.graph, 'add-todo-linked-resource', params))['error']).toEqual({ source: 'validation', message: '--web-url must be an http or https URL' });
  });
});

describe('add-todo-task-attachment and remove-todo-task-attachment', () => {
  const files = createLocalFilesFake({ 'notes.txt': 'hello' });

  it('attaches a local file, and the undo removes it', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'add-todo-task-attachment', { list: 'Tasks', taskId: 'task-a', file: 'notes.txt' }, { journal, files });
    const id = [...todo.files.keys()][0];
    expect(run['changes']).toEqual([
      {
        action: 'add-attachment',
        target: '/me/todo/lists/list-a/tasks/task-a/attachments',
        before: { ...TASK_IDS, file: 'notes.txt', size: '5' },
        after: { listId: 'list-a', taskId: 'task-a', attachmentId: id },
      },
    ]);
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual(['undone']);
    expect(todo.files.size).toBe(0);
  });

  it('refuses a missing file on a dry-run, before any read of Graph', async () => {
    const todo = createSeededTodoStore();
    const dry = await runWrite(todo.graph, 'add-todo-task-attachment', { list: 'Tasks', taskId: 'task-a', file: 'none.txt' }, { dryRun: true, files });
    expect(dry['error']).toMatchObject({ code: 'local_file_not_found' });
    expect(todo.calls).toEqual([]);
  });

  it('removes an attachment, and the undo attaches the same bytes again under a new id', async () => {
    const todo = createSeededTodoStore();
    await runWrite(todo.graph, 'add-todo-task-attachment', { list: 'Tasks', taskId: 'task-a', file: 'notes.txt' }, { files });
    const id = [...todo.files.keys()][0] ?? '';
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'remove-todo-task-attachment', { list: 'Tasks', taskId: 'task-a', attachmentId: id }, { journal });
    expect(run['changes']).toEqual([
      {
        action: 'remove-attachment',
        target: `/me/todo/lists/list-a/tasks/task-a/attachments/${id}`,
        before: { ...TASK_IDS, attachmentId: id, file: 'notes.txt', size: '5' },
        after: {},
      },
    ]);
    expect(todo.files.size).toBe(0);
    const undo = await undoWrite(todo.graph, journal, true);
    const back = [...todo.files.values()][0];
    expect(undo.ok && undo.value.changes.map((change) => [change.status, change.now])).toEqual([['undone', { listId: 'list-a', taskId: 'task-a', attachmentId: back?.id }]]);
    expect([back?.['name'], back?.['contentBytes']]).toEqual(['notes.txt', 'aGVsbG8=']);
  });

  it('refuses an attachment id that the task does not have', async () => {
    const todo = createSeededTodoStore();
    expect((await runWrite(todo.graph, 'remove-todo-task-attachment', { list: 'Tasks', taskId: 'task-a', attachmentId: 'file-9' }))['error']).toEqual({
      source: 'validation',
      code: 'todo_attachment_not_found',
      message: '--attachment-id: the task has no attachment with this id. get-todo-task with --expand attachments gives the ids.',
    });
  });
});

describe('create-todo-list and update-todo-list', () => {
  it('creates a list, and its undo keeps the list and says why', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'create-todo-list', { name: 'Errands' }, { journal });
    const id = [...todo.lists.keys()][2];
    expect(run['changes']).toEqual([{ action: 'create-list', target: '/me/todo/lists', before: {}, after: { listId: id, list: 'Errands' } }]);
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual([
      'failed: No command of this CLI deletes a To Do list, so this undo keeps the list that the run created. Delete it in To Do if you do not want it.',
    ]);
    expect(todo.lists.has(id ?? '')).toBe(true);
  });

  it('gives a list a new name, and the undo gives the name of before back', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'update-todo-list', { list: 'Later', name: 'Someday' }, { journal });
    expect(run['changes']).toEqual([
      { action: 'rename-list', target: '/me/todo/lists/list-b', before: { listId: 'list-b', list: 'Later' }, after: { listId: 'list-b', list: 'Someday' } },
    ]);
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual(['undone']);
    expect(todo.lists.get('list-b')?.['displayName']).toBe('Later');
  });

  it('does not give the name back when the name changed after the run, or when the list is gone', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'update-todo-list', { list: 'Later', name: 'Someday' }, { journal });
    await todo.graph.patch('/me/todo/lists/list-b', { displayName: 'By hand' });
    expect(statusesOf(await undoWrite(todo.graph, journal, false))).toEqual([expect.stringMatching(/^changed: /)]);
    todo.lists.delete('list-b');
    expect(statusesOf(await undoWrite(todo.graph, journal, false))).toEqual([expect.stringMatching(/^missing: /)]);
  });

  it('passes over a list that has the name already', async () => {
    const todo = createSeededTodoStore();
    const dry = await runWrite(todo.graph, 'update-todo-list', { list: 'list-b', name: 'Later' }, { dryRun: true });
    expect(dry['skipped']).toEqual([{ target: '/me/todo/lists/list-b', before: { listId: 'list-b', list: 'Later' }, reason: 'The list has this name already.' }]);
  });
});
