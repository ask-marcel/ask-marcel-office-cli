import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import type { TodoStoreFake } from '../../test-helpers/todo-store-fake.ts';
import { createSeededTodoStore, TIME_ZONE } from '../../test-helpers/todo-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The edges of the undo of the To Do commands: undo data that a person edited
 * in the run journal, an item that changed or is gone after the run, the
 * attachments in the two shapes that Graph gives, and the compare of fields.
 */

describe('the undo data of a To Do run', () => {
  const editJournal = (journal: ReturnType<typeof createRunJournalFake>, data: Record<string, unknown>): void => {
    const record = journal.records.get(RUN_ID) as { changes: Array<{ undo: { data: Record<string, unknown> } }> };
    const change = record.changes[0];
    if (change !== undefined) change.undo.data = { ...change.undo.data, ...data };
    journal.records.set(RUN_ID, record);
  };

  it('is refused when a person edited it, and nothing is sent to Graph', async () => {
    const BAD = 'failed: The run journal names an id that is not a Graph id, or data that this CLI did not write, so this undo does not run.';
    const cases: ReadonlyArray<readonly [string, Record<string, string>]> = [
      ['create-todo-task', { list: 'Tasks', title: 'A' }],
      ['complete-todo-task', { list: 'Tasks', taskId: 'task-a' }],
      ['delete-todo-task', { list: 'Tasks', taskId: 'task-a' }],
      ['move-todo-task', { list: 'Tasks', taskId: 'task-a', to: 'Later' }],
      ['add-todo-task-step', { list: 'Tasks', taskId: 'task-a', title: 'x' }],
      ['update-todo-task-step', { list: 'Tasks', taskId: 'task-a', stepId: 'step-a', checked: 'true' }],
      ['delete-todo-task-step', { list: 'Tasks', taskId: 'task-a', stepId: 'step-a' }],
      ['add-todo-linked-resource', { list: 'Tasks', taskId: 'task-a', displayName: 'x', applicationName: 'y' }],
      ['create-todo-list', { name: 'x' }],
      ['update-todo-list', { list: 'Later', name: 'x' }],
    ];
    for (const [name, params] of cases) {
      const todo = createSeededTodoStore();
      const journal = createRunJournalFake();
      await runWrite(todo.graph, name, params, { journal });
      editJournal(journal, { listId: '../x' });
      const writes = todo.calls.length;
      expect([name, statusesOf(await undoWrite(todo.graph, journal, true))]).toEqual([name, [BAD]]);
      expect(todo.calls.slice(writes).filter((call) => !call.startsWith('GET'))).toEqual([]);
    }
  });

  it('refuses an applied undo whose data a person edited after the check', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'add-todo-task-step', { list: 'Tasks', taskId: 'task-a', title: 'x' }, { journal });
    const record = journal.records.get(RUN_ID) as { changes: Array<{ undo: { kind: string } }> };
    const { undoHandlers } = await import('./write-commands.ts');
    const handler = undoHandlers[record.changes[0]?.undo.kind ?? ''];
    expect(await handler?.apply(todo.graph, { listId: '../x' })).toEqual({
      ok: false,
      error: { type: 'validation_error', message: 'The run journal names an id that is not a Graph id, or data that this CLI did not write, so this undo does not run.' },
    });
  });

  it('says missing for the undo of an added attachment that is gone, and changed for one of another name', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    const { createLocalFilesFake } = await import('@ask-marcel/office-core/test-helpers/local-files-fake.ts');
    await runWrite(todo.graph, 'add-todo-task-attachment', { list: 'Tasks', taskId: 'task-a', file: 'a.txt' }, { journal, files: createLocalFilesFake({ 'a.txt': 'abc' }) });
    const [id, file] = [...todo.files.entries()][0] ?? [];
    todo.files.set(id ?? '', { ...file, id: id ?? '', size: 358 });
    expect(statusesOf(await undoWrite(todo.graph, journal, false))).toEqual(['will-undo']);
    todo.files.set(id ?? '', { ...file, id: id ?? '', name: 'b.txt', size: 99 });
    expect(statusesOf(await undoWrite(todo.graph, journal, false))).toEqual([expect.stringMatching(/^changed: /)]);
    todo.files.clear();
    expect(statusesOf(await undoWrite(todo.graph, journal, false))).toEqual([expect.stringMatching(/^missing: /)]);
    todo.tasks.clear();
    expect(statusesOf(await undoWrite(todo.graph, journal, false))).toEqual([expect.stringMatching(/^missing: /)]);
  });

  it('says changed for the undo of a linked resource that changed, and missing when it is gone', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'add-todo-linked-resource', { list: 'Tasks', taskId: 'task-a', displayName: 'x', applicationName: 'y' }, { journal });
    const [id, link] = [...todo.links.entries()][0] ?? [];
    todo.links.set(id ?? '', { ...link, id: id ?? '', displayName: 'z' });
    expect(statusesOf(await undoWrite(todo.graph, journal, false))).toEqual([expect.stringMatching(/^changed: /)]);
    todo.links.clear();
    expect(statusesOf(await undoWrite(todo.graph, journal, false))).toEqual([expect.stringMatching(/^missing: /)]);
  });

  it('says missing for the undo of a step when its task is gone, and for the undo of a field update', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'update-todo-task-step', { list: 'Tasks', taskId: 'task-a', stepId: 'step-a', checked: 'true' }, { journal });
    todo.tasks.clear();
    expect(statusesOf(await undoWrite(todo.graph, journal, false))).toEqual([expect.stringMatching(/^missing: /)]);
  });

  it('refuses the applied undo of a new list too: no command deletes a list', async () => {
    const { undoHandlers } = await import('./write-commands.ts');
    expect(await undoHandlers['todo-list-created']?.apply(createSeededTodoStore().graph, { listId: 'list-a' })).toEqual({
      ok: false,
      error: {
        type: 'validation_error',
        code: 'no_list_delete',
        message: 'No command of this CLI deletes a To Do list, so this undo keeps the list that the run created. Delete it in To Do if you do not want it.',
      },
    });
  });
});

describe('the attachments of a task', () => {
  // Graph documents the answer of one attachment inside `value`; the store answers it bare.
  const wrapped = (todo: TodoStoreFake): WriteGraph => ({
    ...todo.graph,
    get: async (path) => {
      const read = await todo.graph.get(path);
      return /\/attachments\/[^/]+$/.test(path) && read.ok ? ok({ value: read.value }) : read;
    },
  });

  it('reads an attachment that Graph gives inside value, for a move and for a remove', async () => {
    const { createLocalFilesFake } = await import('@ask-marcel/office-core/test-helpers/local-files-fake.ts');
    const todo = createSeededTodoStore();
    await runWrite(todo.graph, 'add-todo-task-attachment', { list: 'Tasks', taskId: 'task-a', file: 'a.txt' }, { files: createLocalFilesFake({ 'a.txt': 'abc' }) });
    expect(await runWrite(wrapped(todo), 'move-todo-task', { list: 'Tasks', taskId: 'task-a', to: 'Later' })).toMatchObject({ done: 1 });
    const moved = [...todo.tasks.keys()][0] ?? '';
    const id = [...todo.files.values()].find((file) => file['taskId'] === moved)?.id ?? '';
    expect(await runWrite(wrapped(todo), 'remove-todo-task-attachment', { list: 'Later', taskId: moved, attachmentId: id })).toMatchObject({ done: 1 });
  });

  it('refuses an attachment whose bytes Graph does not give', async () => {
    const { createLocalFilesFake } = await import('@ask-marcel/office-core/test-helpers/local-files-fake.ts');
    const todo = createSeededTodoStore();
    await runWrite(todo.graph, 'add-todo-task-attachment', { list: 'Tasks', taskId: 'task-a', file: 'a.txt' }, { files: createLocalFilesFake({ 'a.txt': 'abc' }) });
    const [id] = [...todo.files.keys()];
    const graph = { ...todo.graph, get: async (path: string) => (/\/attachments\/[^/]+$/.test(path) ? ok({ name: 'a.txt' }) : todo.graph.get(path)) };
    expect(await runWrite(graph, 'remove-todo-task-attachment', { list: 'Tasks', taskId: 'task-a', attachmentId: id ?? '' })).toMatchObject({
      error: { message: 'Graph gave the bytes of the attachment in a shape that this CLI cannot read.' },
    });
    expect(await runWrite(graph, 'delete-todo-task', { list: 'Tasks', taskId: 'task-a' })).toMatchObject({
      error: { message: 'Graph gave an attachment of a To Do task in a shape that this CLI cannot read.' },
    });
    expect(todo.tasks.has('task-a')).toBe(true);
  });

  it('passes over an attachment that is gone between the plan and the write', async () => {
    const { createLocalFilesFake } = await import('@ask-marcel/office-core/test-helpers/local-files-fake.ts');
    const todo = createSeededTodoStore();
    await runWrite(todo.graph, 'add-todo-task-attachment', { list: 'Tasks', taskId: 'task-a', file: 'a.txt' }, { files: createLocalFilesFake({ 'a.txt': 'abc' }) });
    const [id] = [...todo.files.keys()];
    let lists = 0;
    const graph = {
      ...todo.graph,
      get: async (path: string) => {
        if (path.endsWith('/attachments') && ++lists === 2) todo.files.clear();
        return todo.graph.get(path);
      },
    };
    const run = await runWrite(graph, 'remove-todo-task-attachment', { list: 'Tasks', taskId: 'task-a', attachmentId: id ?? '' });
    expect(run).toMatchObject({ done: 0, skipped: [{ reason: 'The item is no longer there: it was moved or deleted after the plan.' }] });
  });
});

describe('the undo of an update of the fields', () => {
  it('compares the fields with no regard to the order of their keys', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'update-todo-task', { list: 'Tasks', taskId: 'task-a', due: '2026-12-01' }, { journal });
    todo.edit('task-a', { dueDateTime: { timeZone: TIME_ZONE, dateTime: '2026-12-01T00:00:00' } });
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual(['undone']);
  });

  it('gives back no categories as an empty list, and no due day as null', async () => {
    const todo = createSeededTodoStore();
    const created = await todo.graph.post('/me/todo/lists/list-a/tasks', { title: 'Bare' });
    const id = created.ok ? String((created.value as { id: string }).id) : '';
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'update-todo-task', { list: 'Tasks', taskId: id, categories: 'Red', due: '2026-12-01' }, { journal });
    await undoWrite(todo.graph, journal, true);
    expect(todo.tasks.get(id)).toMatchObject({ categories: [], dueDateTime: null });
  });
});
