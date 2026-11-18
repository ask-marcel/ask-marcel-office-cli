import { describe, expect, it } from 'bun:test';
import { createLocalFilesFake } from '@ask-marcel/office-core/test-helpers/local-files-fake.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createSeededTodoStore, TIME_ZONE } from '../../test-helpers/todo-store-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The batch command: the ops of a local JSON file, in the form of the edit
 * script of the user. The plan reads each list and task; one wrong op refuses the whole
 * file before any change, and the refusal never quotes the file.
 */

const opsFile = (value: unknown): ReturnType<typeof createLocalFilesFake> => createLocalFilesFake({ 'ops.json': typeof value === 'string' ? value : JSON.stringify(value) });

const apply = async (value: unknown, options: { readonly dryRun?: boolean; readonly journal?: ReturnType<typeof createRunJournalFake> } = {}): Promise<Record<string, unknown>> => {
  const todo = createSeededTodoStore();
  return runWrite(todo.graph, 'apply-todo-ops', { file: 'ops.json' }, { ...options, files: opsFile(value) });
};

const refusal = async (value: unknown): Promise<string> => ((await apply(value))['error'] as { message: string }).message;

describe('apply-todo-ops', () => {
  it('applies an add, an update with new steps, a move and a delete in one run, and the undo reverses each change', async () => {
    const todo = createSeededTodoStore();
    const old = await todo.graph.post('/me/todo/lists/list-b/tasks', { title: 'Old' });
    const moved = await todo.graph.post('/me/todo/lists/list-b/tasks', { title: 'Moved' });
    const idOf = (answer: typeof old): string => (answer.ok ? String((answer.value as { id: string }).id) : '');
    const files = opsFile({
      ops: [
        { op: 'add', list: 'Later', title: 'New', due: '2026-11-20', categories: ['Red'] },
        { op: 'update', list: 'Tasks', id: 'task-a', set: { status: 'completed', due: null }, addChecklist: ['Print it'] },
        { op: 'delete', list: 'Later', id: idOf(old) },
        { op: 'move', list: 'Later', id: idOf(moved), to: 'Tasks' },
      ],
    });
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'apply-todo-ops', { file: 'ops.json' }, { journal, files });
    expect(run).toMatchObject({ done: 4, failed: [] });
    expect((run['changes'] as ReadonlyArray<{ action: string }>).map((change) => change.action)).toEqual(['create-task', 'update-task', 'add-step', 'delete-task', 'move-task']);
    expect([...todo.tasks.values()].find((task) => task['title'] === 'New')).toMatchObject({
      dueDateTime: { dateTime: '2026-11-20T00:00:00', timeZone: TIME_ZONE },
      categories: ['Red'],
    });
    expect(todo.tasks.get('task-a')).toMatchObject({ status: 'completed', dueDateTime: null });
    const undo = await undoWrite(todo.graph, journal, true);
    expect(statusesOf(undo)).toEqual(['undone', 'undone', 'undone', 'undone', 'undone']);
    expect(todo.tasks.get('task-a')).toMatchObject({ status: 'notStarted', dueDateTime: { dateTime: '2026-11-19T16:00:00.0000000', timeZone: 'UTC' } });
    expect([...todo.tasks.values()].map((task) => [task['listId'], task['title']]).toSorted((a, b) => String(a).localeCompare(String(b)))).toEqual([
      ['list-a', 'Send the report'],
      ['list-b', 'Moved'],
      ['list-b', 'Old'],
    ]);
    expect([...todo.steps.values()].map((step) => step['displayName'])).toEqual(['Read the draft']);
  });

  it('reads the time zone once, and only when an op has a date', async () => {
    const todo = createSeededTodoStore();
    const files = opsFile({
      ops: [
        { op: 'add', list: 'Tasks', title: 'A' },
        { op: 'update', list: 'Tasks', id: 'task-a', set: { importance: 'high' } },
      ],
    });
    await runWrite(todo.graph, 'apply-todo-ops', { file: 'ops.json' }, { dryRun: true, files });
    expect(todo.calls.filter((call) => call.includes('mailboxSettings'))).toEqual([]);
    expect(todo.calls.filter((call) => call.startsWith('GET /me/todo/lists?'))).toHaveLength(1);
  });

  it('goes on after an op that fails, and reports it', async () => {
    const todo = createSeededTodoStore();
    todo.failures.set('POST /me/todo/lists/list-a/tasks', { type: 'api_error', status: 500, message: 'Graph failed.' });
    const files = opsFile({
      ops: [
        { op: 'add', list: 'Tasks', title: 'A' },
        { op: 'update', list: 'Tasks', id: 'task-a', set: { importance: 'high' } },
      ],
    });
    const run = await runWrite(todo.graph, 'apply-todo-ops', { file: 'ops.json' }, { files });
    expect(run).toMatchObject({
      done: 1,
      failed: [{ target: '/me/todo/lists/list-a/tasks', before: { listId: 'list-a', list: 'Tasks' }, reason: 'Graph failed.' }],
      changes: [{ action: 'update-task' }],
    });
  });

  it('records the changes of an op that failed in part, so the undo can reverse them', async () => {
    const todo = createSeededTodoStore();
    todo.failures.set('POST /me/todo/lists/list-a/tasks/task-a/checklistItems', { type: 'api_error', status: 500, message: 'Graph failed.' });
    const journal = createRunJournalFake();
    const files = opsFile({ ops: [{ op: 'update', list: 'Tasks', id: 'task-a', set: { title: 'B' }, addChecklist: ['x'] }] });
    const run = await runWrite(todo.graph, 'apply-todo-ops', { file: 'ops.json' }, { journal, files });
    expect(run).toMatchObject({ done: 0, failed: [{ reason: 'Graph failed.' }], changes: [{ action: 'update-task' }] });
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual(['undone']);
    expect(todo.tasks.get('task-a')?.['title']).toBe('Send the report');
  });

  it('answers the error when no op made a change', async () => {
    const todo = createSeededTodoStore();
    todo.failures.set('POST /me/todo/lists/list-a/tasks', { type: 'api_error', status: 500, message: 'Graph failed.' });
    const run = await runWrite(todo.graph, 'apply-todo-ops', { file: 'ops.json' }, { files: opsFile({ ops: [{ op: 'add', list: 'Tasks', title: 'A' }] }) });
    expect(run['error']).toMatchObject({ message: 'Graph failed.' });
  });

  it('passes over an op whose task is gone between the plan and the write', async () => {
    const todo = createSeededTodoStore();
    let reads = 0;
    const graph = {
      ...todo.graph,
      get: async (path: string) => {
        reads += 1;
        if (reads === 3) todo.tasks.delete('task-a');
        return todo.graph.get(path);
      },
    };
    const files = opsFile({ ops: [{ op: 'update', list: 'Tasks', id: 'task-a', set: { title: 'B' } }] });
    const run = await runWrite(graph, 'apply-todo-ops', { file: 'ops.json' }, { files });
    expect(run).toMatchObject({ done: 0, skipped: [{ reason: 'The item is no longer there: it was moved or deleted after the plan.' }] });
  });

  it('refuses a file that is not JSON, and does not quote it', async () => {
    expect(await refusal('{ "ops": [ secret ')).toBe('--file: the file is not valid JSON. Write it as UTF-8, for example with a file tool or Out-File -Encoding utf8.');
  });

  it('reads a file with a byte order mark', async () => {
    expect(await apply(`${String.fromCodePoint(0xfeff)}${JSON.stringify({ ops: [{ op: 'add', list: 'Tasks', title: 'A' }] })}`, { dryRun: true })).toMatchObject({ dryRun: true });
  });

  it('names the op and the field of each refusal, and quotes no value or key of the file', async () => {
    expect(await refusal({ ops: [] })).toBe('--file: ops must have one op or more.');
    expect(await refusal({ ops: Array.from({ length: 101 }, () => ({ op: 'add', list: 'Tasks', title: 'A' })) })).toBe('--file: ops must have 100 ops or fewer.');
    expect(await refusal({})).toBe('--file: ops is missing.');
    expect(await refusal([])).toBe('--file: the file must be of the type object.');
    expect(
      await refusal({
        ops: [
          { op: 'add', list: 'Tasks', title: 'A' },
          { op: 'rename', list: 'secret' },
        ],
      })
    ).toBe('--file: op 2 op has no "op" that apply-todo-ops knows: add, update, move or delete.');
    expect(await refusal({ ops: [{ op: 'add', list: 'Tasks', title: 'A', secretKey: 1 }] })).toBe('--file: op 1 has a field that apply-todo-ops does not know.');
    expect(await refusal({ ops: [{ op: 'add', list: 'Tasks' }] })).toBe('--file: op 1 title is missing.');
    expect(await refusal({ ops: [{ op: 'add', list: 'Tasks', title: 7 }] })).toBe('--file: op 1 title must be of the type string.');
    expect(await refusal({ ops: [{ op: 'add', list: 'Tasks', title: '  ' }] })).toBe('--file: op 1 title must not be empty.');
    expect(await refusal({ ops: [{ op: 'add', list: 'Tasks', title: 'A', due: '2026-02-30' }] })).toBe('--file: op 1 due must be a day in the form YYYY-MM-DD, or null.');
    expect(await refusal({ ops: [{ op: 'add', list: 'Tasks', title: 'A', reminder: '2026-11-19T24:00' }] })).toBe(
      '--file: op 1 reminder must be a day and a time in the form YYYY-MM-DDTHH:mm, or null.'
    );
    expect(await refusal({ ops: [{ op: 'update', list: 'Tasks', id: 'task-a', addChecklist: ['x'.repeat(256)] }] })).toBe(
      '--file: op 1 addChecklist 0 must have 255 characters or fewer.'
    );
  });

  it('refuses an update with nothing to change, a task in two ops, a move to its own list, and a bad id, without a quote', async () => {
    expect(await refusal({ ops: [{ op: 'update', list: 'Tasks', id: 'task-a' }] })).toBe('--file: op 1 needs "set", "addChecklist", or both.');
    expect(
      await refusal({
        ops: [
          { op: 'delete', list: 'Tasks', id: 'task-a' },
          { op: 'update', list: 'Tasks', id: 'task-a', set: { title: 'B' } },
        ],
      })
    ).toBe('--file: op 2 names the task of op 1. Give each task in one op only.');
    expect(await refusal({ ops: [{ op: 'move', list: 'Tasks', id: 'task-a', to: 'list-a' }] })).toBe('--file: op 1 moves the task to the list it is in.');
    expect(await refusal({ ops: [{ op: 'delete', list: 'Tasks', id: 'secret/../x' }] })).toBe(
      '--file: op 1 id is not a Graph id. Give the id of a task, as list-todo-tasks returns it.'
    );
    expect(await refusal({ ops: [{ op: 'delete', list: 'secret list', id: 'task-a' }] })).toBe(
      'op 1 list: no To Do list has this name or id. list-todo-task-lists gives the name and the id of each list.'
    );
    expect(await refusal({ ops: [{ op: 'move', list: 'Tasks', id: 'task-a', to: 'secret list' }] })).toBe(
      'op 1 to: no To Do list has this name or id. list-todo-task-lists gives the name and the id of each list.'
    );
  });

  it('refuses a file larger than 1 MB', async () => {
    const big = JSON.stringify({ ops: [{ op: 'add', list: 'Tasks', title: 'x'.repeat(1_000_000) }] });
    expect((await apply(big))['error']).toMatchObject({ code: 'local_file_too_large' });
  });
});
