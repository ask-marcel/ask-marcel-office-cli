import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createSeededTodoStore, TIME_ZONE } from '../../test-helpers/todo-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The commands that create, change and complete a task. Each shows the task in a dry-run,
 * writes it, records the run, and its undo reverses the change only if the
 * task is as the run left it.
 */

const writes = (calls: ReadonlyArray<string>): ReadonlyArray<string> => calls.filter((call) => !call.startsWith('GET'));

const TASK_BEFORE = { listId: 'list-a', list: 'Tasks', taskId: 'task-a', title: 'Send the report', status: 'notStarted' };

describe('create-todo-task', () => {
  it('creates a task with its fields in the time zone of the mailbox, and the undo deletes it', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    const params = { list: 'tasks', title: 'Book the room', due: '2026-11-20', reminder: '2026-11-19T09:30', importance: 'high', categories: 'Red, Blue', body: 'Floor 3' };
    const run = await runWrite(todo.graph, 'create-todo-task', params, { journal });
    expect(run).toMatchObject({ done: 1, skipped: [], failed: [], runId: RUN_ID });
    const created = [...todo.tasks.values()].find((task) => task['title'] === 'Book the room');
    expect(created).toMatchObject({
      listId: 'list-a',
      importance: 'high',
      categories: ['Red', 'Blue'],
      body: { content: 'Floor 3', contentType: 'text' },
      dueDateTime: { dateTime: '2026-11-20T00:00:00', timeZone: TIME_ZONE },
      reminderDateTime: { dateTime: '2026-11-19T09:30:00', timeZone: TIME_ZONE },
      isReminderOn: true,
    });
    expect(run['changes']).toEqual([
      {
        action: 'create-task',
        target: '/me/todo/lists/list-a/tasks',
        before: { listId: 'list-a', list: 'Tasks' },
        after: { listId: 'list-a', taskId: created?.id, title: 'Book the room' },
      },
    ]);
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual(['undone']);
    expect(todo.tasks.has(created?.id ?? '')).toBe(false);
  });

  it('shows the list on a dry-run, reads no time zone without a date, and creates nothing', async () => {
    const todo = createSeededTodoStore();
    const dry = await runWrite(todo.graph, 'create-todo-task', { list: 'list-b', title: 'Call back' }, { dryRun: true });
    expect(dry).toEqual({
      dryRun: true,
      command: 'create-todo-task',
      changes: [{ action: 'create-task', target: '/me/todo/lists/list-b/tasks', before: { listId: 'list-b', list: 'Later' } }],
    });
    expect(todo.calls).toEqual(['GET /me/todo/lists?$top=200']);
  });

  it('sends only the fields given: a task with a title alone has no other field', async () => {
    const todo = createSeededTodoStore();
    await runWrite(todo.graph, 'create-todo-task', { list: 'Tasks', title: 'Plain' });
    const created = [...todo.tasks.values()].find((task) => task['title'] === 'Plain');
    expect(Object.keys(created ?? {}).toSorted((a, b) => a.localeCompare(b))).toEqual(['id', 'importance', 'lastModifiedDateTime', 'listId', 'status', 'title']);
  });

  it('keeps the task when it changed after the run', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'create-todo-task', { list: 'Tasks', title: 'Plain' }, { journal });
    const id = (run['changes'] as ReadonlyArray<{ after: { taskId: string } }>)[0]?.after.taskId ?? '';
    todo.edit(id, { title: 'Plain, by hand' });
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual([
      'changed: The item changed after the run, so the undo passes over it. Look at it and reverse it by hand if you still want to.',
    ]);
    expect(todo.tasks.has(id)).toBe(true);
  });

  it('says missing when the created task is gone', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'create-todo-task', { list: 'Tasks', title: 'Plain' }, { journal });
    todo.tasks.delete((run['changes'] as ReadonlyArray<{ after: { taskId: string } }>)[0]?.after.taskId ?? '');
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual(['missing: The item is no longer there, so there is nothing to undo.']);
  });
});

describe('update-todo-task', () => {
  it('writes only the fields given, removes the due day with none, and the undo writes the fields of before back', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'update-todo-task', { list: 'Tasks', taskId: 'task-a', due: 'none', status: 'inProgress', categories: '' }, { journal });
    expect(run['changes']).toEqual([
      {
        action: 'update-task',
        target: '/me/todo/lists/list-a/tasks/task-a',
        before: TASK_BEFORE,
        after: { listId: 'list-a', taskId: 'task-a', fields: 'status, categories, dueDateTime' },
      },
    ]);
    expect(todo.tasks.get('task-a')).toMatchObject({ status: 'inProgress', categories: [], dueDateTime: null, title: 'Send the report' });
    expect(writes(todo.calls)).toEqual(['PATCH /me/todo/lists/list-a/tasks/task-a']);
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual(['undone']);
    expect(todo.tasks.get('task-a')).toMatchObject({ status: 'notStarted', categories: ['Blue'], dueDateTime: { dateTime: '2026-11-19T16:00:00.0000000', timeZone: 'UTC' } });
  });

  it('turns the reminder off with none, and gives an empty field of before back as null', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'update-todo-task', { list: 'Tasks', taskId: 'task-a', reminder: 'none', start: '2026-11-18' }, { journal });
    expect(todo.tasks.get('task-a')).toMatchObject({ isReminderOn: false, startDateTime: { dateTime: '2026-11-18T00:00:00', timeZone: TIME_ZONE } });
    await undoWrite(todo.graph, journal, true);
    expect(todo.tasks.get('task-a')).toMatchObject({ isReminderOn: false, startDateTime: null });
  });

  it('undoes the fields although a step of the task changed after the run', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'update-todo-task', { list: 'Tasks', taskId: 'task-a', title: 'Send the final report' }, { journal });
    await runWrite(todo.graph, 'add-todo-task-step', { list: 'Tasks', taskId: 'task-a', title: 'Print it' });
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual(['undone']);
    expect(todo.tasks.get('task-a')?.['title']).toBe('Send the report');
  });

  it('does not undo a field that changed after the run', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'update-todo-task', { list: 'Tasks', taskId: 'task-a', title: 'Send the final report' }, { journal });
    todo.edit('task-a', { title: 'Another title' });
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual([expect.stringMatching(/^changed: /)]);
    expect(todo.tasks.get('task-a')?.['title']).toBe('Another title');
  });

  it('refuses an update with no field to change, and reads nothing', async () => {
    const todo = createSeededTodoStore();
    expect((await runWrite(todo.graph, 'update-todo-task', { list: 'Tasks', taskId: 'task-a' }))['error']).toEqual({
      source: 'validation',
      message: 'update-todo-task: give one field to change or more: --title, --body, --status, --importance, --due, --start, --reminder or --categories.',
    });
    expect(todo.calls).toEqual([]);
  });

  it('passes over the task when it changed between the plan and the write', async () => {
    const todo = createSeededTodoStore();
    let reads = 0;
    const graph = {
      ...todo.graph,
      get: async (path: string) => {
        reads += 1;
        if (reads === 3) todo.edit('task-a', { title: 'Changed in between' });
        return todo.graph.get(path);
      },
    };
    const run = await runWrite(graph, 'update-todo-task', { list: 'Tasks', taskId: 'task-a', importance: 'low' });
    expect(run).toMatchObject({
      done: 0,
      skipped: [
        {
          target: '/me/todo/lists/list-a/tasks/task-a',
          reason: 'The item changed after the plan, so this run passes over it. Run the command again to see it as it is now.',
        },
      ],
    });
    expect(writes(todo.calls)).toEqual([]);
  });
});

describe('complete-todo-task', () => {
  it('marks the task completed, and the undo gives it its status of before', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'complete-todo-task', { list: 'Tasks', taskId: 'task-a' }, { journal });
    expect(run['changes']).toEqual([
      { action: 'complete-task', target: '/me/todo/lists/list-a/tasks/task-a', before: TASK_BEFORE, after: { listId: 'list-a', taskId: 'task-a', fields: 'status' } },
    ]);
    expect(todo.tasks.get('task-a')?.['status']).toBe('completed');
    expect(todo.calls.some((call) => call.includes('mailboxSettings'))).toBe(false);
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual(['undone']);
    expect(todo.tasks.get('task-a')?.['status']).toBe('notStarted');
  });

  it('passes over a task that is completed already', async () => {
    const todo = createSeededTodoStore();
    todo.edit('task-a', { status: 'completed' });
    const dry = await runWrite(todo.graph, 'complete-todo-task', { list: 'Tasks', taskId: 'task-a' }, { dryRun: true });
    expect(dry).toEqual({
      dryRun: true,
      command: 'complete-todo-task',
      changes: [],
      skipped: [{ target: '/me/todo/lists/list-a/tasks/task-a', before: { ...TASK_BEFORE, status: 'completed' }, reason: 'The task is completed already.' }],
    });
    const run = await runWrite(todo.graph, 'complete-todo-task', { list: 'Tasks', taskId: 'task-a' });
    expect(run).toEqual({ done: 0, skipped: [expect.objectContaining({ reason: 'The task is completed already.' })], failed: [], changes: [] });
  });
});
