import { describe, expect, it } from 'bun:test';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { TodoStoreFake } from '../../test-helpers/todo-store-fake.ts';
import { createSeededTodoStore } from '../../test-helpers/todo-store-fake.ts';
import { RUN_ID, runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';

/*
 * The move and the delete of a task. Graph has no move of a task: the
 * command copies the task with its steps, linked resources and attachments to
 * the other list, and deletes the original only after the full copy. A copy
 * that fails is deleted again, and the original stays. The delete keeps a
 * full copy in the run journal, so its undo creates the task again.
 */

const API_ERROR = { type: 'api_error', status: 500, message: 'Graph failed.' } as const;

// The task of the seed with a linked resource and a file attachment beside its step.
const withChildren = async (todo: TodoStoreFake): Promise<void> => {
  await todo.graph.post('/me/todo/lists/list-a/tasks/task-a/linkedResources', { displayName: 'Spec', applicationName: 'Browser', webUrl: 'https://example.com/spec' });
  await todo.graph.uploadAttachment(
    { kind: 'todo-task', listId: 'list-a', taskId: 'task-a' },
    { name: 'a.txt', bytes: new TextEncoder().encode('hello'), contentType: 'text/plain' }
  );
};

const childrenOf = (todo: TodoStoreFake, taskId: string): ReadonlyArray<unknown> => [
  [...todo.steps.values()].filter((item) => item['taskId'] === taskId).map((item) => [item['displayName'], item['isChecked']]),
  [...todo.links.values()].filter((item) => item['taskId'] === taskId).map((item) => [item['displayName'], item['applicationName'], item['webUrl']]),
  [...todo.files.values()].filter((item) => item['taskId'] === taskId).map((item) => [item['name'], item['contentType'], item['contentBytes']]),
];

const FULL = [[['Read the draft', false]], [['Spec', 'Browser', 'https://example.com/spec']], [['a.txt', 'text/plain', 'aGVsbG8=']]];

const tasksIn = (todo: TodoStoreFake, listId: string): ReadonlyArray<string> => [...todo.tasks.values()].filter((task) => task['listId'] === listId).map((task) => task.id);

const newTaskIn = (todo: TodoStoreFake, listId: string): string => tasksIn(todo, listId)[0] ?? '';

describe('move-todo-task', () => {
  it('shows the task, its items and the target list on a dry-run, and changes nothing', async () => {
    const todo = createSeededTodoStore();
    await withChildren(todo);
    const dry = await runWrite(todo.graph, 'move-todo-task', { list: 'Tasks', taskId: 'task-a', to: 'later' }, { dryRun: true });
    expect(dry['changes']).toEqual([
      {
        action: 'move-task',
        target: '/me/todo/lists/list-a/tasks/task-a',
        before: {
          listId: 'list-a',
          list: 'Tasks',
          taskId: 'task-a',
          title: 'Send the report',
          status: 'notStarted',
          steps: '1',
          linkedResources: '1',
          attachments: '1',
          toListId: 'list-b',
          to: 'Later',
        },
      },
    ]);
    expect(tasksIn(todo, 'list-b')).toEqual([]);
  });

  it('copies the fields, the steps, the linked resources and the attachments, then deletes the original; the answer gives the new id', async () => {
    const todo = createSeededTodoStore();
    await withChildren(todo);
    const before = todo.tasks.get('task-a');
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'move-todo-task', { list: 'Tasks', taskId: 'task-a', to: 'list-b' }, { journal });
    const id = newTaskIn(todo, 'list-b');
    expect(run).toMatchObject({ done: 1, runId: RUN_ID, changes: [{ action: 'move-task', after: { listId: 'list-b', list: 'Later', taskId: id } }] });
    expect(todo.tasks.has('task-a')).toBe(false);
    expect(todo.tasks.get(id)).toMatchObject({
      title: before?.['title'],
      body: before?.['body'],
      categories: before?.['categories'],
      dueDateTime: before?.['dueDateTime'],
      importance: 'normal',
    });
    expect(todo.tasks.get(id)?.['createdDateTime']).toBeUndefined();
    expect(childrenOf(todo, id)).toEqual(FULL);
  });

  it('keeps the original and deletes the part copy when a part of the copy fails', async () => {
    const todo = createSeededTodoStore();
    await withChildren(todo);
    const graph = {
      ...todo.graph,
      post: async (path: string, body: unknown) => (/list-b\/tasks\/[^/]+\/linkedResources$/.test(path) ? ({ ok: false, error: API_ERROR } as const) : todo.graph.post(path, body)),
    };
    const run = await runWrite(graph, 'move-todo-task', { list: 'Tasks', taskId: 'task-a', to: 'Later' });
    expect(run['error']).toEqual({ source: 'validation', code: 'todo_copy_failed', message: 'The copy of the task failed: Graph failed. The CLI deleted the part copy.' });
    expect(tasksIn(todo, 'list-a')).toEqual(['task-a']);
    expect(tasksIn(todo, 'list-b')).toEqual([]);
  });

  it('says which part copy stays when its delete fails too', async () => {
    const todo = createSeededTodoStore();
    todo.failures.set('POST /me/todo/lists/list-b/tasks/task-1/checklistItems', API_ERROR);
    todo.failures.set('DELETE /me/todo/lists/list-b/tasks/task-1', API_ERROR);
    const run = await runWrite(todo.graph, 'move-todo-task', { list: 'Tasks', taskId: 'task-a', to: 'Later' });
    expect((run['error'] as { message: string }).message).toBe(
      'The copy of the task failed: Graph failed. The CLI could not delete the part copy (task task-1 in list list-b): delete it in To Do.'
    );
    expect(tasksIn(todo, 'list-a')).toEqual(['task-a']);
  });

  it('deletes the copy again when the delete of the original fails', async () => {
    const todo = createSeededTodoStore();
    todo.failures.set('DELETE /me/todo/lists/list-a/tasks/task-a', API_ERROR);
    const run = await runWrite(todo.graph, 'move-todo-task', { list: 'Tasks', taskId: 'task-a', to: 'Later' });
    expect(run['error']).toEqual({
      source: 'validation',
      code: 'todo_move_failed',
      message: 'The copy of the task is made, but the delete of the original failed: Graph failed. The CLI deleted the copy again, so the task stays in its list.',
    });
    expect([tasksIn(todo, 'list-a'), tasksIn(todo, 'list-b')]).toEqual([['task-a'], []]);
  });

  it('says that the task is in both lists when the copy cannot be deleted again', async () => {
    const todo = createSeededTodoStore();
    todo.failures.set('DELETE /me/todo/lists/list-a/tasks/task-a', API_ERROR);
    todo.failures.set('DELETE /me/todo/lists/list-b/tasks/task-1', API_ERROR);
    const run = await runWrite(todo.graph, 'move-todo-task', { list: 'Tasks', taskId: 'task-a', to: 'Later' });
    expect((run['error'] as { message: string }).message).toEndWith(
      'The CLI could not delete the copy (task task-1 in list list-b): the task is now in both lists. Delete one in To Do.'
    );
  });

  it('refuses a move to the list the task is in', async () => {
    const todo = createSeededTodoStore();
    expect((await runWrite(todo.graph, 'move-todo-task', { list: 'Tasks', taskId: 'task-a', to: 'list-a' }))['error']).toEqual({
      source: 'validation',
      code: 'todo_same_list',
      message: '--to: the task is in this list already.',
    });
  });

  it('moves the task back on undo, with its items, under a new id', async () => {
    const todo = createSeededTodoStore();
    await withChildren(todo);
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'move-todo-task', { list: 'Tasks', taskId: 'task-a', to: 'Later' }, { journal });
    const undo = await undoWrite(todo.graph, journal, true);
    const back = newTaskIn(todo, 'list-a');
    expect(undo.ok && undo.value.changes.map((change) => [change.status, change.now])).toEqual([['undone', { listId: 'list-a', taskId: back }]]);
    expect(tasksIn(todo, 'list-b')).toEqual([]);
    expect(childrenOf(todo, back)).toEqual(FULL);
  });

  it('does not move the task back when it changed after the run, or when its list of before is gone', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'move-todo-task', { list: 'Tasks', taskId: 'task-a', to: 'Later' }, { journal });
    const id = newTaskIn(todo, 'list-b');
    todo.edit(id, { status: 'completed' });
    expect(statusesOf(await undoWrite(todo.graph, journal, false))).toEqual([expect.stringMatching(/^changed: /)]);
    todo.lists.delete('list-a');
    expect(statusesOf(await undoWrite(todo.graph, journal, false))).toEqual([expect.stringMatching(/^missing: /)]);
    expect(tasksIn(todo, 'list-b')).toEqual([id]);
  });
});

describe('delete-todo-task', () => {
  it('shows the task and the count of its items on a dry-run, and deletes nothing', async () => {
    const todo = createSeededTodoStore();
    const dry = await runWrite(todo.graph, 'delete-todo-task', { list: 'Tasks', taskId: 'task-a' }, { dryRun: true });
    expect(dry['changes']).toEqual([
      {
        action: 'delete-task',
        target: '/me/todo/lists/list-a/tasks/task-a',
        before: { listId: 'list-a', list: 'Tasks', taskId: 'task-a', title: 'Send the report', status: 'notStarted', steps: '1', linkedResources: '0', attachments: '0' },
      },
    ]);
    expect(todo.tasks.has('task-a')).toBe(true);
  });

  it('deletes the task, and the undo creates it again with its items from the run journal', async () => {
    const todo = createSeededTodoStore();
    await withChildren(todo);
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'delete-todo-task', { list: 'Tasks', taskId: 'task-a' }, { journal });
    expect(run).toMatchObject({ done: 1, changes: [{ action: 'delete-task', after: {} }] });
    expect(todo.tasks.has('task-a')).toBe(false);
    const undo = await undoWrite(todo.graph, journal, true);
    const back = newTaskIn(todo, 'list-a');
    expect(undo.ok && undo.value.changes.map((change) => [change.status, change.now])).toEqual([['undone', { listId: 'list-a', taskId: back }]]);
    expect(todo.tasks.get(back)).toMatchObject({ title: 'Send the report', categories: ['Blue'], body: { content: 'The note', contentType: 'text' } });
    expect(childrenOf(todo, back)).toEqual(FULL);
  });

  it('does not create the task again when its list is gone', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'delete-todo-task', { list: 'Tasks', taskId: 'task-a' }, { journal });
    todo.lists.delete('list-a');
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual(['missing: The item is no longer there, so there is nothing to undo.']);
  });

  it('refuses the delete of a task with more items than one page of Graph, so the copy is never part of the task', async () => {
    const todo = createSeededTodoStore();
    const graph = {
      ...todo.graph,
      get: async (path: string) =>
        path.endsWith('/checklistItems') ? ({ ok: true, value: { value: [], '@odata.nextLink': 'https://graph.microsoft.com/next' } } as const) : todo.graph.get(path),
    };
    expect((await runWrite(graph, 'delete-todo-task', { list: 'Tasks', taskId: 'task-a' }))['error']).toEqual({
      source: 'validation',
      code: 'todo_copy_too_large',
      message: 'The task has more steps than Graph gives in one page, so this CLI does not copy it: the copy would not be complete.',
    });
  });

  it('reports the delete that Graph refused, and records no run', async () => {
    const todo = createSeededTodoStore();
    todo.failures.set('DELETE /me/todo/lists/list-a/tasks/task-a', API_ERROR);
    const journal = createRunJournalFake();
    const run = await runWrite(todo.graph, 'delete-todo-task', { list: 'Tasks', taskId: 'task-a' }, { journal });
    expect(run['error']).toMatchObject({ message: 'Graph failed.' });
    expect(todo.tasks.has('task-a')).toBe(true);
  });
});
