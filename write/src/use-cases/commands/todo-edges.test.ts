import { describe, expect, it } from 'bun:test';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import type { TodoStoreFake } from '../../test-helpers/todo-store-fake.ts';
import { createSeededTodoStore, TIME_ZONE } from '../../test-helpers/todo-store-fake.ts';
import { runWrite } from '../../test-helpers/run-write.ts';

/*
 * The edges of the To Do commands: the lookup of a list, the refusal of each
 * wrong option, the limits of a date and a time, a Graph answer that the CLI
 * cannot read, and undo data that a person edited in the run journal.
 */

const errorOf = async (graph: WriteGraph, name: string, params: Record<string, string>): Promise<unknown> => (await runWrite(graph, name, params, { dryRun: true }))['error'];

// The graph of the store, with the answer of the reads whose path starts with `prefix` replaced.
const answering = (todo: TodoStoreFake, prefix: string, answer: Result<unknown, GraphError>): WriteGraph => ({
  ...todo.graph,
  get: async (path) => (path.startsWith(prefix) ? answer : todo.graph.get(path)),
});

const CREATE = { list: 'Tasks', title: 'A' };

describe('the lookup of a To Do list', () => {
  it('finds a list by its name with no case, or by its id', async () => {
    const todo = createSeededTodoStore();
    const dry = (params: Record<string, string>): Promise<Record<string, unknown>> => runWrite(todo.graph, 'create-todo-task', params, { dryRun: true });
    expect((await dry({ list: 'LATER', title: 'A' }))['changes']).toEqual([expect.objectContaining({ target: '/me/todo/lists/list-b/tasks' })]);
    expect((await dry({ list: 'list-b', title: 'A' }))['changes']).toEqual([expect.objectContaining({ target: '/me/todo/lists/list-b/tasks' })]);
  });

  it('refuses a name that two lists have, and a list that is not there, without a quote', async () => {
    const todo = createSeededTodoStore();
    todo.lists.set('list-c', { id: 'list-c', displayName: 'later' });
    expect(await errorOf(todo.graph, 'create-todo-task', { list: 'Later', title: 'A' })).toEqual({
      source: 'validation',
      code: 'todo_list_ambiguous',
      message: '--list: 2 To Do lists have this name. Give the id of the list, as list-todo-task-lists returns it.',
    });
    expect(await errorOf(todo.graph, 'create-todo-task', { list: 'Nowhere list', title: 'A' })).toEqual({
      source: 'validation',
      code: 'todo_list_not_found',
      message: '--list: no To Do list has this name or id. list-todo-task-lists gives the name and the id of each list.',
    });
    expect(todo.calls.filter((call) => call !== 'GET /me/todo/lists?$top=200')).toEqual([]);
  });

  it('reads a list that is not in the first page by its id, and refuses an id that Graph does not know', async () => {
    const todo = createSeededTodoStore();
    const graph = answering(todo, '/me/todo/lists?', ok({ value: [], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/todo/lists?$skip=200' }));
    expect((await runWrite(graph, 'create-todo-task', { list: 'list-b', title: 'A' }, { dryRun: true }))['changes']).toEqual([
      expect.objectContaining({ before: { listId: 'list-b', list: 'Later' } }),
    ]);
    expect(await errorOf(graph, 'create-todo-task', { list: 'list-z', title: 'A' })).toMatchObject({ code: 'todo_list_not_found' });
    const malformed = {
      ...graph,
      get: async (path: string) => (path === '/me/todo/lists/Errands' ? err<GraphError>({ type: 'api_error', status: 400, message: 'Invalid id.' }) : graph.get(path)),
    };
    expect(await errorOf(malformed, 'create-todo-task', { list: 'Errands', title: 'A' })).toMatchObject({ code: 'todo_list_not_found' });
  });

  it('gives the error of Graph when the read of the lists or of a list fails', async () => {
    const todo = createSeededTodoStore();
    const failure: GraphError = { type: 'api_error', status: 503, message: 'Graph is busy.' };
    expect(await errorOf(answering(todo, '/me/todo/lists?', err(failure)), 'create-todo-task', CREATE)).toMatchObject({ message: 'Graph is busy.' });
    const second = answering(todo, '/me/todo/lists', err(failure));
    const firstPage = { ...second, get: async (path: string) => (path.startsWith('/me/todo/lists?') ? ok({ value: [] }) : second.get(path)) };
    expect(await errorOf(firstPage, 'create-todo-task', { list: 'list-b', title: 'A' })).toMatchObject({ message: 'Graph is busy.' });
  });

  it('refuses the answers of Graph that it cannot read', async () => {
    const todo = createSeededTodoStore();
    expect(await errorOf(answering(todo, '/me/todo/lists?', ok({ items: [] })), 'create-todo-task', CREATE)).toEqual({
      source: 'validation',
      message: 'Graph gave the To Do lists in a shape that this CLI cannot read.',
    });
    expect(await errorOf(answering(todo, '/me/todo/lists?', ok({ value: [{ id: 'bad id', displayName: 'Tasks' }] })), 'create-todo-task', CREATE)).toEqual({
      source: 'validation',
      message: 'Graph gave a To Do list in a shape that this CLI cannot read.',
    });
    expect(await errorOf(answering(todo, '/me/todo/lists/list-a/tasks/task-a', ok({ id: 'task-a' })), 'complete-todo-task', { list: 'Tasks', taskId: 'task-a' })).toEqual({
      source: 'validation',
      message: 'Graph gave a To Do task in a shape that this CLI cannot read.',
    });
    expect(await errorOf(answering(todo, '/me/mailboxSettings', ok({ value: '' })), 'create-todo-task', { ...CREATE, due: '2026-11-20' })).toEqual({
      source: 'validation',
      message: 'Graph gave the time zone of the mailbox in a shape that this CLI cannot read.',
    });
  });

  it('refuses a task id that is not a Graph id before any read of the task', async () => {
    const todo = createSeededTodoStore();
    expect(await errorOf(todo.graph, 'complete-todo-task', { list: 'Tasks', taskId: 'a/b' })).toEqual({
      source: 'validation',
      code: 'invalid_id',
      message: '--task-id: "a/b" is not a Graph id. An id has only letters, digits, "_", "-" and "=", as list-mail-folder-messages and get-mail-message return it.',
    });
    expect(todo.calls).toEqual(['GET /me/todo/lists?$top=200']);
  });
});

describe('the field options of a task', () => {
  it('refuse a wrong value, each with its rule', async () => {
    const todo = createSeededTodoStore();
    const refused = async (params: Record<string, string>): Promise<unknown> =>
      ((await errorOf(todo.graph, 'create-todo-task', { ...CREATE, ...params })) as { message: string }).message;
    expect(await refused({ due: '2026-02-30' })).toBe('--due must be a day in the form YYYY-MM-DD, for example 2026-11-20, or none to remove it');
    expect(await refused({ start: '20261120' })).toBe('--start must be a day in the form YYYY-MM-DD, for example 2026-11-20, or none to remove it');
    expect(await refused({ reminder: '2026-11-19T24:00' })).toBe(
      '--reminder must be a day and a time in the form YYYY-MM-DDTHH:mm, for example 2026-11-19T09:00, or none to remove it'
    );
    expect(await refused({ reminder: '2026-11-19T09:60' })).toMatch(/^--reminder must be/);
    expect(await refused({ reminder: '2026-02-30T09:00' })).toMatch(/^--reminder must be/);
    expect(await refused({ reminder: 'x2026-11-19T09:00' })).toMatch(/^--reminder must be/);
    expect(await refused({ reminder: '2026-11-19T09:001' })).toMatch(/^--reminder must be/);
    expect(await refused({ due: '2026-13-01' })).toMatch(/^--due must be/);
    expect(await refused({ due: '2026-11-31' })).toMatch(/^--due must be/);
    expect(await refused({ due: '2026-11-200' })).toMatch(/^--due must be/);
    expect(await refused({ due: 'x2026-11-20' })).toMatch(/^--due must be/);
    expect(await refused({ title: '   ' })).toBe('--title is empty');
    expect(await refused({ status: 'done' })).toMatch(/^--status /);
    expect(await refused({ importance: 'urgent' })).toMatch(/^--importance /);
  });

  it('take the first and the last minute of a day', async () => {
    const todo = createSeededTodoStore();
    await runWrite(todo.graph, 'update-todo-task', { list: 'Tasks', taskId: 'task-a', reminder: '2026-11-19T23:59' });
    expect(todo.tasks.get('task-a')?.['reminderDateTime']).toEqual({ dateTime: '2026-11-19T23:59:00', timeZone: TIME_ZONE });
    await runWrite(todo.graph, 'update-todo-task', { list: 'Tasks', taskId: 'task-a', reminder: '2026-11-19T00:00', title: '  Trimmed  ' });
    expect(todo.tasks.get('task-a')).toMatchObject({ reminderDateTime: { dateTime: '2026-11-19T00:00:00', timeZone: TIME_ZONE }, title: 'Trimmed' });
  });

  it('read the time zone only for a date or a reminder, and give its error', async () => {
    const todo = createSeededTodoStore();
    await runWrite(todo.graph, 'update-todo-task', { list: 'Tasks', taskId: 'task-a', due: 'none', start: 'none', reminder: 'none' });
    expect(todo.calls.some((call) => call.includes('mailboxSettings'))).toBe(false);
    const failure: GraphError = { type: 'api_error', status: 403, message: 'No scope.' };
    expect(await errorOf(answering(todo, '/me/mailboxSettings', err(failure)), 'create-todo-task', { ...CREATE, start: '2026-11-20' })).toMatchObject({ message: 'No scope.' });
  });
});
