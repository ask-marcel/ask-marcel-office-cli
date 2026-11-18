import { describe, expect, it } from 'bun:test';
import type { AttachmentFile } from '@ask-marcel/office-core/domain/attachment-upload.ts';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import { recordingUploads, tasksIn, withItems } from '../../test-helpers/todo-scenarios.ts';
import { createSeededTodoStore } from '../../test-helpers/todo-store-fake.ts';
import { runWrite, statusesOf, undoWrite } from '../../test-helpers/run-write.ts';
import { TASK_FIELD_OPTIONS } from './todo-task-options.ts';
import { commands } from './write-commands.ts';

/*
 * The rules of the To Do commands that the main scenarios do not show: a
 * library call with no params, the trim of a text, a change between the plan
 * and the write, the copy of each kind of item, and the exact words of the
 * refusals and of the docs.
 */

const TASK = { list: 'Tasks', taskId: 'task-a' };

describe('a library call of a To Do command', () => {
  it('refuses the params that are missing, before any read', async () => {
    const refusals: Record<string, string> = {};
    for (const name of Object.keys(commands).filter((each) => each.includes('-todo-'))) {
      const result = await commands[name]?.execute(createSeededTodoStore().graph, {});
      refusals[name] = result?.ok === false ? result.error.message : 'run';
    }
    expect(refusals).toEqual({
      'create-todo-task': '--list is missing; --title is missing',
      'update-todo-task': '--list is missing; --task-id is missing',
      'complete-todo-task': '--list is missing; --task-id is missing',
      'move-todo-task': '--list is missing; --task-id is missing; --to is missing',
      'delete-todo-task': '--list is missing; --task-id is missing',
      'add-todo-linked-resource': '--list is missing; --task-id is missing; --display-name is missing; --application-name is missing',
      'apply-todo-ops': '--file is missing',
      'add-todo-task-step': '--list is missing; --task-id is missing; --title is missing',
      'update-todo-task-step': '--list is missing; --task-id is missing; --step-id is missing',
      'delete-todo-task-step': '--list is missing; --task-id is missing; --step-id is missing',
      'add-todo-task-attachment': '--list is missing; --task-id is missing; --file is missing',
      'remove-todo-task-attachment': '--list is missing; --task-id is missing; --attachment-id is missing',
      'create-todo-list': '--name is missing',
      'update-todo-list': '--list is missing; --name is missing',
    });
  });
});

describe('the texts that a To Do command gets', () => {
  it('trims the names and titles before they go to Graph', async () => {
    const todo = createSeededTodoStore();
    await runWrite(todo.graph, 'add-todo-linked-resource', { ...TASK, displayName: '  Spec  ', applicationName: '  Browser  ', webUrl: 'http://example.com/spec' });
    expect([...todo.links.values()][0]).toMatchObject({ displayName: 'Spec', applicationName: 'Browser', webUrl: 'http://example.com/spec' });
    await runWrite(todo.graph, 'add-todo-task-step', { ...TASK, title: '  Print  ' });
    await runWrite(todo.graph, 'update-todo-task-step', { ...TASK, stepId: 'step-a', title: '  Read  ' });
    expect([...todo.steps.values()].map((step) => step['displayName'])).toEqual(['Read', 'Print']);
    await runWrite(todo.graph, 'create-todo-list', { name: '  Errands  ' });
    await runWrite(todo.graph, 'update-todo-list', { list: 'Later', name: '  Someday  ' });
    expect([...todo.lists.values()].map((list) => list['displayName'])).toEqual(['Tasks', 'Someday', 'Errands']);
  });

  it('names the option of a list that is not there: --list or --to', async () => {
    const todo = createSeededTodoStore();
    const NO_LIST = ': no To Do list has this name or id. list-todo-task-lists gives the name and the id of each list.';
    const messageOf = async (name: string, params: Record<string, string>): Promise<unknown> =>
      ((await runWrite(todo.graph, name, params))['error'] as { message: string }).message;
    expect(await messageOf('complete-todo-task', { list: 'No list', taskId: 'task-a' })).toBe(`--list${NO_LIST}`);
    expect(await messageOf('update-todo-list', { list: 'No list', name: 'x' })).toBe(`--list${NO_LIST}`);
    expect(await messageOf('move-todo-task', { ...TASK, to: 'No list' })).toBe(`--to${NO_LIST}`);
  });
});

describe('the fields of a task', () => {
  it('turn a reminder off with none', async () => {
    const todo = createSeededTodoStore();
    todo.edit('task-a', { isReminderOn: true });
    await runWrite(todo.graph, 'update-todo-task', { ...TASK, reminder: 'none' });
    expect(todo.tasks.get('task-a')?.['isReminderOn']).toBe(false);
  });

  it('refuse a day of seven characters, and a reminder after another text', async () => {
    const result = await runWrite(createSeededTodoStore().graph, 'create-todo-task', { list: 'Tasks', title: 'A', due: '2026-11' });
    expect(result['error']).toMatchObject({ message: expect.stringMatching(/^--due must be/) });
    const twice = await runWrite(createSeededTodoStore().graph, 'create-todo-task', { list: 'Tasks', title: 'A', reminder: '2026-11-192026-11-19T09:00' });
    expect(twice['error']).toMatchObject({ message: expect.stringMatching(/^--reminder must be/) });
  });

  it('undo nothing when one of the fields that the run set changed', async () => {
    const todo = createSeededTodoStore();
    const journal = createRunJournalFake();
    await runWrite(todo.graph, 'update-todo-task', { ...TASK, title: 'B', importance: 'high' }, { journal });
    todo.edit('task-a', { importance: 'low' });
    expect(statusesOf(await undoWrite(todo.graph, journal, true))).toEqual([expect.stringMatching(/^changed: /)]);
  });

  it('show the hint of each field option', () => {
    expect(TASK_FIELD_OPTIONS.map((option) => [option.name, option.argumentHint?.kind ?? 'none'])).toEqual([
      ['title', 'none'],
      ['body', 'none'],
      ['status', 'magicValue'],
      ['importance', 'magicValue'],
      ['due', 'iso8601'],
      ['start', 'iso8601'],
      ['reminder', 'iso8601'],
      ['categories', 'none'],
    ]);
    expect(TASK_FIELD_OPTIONS.filter((option) => Object.hasOwn(option, 'argumentHint')).map((option) => option.name)).toEqual(['status', 'importance', 'due', 'start', 'reminder']);
  });
});

describe('the copy of a task', () => {
  it('keeps a checked step, each field of a link, and the content type of each file, and sends no empty field', async () => {
    const todo = createSeededTodoStore();
    await withItems(todo);
    todo.edit('task-a', { startDateTime: null });
    const uploads: AttachmentFile[] = [];
    await runWrite(recordingUploads(todo, uploads), 'move-todo-task', { ...TASK, to: 'Later' });
    const copy = tasksIn(todo, 'list-b')[0] ?? {};
    expect(Object.keys(copy).toSorted((a, b) => a.localeCompare(b))).toEqual([
      'body',
      'categories',
      'dueDateTime',
      'id',
      'importance',
      'isReminderOn',
      'lastModifiedDateTime',
      'listId',
      'status',
      'title',
    ]);
    expect([...todo.steps.values()].filter((step) => step['taskId'] === copy['id']).map((step) => [step['displayName'], step['isChecked']])).toEqual([
      ['Read the draft', false],
      ['Done step', true],
    ]);
    expect([...todo.links.values()].find((link) => link['taskId'] === copy['id'])).toMatchObject({ externalId: 'spec-1', webUrl: 'https://example.com/spec' });
    expect(uploads.map((file) => [file.name, Object.hasOwn(file, 'contentType') ? file.contentType : 'none'])).toEqual([
      ['a.txt', 'text/plain'],
      ['b.bin', 'none'],
    ]);
  });
});

describe('the answers of Graph with an id that is not a Graph id', () => {
  const BAD = 'bad id';
  const unreadableOf = async (graph: WriteGraph, name: string, params: Record<string, string>): Promise<unknown> =>
    ((await runWrite(graph, name, params))['error'] as { message: string }).message;

  it('are refused for a new task, a step and an attachment', async () => {
    const todo = createSeededTodoStore();
    const post: WriteGraph = { ...todo.graph, post: async () => ok({ id: BAD, title: 'A', displayName: 'A', isChecked: false, lastModifiedDateTime: 'x' }) };
    expect(await unreadableOf(post, 'create-todo-task', { list: 'Tasks', title: 'A' })).toBe('Graph gave the task it created in a shape that this CLI cannot read.');
    expect(await unreadableOf(post, 'add-todo-task-step', { ...TASK, title: 'A' })).toBe('Graph gave a step of a To Do task in a shape that this CLI cannot read.');
    const get: WriteGraph = {
      ...todo.graph,
      get: async (path, headers) => (path.endsWith('/attachments') ? ok({ value: [{ id: BAD, name: 'a', size: 1 }] }) : todo.graph.get(path, headers)),
    };
    expect(await unreadableOf(get, 'delete-todo-task', TASK)).toBe('Graph gave an attachment of a To Do task in a shape that this CLI cannot read.');
  });

  it('refuse an update whose answer is not an object', async () => {
    const todo = createSeededTodoStore();
    const patch: WriteGraph = { ...todo.graph, patch: async () => ok(null) };
    expect(await unreadableOf(patch, 'update-todo-task', { ...TASK, title: 'B' })).toBe('Graph gave the task it changed in a shape that this CLI cannot read.');
  });

  it('refuse an attachment id that is not a Graph id, with the option name', async () => {
    const result = await runWrite(createSeededTodoStore().graph, 'remove-todo-task-attachment', { ...TASK, attachmentId: 'a/b' });
    expect(result['error']).toMatchObject({ code: 'invalid_id', message: expect.stringMatching(/^--attachment-id: "a\/b" is not a Graph id/) });
  });

  it('refuse the commands on a task that is not there', async () => {
    const todo = createSeededTodoStore();
    const NOT_FOUND = 'The specified object was not found in the store.';
    const missing = { list: 'Tasks', taskId: 'task-z' };
    for (const [name, params] of [
      ['move-todo-task', { ...missing, to: 'Later' }],
      ['update-todo-task', { ...missing, title: 'B' }],
      ['update-todo-task-step', { ...missing, stepId: 'step-a', checked: 'true' }],
    ] as const) {
      expect([name, ((await runWrite(todo.graph, name, params))['error'] as { message: string }).message]).toEqual([name, NOT_FOUND]);
    }
    todo.failures.set('GET /me/mailboxSettings/timeZone', { type: 'api_error', status: 403, message: 'No scope.' });
    expect(await unreadableOf(todo.graph, 'update-todo-task', { ...TASK, due: '2026-12-01' })).toBe('No scope.');
  });
});
