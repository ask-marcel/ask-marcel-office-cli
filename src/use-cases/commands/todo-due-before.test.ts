import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const pathOf = async (name: string, params: Record<string, string>): Promise<string> => {
  const command = commands[name];
  if (!command) throw new Error(`${name} is not registered`);
  let seen = '';
  const result = await command.execute(
    fakeGraphClient({
      get: async (path: string) => {
        seen = path;
        return ok({ value: [] });
      },
    }),
    params
  );
  if (!result.ok) throw new Error(result.error.message);
  return decodeURIComponent(seen);
};

describe('--due-before on the To Do task listings', () => {
  it('adds the due bound to the open-task filter, as a UTC instant without zone suffix', async () => {
    expect(await pathOf('list-incomplete-todo-tasks', { todoTaskListId: 'L1', dueBefore: '2026-09-26T16:30:00Z' })).toBe(
      "/me/todo/lists/L1/tasks?$filter=status ne 'completed' and dueDateTime/dateTime lt '2026-09-26T16:30:00'"
    );
    expect(await pathOf('list-incomplete-todo-tasks', { todoTaskListId: 'L1' })).toBe("/me/todo/lists/L1/tasks?$filter=status ne 'completed'");
  });

  it('filters every task by due date, and joins a caller filter with and', async () => {
    expect(await pathOf('list-todo-tasks', { todoTaskListId: 'L1', dueBefore: '2026-09-27' })).toBe(
      "/me/todo/lists/L1/tasks?$filter=dueDateTime/dateTime lt '2026-09-27T00:00:00'"
    );
    expect(await pathOf('list-todo-tasks', { todoTaskListId: 'L1', dueBefore: '2026-09-27', filter: "importance eq 'high'" })).toBe(
      "/me/todo/lists/L1/tasks?$filter=(importance eq 'high') and dueDateTime/dateTime lt '2026-09-27T00:00:00'"
    );
    expect(await pathOf('list-todo-tasks', { todoTaskListId: 'L1', filter: "importance eq 'high'" })).toBe("/me/todo/lists/L1/tasks?$filter=importance eq 'high'");
  });

  it('accepts a named day, resolved in the run zone', async () => {
    expect(await pathOf('list-incomplete-todo-tasks', { todoTaskListId: 'L1', dueBefore: 'tomorrow' })).toMatch(
      /and dueDateTime\/dateTime lt '\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}'$/
    );
  });

  it('refuses a date it cannot read, and lists the flag on both commands', async () => {
    const command = commands['list-todo-tasks'];
    if (!command) throw new Error('list-todo-tasks is not registered');
    const refused = await command.execute(fakeGraphClient(), { todoTaskListId: 'L1', dueBefore: 'someday' });
    expect(refused.ok).toBe(false);
    if (!refused.ok) expect(refused.error.type).toBe('validation_error');
    for (const name of ['list-todo-tasks', 'list-incomplete-todo-tasks']) expect(commands[name]?.meta.options.map((o) => o.name)).toContain('due-before');
  });
});
