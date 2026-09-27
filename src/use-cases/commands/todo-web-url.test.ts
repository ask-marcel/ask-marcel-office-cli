import { describe, expect, it } from 'bun:test';
import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

// The shape the To Do web app opens a task with, confirmed by hand 2026-09-27:
// the id exactly as Graph gives it, its `=` padding left raw.
const ID = 'AAMkADc3NTlhODYyLTgwODItNGVkNC05MzUxLWU5MjI4OTNlYWY1OQBGAAAAAAB336IjVATmRbEl82Y4LUTvBwDVDnn6NxJXSpSeAn2vlpdGAABNESPsAAA=';
const LINK = `https://to-do.office.com/tasks/id/${ID}/details`;

const run = async (name: string, answer: unknown, params: Record<string, string>): Promise<Result<unknown, GraphError>> => {
  const command = commands[name];
  if (!command) throw new Error(`${name} is not registered`);
  return command.execute(fakeGraphClient({ get: async () => ok(answer) }), params);
};

const page = { '@odata.context': 'ctx', value: [{ id: ID, title: 'Send the rates' }, { title: 'no id when --select leaves it out' }] };

describe('the web link of a To Do task', () => {
  for (const name of ['list-todo-tasks', 'list-incomplete-todo-tasks', 'list-todo-tasks-delta']) {
    it(`${name} gives each listed task the link the To Do app opens it with`, async () => {
      const result = await run(name, page, { todoTaskListId: 'L1' });
      if (!result.ok) throw new Error(result.error.message);
      expect(result.value).toEqual({ '@odata.context': 'ctx', value: [{ id: ID, title: 'Send the rates', webUrl: LINK }, { title: 'no id when --select leaves it out' }] });
    });
  }

  it('get-todo-task gives the task its link, and an id with other characters is encoded around its padding', async () => {
    const one = await run('get-todo-task', { id: ID, title: 'Send the rates' }, { todoTaskListId: 'L1', todoTaskId: ID });
    expect(one).toEqual(ok({ id: ID, title: 'Send the rates', webUrl: LINK }));
    const odd = await run('get-todo-task', { id: 'a/b+c==' }, { todoTaskListId: 'L1', todoTaskId: 'x' });
    expect(odd).toEqual(ok({ id: 'a/b+c==', webUrl: 'https://to-do.office.com/tasks/id/a%2Fb%2Bc==/details' }));
  });
});
