import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createLocalFilesFake } from '@ask-marcel/office-core/test-helpers/local-files-fake.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { TodoStoreFake } from '../test-helpers/todo-store-fake.ts';
import { createSeededTodoStore } from '../test-helpers/todo-store-fake.ts';
import { recorderOf } from '../test-helpers/run-write.ts';
import { buildWriteMcpServer } from './write-server.ts';

/*
 * The To Do commands over MCP. The commands that delete or move a task, a step
 * or an attachment, and the batch of ops, are on run-destructive-command, the
 * tool with the destructive mark; the other To Do commands are on
 * run-write-command. The batch reads its local file through the reader of the
 * bin.
 */

const FILES = createLocalFilesFake({ 'ops.json': JSON.stringify({ ops: [{ op: 'add', list: 'Tasks', title: 'From a file' }] }) });

const connect = async (todo: TodoStoreFake): Promise<Client> => {
  const server = buildWriteMcpServer({ graph: todo.graph, fs: createFileSystemFake(), runs: recorderOf(createRunJournalFake()), files: FILES, version: '1.0.0' });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const call = async (client: Client, name: string, args: Record<string, unknown>): Promise<{ readonly isError: boolean; readonly text: string }> => {
  const result = (await client.callTool({ name, arguments: args })) as { isError?: boolean; content: ReadonlyArray<{ text?: string }> };
  return { isError: result.isError === true, text: result.content.map((part) => part.text ?? '').join('') };
};

describe('the To Do commands over MCP', () => {
  it('refuse a move on run-write-command, and move the task on run-destructive-command', async () => {
    const todo = createSeededTodoStore();
    const client = await connect(todo);
    const params = { list: 'Tasks', 'task-id': 'task-a', to: 'Later' };
    const refused = await call(client, 'run-write-command', { command: 'move-todo-task', params });
    expect(refused).toEqual({ isError: true, text: expect.stringContaining('"move-todo-task" deletes, so it is not available through run-write-command') });
    expect(todo.tasks.has('task-a')).toBe(true);
    const moved = await call(client, 'run-destructive-command', { command: 'move-todo-task', params });
    expect([moved.isError, todo.tasks.has('task-a')]).toEqual([false, false]);
  });

  it('create a task on run-write-command, as a dry-run first', async () => {
    const todo = createSeededTodoStore();
    const client = await connect(todo);
    const dry = await call(client, 'run-write-command', { command: 'create-todo-task', params: { list: 'Tasks', title: 'New' }, dryRun: true });
    expect([dry.isError, dry.text]).toEqual([false, expect.stringContaining('create-task')]);
    expect(todo.tasks.size).toBe(1);
    const run = await call(client, 'run-write-command', { command: 'create-todo-task', params: { list: 'Tasks', title: 'New' } });
    expect([run.isError, todo.tasks.size]).toEqual([false, 2]);
  });

  it('apply the ops of a local file on run-destructive-command', async () => {
    const todo = createSeededTodoStore();
    const run = await call(await connect(todo), 'run-destructive-command', { command: 'apply-todo-ops', params: { file: 'ops.json' } });
    expect([run.isError, run.text]).toEqual([false, expect.stringContaining('create-task')]);
    expect([...todo.tasks.values()].map((task) => task['title'])).toContain('From a file');
  });
});
