import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createLocalFilesFake } from '@ask-marcel/office-core/test-helpers/local-files-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import type { StoredMessage } from '../test-helpers/mail-store-fake.ts';
import { createMailStoreFake } from '../test-helpers/mail-store-fake.ts';
import { RUN_ID, recorderOf } from '../test-helpers/run-write.ts';
import { buildWriteCli } from './write-program.ts';
import { buildWriteMcpServer } from './write-server.ts';

/*
 * The mail state commands on the two surfaces of the write bin. Over MCP,
 * delete-mail-draft reaches the host on run-destructive-command, the one tool
 * with the destructive mark, and the other mail state commands on
 * run-write-command, which is not marked destructive. On the CLI and over MCP,
 * --ids-file reads a local file through the reader of the bin.
 */

const SEED: ReadonlyArray<StoredMessage> = [
  { id: 'm1', parentFolderId: 'inbox-id', changeKey: 'm1-ck', isRead: false },
  { id: 'd1', parentFolderId: 'drafts-id', changeKey: 'd1-ck', isDraft: true },
];

const FILES = createLocalFilesFake({ 'ids.txt': 'm1\n' });

const connect = async (mailbox = createMailStoreFake(SEED)): Promise<Client> => {
  const server = buildWriteMcpServer({ graph: mailbox.graph, fs: createFileSystemFake(), runs: recorderOf(createRunJournalFake()), files: FILES, version: '1.0.0' });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const call = async (client: Client, name: string, args: Record<string, unknown>): Promise<{ readonly isError: boolean; readonly text: string }> => {
  const result = (await client.callTool({ name, arguments: args })) as { isError?: boolean; content: ReadonlyArray<{ text?: string }> };
  return { isError: result.isError === true, text: result.content.map((part) => part.text ?? '').join('') };
};

describe('the mail state commands over MCP', () => {
  it('serve delete-mail-draft on run-destructive-command only, marked destructive', async () => {
    const tools = (await (await connect()).listTools()).tools;
    const destructive = tools.find((tool) => tool.name === 'run-destructive-command');
    expect([destructive?.annotations?.readOnlyHint, destructive?.annotations?.destructiveHint]).toEqual([false, true]);
    expect(destructive?.description).toStartWith(
      'Run one of the 6 commands that DELETE in Microsoft 365. apply-todo-ops, delete-mail-draft, delete-todo-task, delete-todo-task-step, move-todo-task, remove-todo-task-attachment: delete or move an item'
    );
    const write = tools.find((tool) => tool.name === 'run-write-command');
    expect(write?.annotations?.destructiveHint).toBe(false);
    expect(write?.description).toContain('mark-mail-message, move-mail-message, move-mail-thread, set-mail-categories, set-mail-flag: change the state of a mail message');
    expect(write?.description).toContain(
      'The commands that delete are on run-destructive-command: apply-todo-ops, delete-mail-draft, delete-todo-task, delete-todo-task-step, move-todo-task, remove-todo-task-attachment.'
    );
  });

  it('delete a draft on run-destructive-command, and refuse it on run-write-command', async () => {
    const mailbox = createMailStoreFake(SEED);
    const client = await connect(mailbox);
    const refused = await call(client, 'run-write-command', { command: 'delete-mail-draft', params: { 'message-id': 'd1' } });
    expect(refused).toEqual({ isError: true, text: expect.stringContaining('"delete-mail-draft" deletes, so it is not available through run-write-command') });
    const dry = await call(client, 'run-destructive-command', { command: 'delete-mail-draft', params: { 'message-id': 'd1' }, dryRun: true });
    expect(dry.text).toContain('dryRun: true');
    expect(mailbox.calls.filter((entry) => entry.startsWith('POST'))).toEqual([]);
    const done = await call(client, 'run-destructive-command', { command: 'delete-mail-draft', params: { 'message-id': 'd1' } });
    expect([done.isError, done.text]).toEqual([false, expect.stringContaining(`runId: ${RUN_ID}`)]);
    expect(mailbox.calls.filter((entry) => entry.startsWith('POST'))).toEqual(['POST /me/messages/d1/move']);
  });

  it('read --ids-file through the reader of the server', async () => {
    const mailbox = createMailStoreFake(SEED);
    const done = await call(await connect(mailbox), 'run-write-command', { command: 'mark-mail-message', params: { 'ids-file': 'ids.txt', read: 'true' } });
    expect([done.isError, mailbox.messages.get('m1')?.['isRead']]).toEqual([false, true]);
  });
});

describe('the mail state commands on the CLI', () => {
  const run = async (mailbox: ReturnType<typeof createMailStoreFake>, args: ReadonlyArray<string>): Promise<unknown> => {
    const cli = buildWriteCli({ graph: mailbox.graph, logger: createLoggerFake(), fs: createFileSystemFake(), runs: recorderOf(createRunJournalFake()), files: FILES });
    const original = process.stdout.write.bind(process.stdout);
    let captured = '';
    process.stdout.write = (chunk: string | Uint8Array): boolean => {
      captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
      return true;
    };
    try {
      await cli.parseAsync(['node', 'ask-marcel-office-write', '--output', 'json', ...args]);
    } finally {
      process.stdout.write = original;
    }
    return JSON.parse(captured);
  };

  it('move the messages of --ids-file, with a dry-run first', async () => {
    const mailbox = createMailStoreFake(SEED);
    expect(await run(mailbox, ['move-mail-message', '--ids-file', 'ids.txt', '--to', 'archive', '--dry-run'])).toMatchObject({
      ok: true,
      data: { dryRun: true, changes: [{ before: { messageId: 'm1', folder: 'Inbox' } }] },
    });
    expect(await run(mailbox, ['move-mail-message', '--ids-file', 'ids.txt', '--to', 'archive'])).toMatchObject({ ok: true, data: { done: 1, runId: RUN_ID } });
  });

  it('give --dry-run to each mail state command', () => {
    const cli = buildWriteCli({ graph: createMailStoreFake().graph, logger: createLoggerFake(), fs: createFileSystemFake() });
    const names = ['delete-mail-draft', 'mark-mail-message', 'move-mail-message', 'move-mail-thread', 'set-mail-categories', 'set-mail-flag'];
    const flagged = cli.commands.filter((command) => names.includes(command.name()) && command.options.some((option) => option.long === '--dry-run'));
    expect(flagged.map((command) => command.name()).toSorted((a, b) => a.localeCompare(b))).toEqual(names);
  });
});
