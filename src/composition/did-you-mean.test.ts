import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { createProcessRunnerFake } from '../test-helpers/process-runner-fake.ts';
import { commands } from '../use-cases/commands/index.ts';
import { buildCli } from './cli.ts';
import { buildMcpServer } from './mcp.ts';

const captureStdout = async (run: () => Promise<unknown>): Promise<string> => {
  const original = process.stdout.write.bind(process.stdout);
  let captured = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await run();
  } catch {
    /* commander throws after exitOverride has rendered the error */
  } finally {
    process.stdout.write = original;
  }
  return captured;
};

const runCli = (...args: ReadonlyArray<string>): Promise<string> => {
  const cli = buildCli({ auth: fakeAuthManager(), graph: fakeGraphClient(), logger: createLoggerFake(), processRunner: createProcessRunnerFake(), fs: createFileSystemFake() });
  return captureStdout(() => cli.parseAsync(['node', 'ask-marcel-office', ...args]));
};

const connect = async (): Promise<Client> => {
  const server = buildMcpServer({ auth: fakeAuthManager(), graph: fakeGraphClient(), fs: createFileSystemFake(), version: '9.9.9' });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const textOf = (result: unknown): string => (result as { content: ReadonlyArray<{ text?: string }> }).content.map((c) => c.text ?? '').join('');

describe('a mistyped name on the CLI', () => {
  it('suggests the command whose words hold the guessed ones', async () => {
    expect(await runCli('list-chat-messages')).toContain("unknown command 'list-chat-messages'. Did you mean `list-teams-chat-messages`");
  });

  it('suggests a flag of the command itself, by its words or by a typo', async () => {
    expect(await runCli('list-incomplete-todo-tasks', '--todo-task-list-id', 'L1', '--list-id', 'L1')).toContain("unknown option '--list-id'. Did you mean `--todo-task-list-id`?");
    expect(await runCli('list-mail-messages', '--selct', 'id')).toContain("unknown option '--selct'. Did you mean `--select`?");
  });

  it('suggests a global flag, and adds nothing when no flag is close', async () => {
    expect(await runCli('--ouptut', 'json', 'get-current-user')).toContain("unknown option '--ouptut'. Did you mean `--output`?");
    const out = await runCli('list-drives', '--no-such-flag');
    expect(out).toContain("unknown option '--no-such-flag'");
    expect(out).not.toContain('Did you mean');
  });

  it('suggests a command on the docs and help pages', async () => {
    expect(await runCli('docs', 'list-chat-messages')).toContain('Unknown command "list-chat-messages". Did you mean `list-teams-chat-messages`');
    expect(await runCli('help', 'list-chat-messages')).toContain('Unknown command "list-chat-messages". Did you mean `list-teams-chat-messages`');
  });
});

describe('a mistyped name over MCP and in the library', () => {
  it('suggests a command to run-command and get-command-docs', async () => {
    const client = await connect();
    expect(textOf(await client.callTool({ name: 'run-command', arguments: { command: 'list-chat-messages' } }))).toContain(
      'Unknown command "list-chat-messages". Did you mean `list-teams-chat-messages`'
    );
    expect(textOf(await client.callTool({ name: 'get-command-docs', arguments: { command: 'list-chat-messages' } }))).toContain(
      'Unknown command "list-chat-messages". Did you mean `list-teams-chat-messages`'
    );
  });

  it('suggests the parameter an unknown one was meant to be, per parameter when there are several', async () => {
    const command = commands['list-incomplete-todo-tasks'];
    if (!command) throw new Error('list-incomplete-todo-tasks is not registered');
    const one = await command.execute(fakeGraphClient(), { listId: 'L1' });
    expect(one.ok).toBe(false);
    if (!one.ok) expect(one.error.message).toContain('so it would have been ignored rather than applied. Did you mean `--todo-task-list-id`? Supported:');
    const two = await command.execute(fakeGraphClient(), { listId: 'L1', selct: 'id' });
    expect(two.ok).toBe(false);
    if (!two.ok)
      expect(two.error.message).toContain('ignored rather than applied. For `--list-id`, did you mean `--todo-task-list-id`? For `--selct`, did you mean `--select`? Supported:');
  });
});
