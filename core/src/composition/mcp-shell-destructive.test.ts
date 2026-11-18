import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import type { CommandPackage } from '../use-cases/commands/command-package.ts';
import { describeWrites, WRITE_EFFECTS } from '../use-cases/commands/command-effect.ts';
import { fakeCommand, fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { buildMcpShell } from './mcp-shell.ts';

/*
 * A command that deletes reaches the MCP host with the destructive mark, on a
 * tool of its own (`run-destructive-command`), and the other writes keep a
 * tool that is not marked destructive. Each run tool refuses a command of
 * another tool and names the right one.
 */

const WIDGETS = fakeCommandPackage({
  registry: {
    'list-widgets': fakeCommand(),
    'save-widget-draft': fakeCommand({ summary: 'Save a widget draft.', category: 'mail', graphMethod: 'POST', effect: 'draft' }),
    'drop-widget-draft': fakeCommand({ summary: 'Drop a widget draft.', category: 'mail', graphMethod: 'POST', effect: 'delete' }),
  },
});

const connect = async (commandPackage: CommandPackage = WIDGETS): Promise<Client> => {
  const server = buildMcpShell({ commandPackage, graph: fakeGraphClient(), fs: createFileSystemFake() });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

type Tool = Awaited<ReturnType<Client['listTools']>>['tools'][number];

const toolsOf = async (client: Client): Promise<ReadonlyArray<Tool>> => (await client.listTools()).tools;

const descriptionOf = (tools: ReadonlyArray<Tool>, name: string): string => tools.find((tool) => tool.name === name)?.description ?? '';

const call = async (client: Client, tool: string, command: string): Promise<{ readonly isError: boolean; readonly text: string }> => {
  const result = (await client.callTool({ name: tool, arguments: { command, params: {} } })) as { isError?: boolean; content: ReadonlyArray<{ text: string }> };
  return { isError: result.isError === true, text: result.content.map((part) => part.text).join('') };
};

describe('the destructive write class', () => {
  it('is destructive, and says what a delete does', () => {
    expect(WRITE_EFFECTS.delete.destructive).toBe(true);
    expect(describeWrites({ 'drop-widget-draft': fakeCommand({ effect: 'delete' }) })).toBe(
      'drop-widget-draft: delete or move an item: a mail draft goes to Deleted Items, from where the undo moves it back to Drafts. A To Do task, step or attachment is deleted (a To Do move deletes the original after its copy), and the undo makes it again from the run journal, with a new id'
    );
  });
});

describe('the MCP tool of the commands that delete', () => {
  it('is offered beside run-write-command, marked destructive, while run-write-command is not', async () => {
    const tools = await toolsOf(await connect());
    const marks = Object.fromEntries(tools.map((tool) => [tool.name, [tool.annotations?.readOnlyHint, tool.annotations?.destructiveHint]]));
    expect(marks).toEqual({
      'list-commands': [true, undefined],
      'get-command-docs': [true, undefined],
      'run-command': [true, undefined],
      'run-write-command': [false, false],
      'run-destructive-command': [false, true],
    });
  });

  it('is not offered when no command deletes', async () => {
    expect((await toolsOf(await connect(fakeCommandPackage()))).map((tool) => tool.name)).not.toContain('run-destructive-command');
  });

  it('names its commands, and run-write-command sends a delete there', async () => {
    const tools = await toolsOf(await connect());
    expect(descriptionOf(tools, 'run-destructive-command')).toBe(
      'Run one of the 1 commands that DELETE in Microsoft 365. drop-widget-draft: delete or move an item: a mail draft goes to Deleted Items, from where the undo moves it back to Drafts. A To Do task, step or attachment is deleted (a To Do move deletes the original after its copy), and the undo makes it again from the run journal, with a new id. Each one takes dryRun, and each run has an undo. The other writes are on run-write-command, and the reads on run-command.'
    );
    expect(descriptionOf(tools, 'run-write-command')).toBe(
      'Run one of the 1 commands that WRITE to Microsoft 365. save-widget-draft: create or update an UNSENT mail draft (this CLI cannot send mail). The commands that delete are on run-destructive-command: drop-widget-draft. Read commands are rejected here; use run-command.'
    );
    expect(descriptionOf(tools, 'get-command-docs')).toContain('before run-command or run-write-command or run-destructive-command — it tells you');
    expect(descriptionOf(tools, 'list-commands')).toContain('for all 3 commands');
  });

  it('names the write classes in the order of the registry, not in the order of the names', async () => {
    const registry = {
      'save-widget-draft': fakeCommand({ category: 'mail', effect: 'draft' }),
      'convert-widget-to-pdf': fakeCommand({ category: 'mail', effect: 'transient-upload' }),
      'drop-widget-draft': fakeCommand({ category: 'mail', effect: 'delete' }),
    };
    const description = descriptionOf(await toolsOf(await connect(fakeCommandPackage({ registry }))), 'run-write-command');
    expect(description.indexOf('save-widget-draft:')).toBeLessThan(description.indexOf('convert-widget-to-pdf:'));
  });

  it('runs a command that deletes', async () => {
    expect(await call(await connect(), 'run-destructive-command', 'drop-widget-draft')).toEqual({ isError: false, text: expect.stringContaining('ranWith') });
  });

  it('refuses a command that deletes on run-write-command and on run-command, and names run-destructive-command', async () => {
    const client = await connect();
    expect(await call(client, 'run-write-command', 'drop-widget-draft')).toEqual({
      isError: true,
      text: expect.stringContaining(
        'error: "drop-widget-draft" deletes, so it is not available through run-write-command (which is not declared destructive). Use run-destructive-command instead. run-destructive-command accepts: drop-widget-draft.'
      ),
    });
    expect(await call(client, 'run-command', 'drop-widget-draft')).toEqual({
      isError: true,
      text: expect.stringContaining(
        'error: "drop-widget-draft" writes to Microsoft 365, so it is not available through run-command (which is declared read-only). Use run-destructive-command instead. Write commands: drop-widget-draft, save-widget-draft.'
      ),
    });
  });

  it('refuses a read and a write that does not delete, and names their tools', async () => {
    const client = await connect();
    expect(await call(client, 'run-destructive-command', 'list-widgets')).toEqual({
      isError: true,
      text: expect.stringContaining('error: "list-widgets" is a read command — use run-command. run-destructive-command only accepts: drop-widget-draft.'),
    });
    expect(await call(client, 'run-destructive-command', 'save-widget-draft')).toEqual({
      isError: true,
      text: expect.stringContaining('error: "save-widget-draft" does not delete — use run-write-command. run-destructive-command only accepts: drop-widget-draft.'),
    });
    expect(await call(client, 'run-write-command', 'list-widgets')).toEqual({
      isError: true,
      text: expect.stringContaining('error: "list-widgets" is a read command — use run-command. run-write-command only accepts: save-widget-draft.'),
    });
  });

  it('refuses an unknown command and lists every command', async () => {
    expect(await call(await connect(), 'run-destructive-command', 'drop-widgets')).toMatchObject({ isError: true, text: expect.stringContaining('drop-widget-draft') });
  });
});
