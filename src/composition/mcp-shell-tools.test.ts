import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { fakeCommand, fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import type { CommandPackage } from '../use-cases/commands/command-package.ts';
import { buildMcpShell } from './mcp-shell.ts';

/*
 * The gateway tools on a made-up package (core, package split phase 3): each
 * kind of command runs on its own tool only, and a refusal names the commands
 * the right tool takes.
 */

const connect = async (commandPackage: CommandPackage): Promise<Client> => {
  const server = buildMcpShell({ commandPackage, graph: fakeGraphClient(), fs: createFileSystemFake() });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const textOf = (result: unknown): string => (result as { content: ReadonlyArray<{ text?: string }> }).content.map((c) => c.text ?? '').join('');
const isError = (result: unknown): boolean => (result as { isError?: boolean }).isError === true;

// Two writes, so the write list has an order to keep.
const twoWrites = (): CommandPackage =>
  fakeCommandPackage({
    registry: {
      'list-widgets': fakeCommand(),
      'save-widget-draft': fakeCommand({ summary: 'Save a widget draft.', category: 'mail', graphMethod: 'POST', effect: 'draft' }),
      'convert-widget-to-pdf': fakeCommand({ summary: 'Convert a widget to PDF.', category: 'mail', effect: 'transient-upload' }),
    },
  });

describe('the run tools of the gateway', () => {
  it('runs a write on run-write-command', async () => {
    const result = await (await connect(twoWrites())).callTool({ name: 'run-write-command', arguments: { command: 'save-widget-draft', params: {} } });
    expect(isError(result)).toBe(false);
    expect(textOf(result)).toBe('ranWith:\n');
  });

  it('refuses a write on run-command and names the writes of the package in alphabetical order', async () => {
    const result = await (await connect(twoWrites())).callTool({ name: 'run-command', arguments: { command: 'save-widget-draft' } });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toStartWith(
      'error: "save-widget-draft" writes to Microsoft 365, so it is not available through run-command (which is declared read-only). Use run-write-command instead. Write commands: convert-widget-to-pdf, save-widget-draft.'
    );
  });

  it('refuses a read on run-write-command and sends it to run-command', async () => {
    const result = await (await connect(twoWrites())).callTool({ name: 'run-write-command', arguments: { command: 'list-widgets' } });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toStartWith('error: "list-widgets" is a read command — use run-command. run-write-command only accepts: convert-widget-to-pdf, save-widget-draft.');
  });

  it('refuses an unknown category on list-commands, naming the categories', async () => {
    const result = await (await connect(twoWrites())).callTool({ name: 'list-commands', arguments: { category: 'nope' } });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toStartWith(
      'error: Unknown category "nope". Available categories: calendar, chats, drive, excel, lifecycle, mail, meta, notes, sharepoint, tasks, teams, user.'
    );
  });
});
