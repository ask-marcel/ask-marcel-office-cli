import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { fakeCommand, fakeCommandPackage, SIGN_IN_ENTRY } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import type { CommandPackage } from '../use-cases/commands/command-package.ts';
import { buildMcpShell } from './mcp-shell.ts';

/*
 * A package of the split holds one kind of command (package split, phase 5).
 * Its MCP server offers only the run tool that it can serve, and sends a caller
 * to the server of the other package for the other kind. The lifecycle list and
 * the example of get-command-docs are the package's own.
 */

const BINS = { auth: 'widget-auth', read: 'widget-cli', write: 'widget-write' };

const LIFECYCLE = [
  { ...SIGN_IN_ENTRY, name: 'docs' },
  { ...SIGN_IN_ENTRY, name: 'help-json' },
  { ...SIGN_IN_ENTRY, name: 'mcp' },
];

const READS_ONLY = fakeCommandPackage({
  registry: {
    'list-widgets': fakeCommand(),
    'find-widgets': fakeCommand({ summary: 'Search the widgets.', category: 'meta', graphMethod: 'POST' }),
  },
  lifecycle: LIFECYCLE,
  exampleCommand: 'find-widgets',
  bins: BINS,
});

const WRITES_ONLY = fakeCommandPackage({
  registry: {
    'save-widget-draft': fakeCommand({ summary: 'Save a widget draft.', category: 'mail', graphMethod: 'POST', effect: 'draft' }),
  },
  lifecycle: LIFECYCLE,
  exampleCommand: 'save-widget-draft',
  binName: 'widget-write',
  bins: BINS,
});

type Tool = { readonly name: string; readonly description?: string; readonly inputSchema: { readonly properties?: Record<string, { readonly description?: string }> } };

const toolsOf = async (commandPackage: CommandPackage): Promise<ReadonlyArray<Tool>> => {
  const server = buildMcpShell({ commandPackage, graph: fakeGraphClient(), fs: createFileSystemFake() });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return (await client.listTools()).tools;
};

const namesOf = (tools: ReadonlyArray<Tool>): ReadonlyArray<string> => tools.map((tool) => tool.name).toSorted((a, b) => a.localeCompare(b));
const descriptionOf = (tools: ReadonlyArray<Tool>, name: string): string => tools.find((tool) => tool.name === name)?.description ?? '';

describe('the MCP server of a package that holds only reads', () => {
  it('offers list-commands, get-command-docs and run-command, and no run-write-command', async () => {
    expect(namesOf(await toolsOf(READS_ONLY))).toEqual(['get-command-docs', 'list-commands', 'run-command']);
  });

  it('sends a caller to the server of the write package for the writes', async () => {
    expect(descriptionOf(await toolsOf(READS_ONLY), 'run-command')).toContain(
      'Run any of the 2 READ commands (mail, files, calendar, chats, search, document conversion). Cannot write: the write commands are on another MCP server (`widget-write mcp`). Call get-command-docs first'
    );
  });

  it("names the run tool that exists, the package's own lifecycle commands and its example in get-command-docs", async () => {
    const tools = await toolsOf(READS_ONLY);
    const docs = descriptionOf(tools, 'get-command-docs');
    expect(docs).toContain('Call this after list-commands and before run-command — it tells you');
    expect(docs).toContain('Also covers the lifecycle commands (docs/help-json/mcp).');
    expect(tools.find((tool) => tool.name === 'get-command-docs')?.inputSchema.properties?.['command']?.description).toBe('Command name, e.g. `find-widgets`.');
  });
});

describe('the MCP server of a package that holds only writes', () => {
  it('offers list-commands, get-command-docs and run-write-command, and no run-command', async () => {
    expect(namesOf(await toolsOf(WRITES_ONLY))).toEqual(['get-command-docs', 'list-commands', 'run-write-command']);
  });

  it('sends a caller to the server of the read package for the reads', async () => {
    const description = descriptionOf(await toolsOf(WRITES_ONLY), 'run-write-command');
    expect(description).toStartWith('Run one of the 1 commands that WRITE to Microsoft 365. save-widget-draft: create or update an UNSENT mail draft');
    expect(description).toEndWith('This CLI has no other write surface. The read commands are on another MCP server (`widget-cli mcp`).');
  });

  it('names only run-write-command and the example of the package in get-command-docs', async () => {
    const tools = await toolsOf(WRITES_ONLY);
    expect(descriptionOf(tools, 'get-command-docs')).toContain('before run-write-command — it tells you');
    expect(tools.find((tool) => tool.name === 'get-command-docs')?.inputSchema.properties?.['command']?.description).toBe('Command name, e.g. `save-widget-draft`.');
  });
});

describe('the MCP server of a package that holds reads and writes', () => {
  it('names both run tools in get-command-docs', async () => {
    expect(descriptionOf(await toolsOf(fakeCommandPackage()), 'get-command-docs')).toContain('before run-command or run-write-command — it tells you');
  });

  it('offers both run tools, and each one sends a caller to the other for the other kind', async () => {
    const tools = await toolsOf(fakeCommandPackage());
    expect(namesOf(tools)).toEqual(['get-command-docs', 'list-commands', 'run-command', 'run-write-command']);
    expect(descriptionOf(tools, 'run-command')).toContain('Cannot write: the 1 write commands are rejected here and live on run-write-command. Call get-command-docs first');
    expect(descriptionOf(tools, 'run-write-command')).toEndWith('This CLI has no other write surface. Read commands are rejected here; use run-command.');
  });
});
