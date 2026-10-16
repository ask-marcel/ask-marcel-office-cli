import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import type { McpShellDeps } from './mcp-shell.ts';
import { buildMcpShell, okText } from './mcp-shell.ts';

// The MCP gateway is core: it serves whatever package it is given. A real
// client over the SDK's in-memory transport, a made-up package, fake ports.
const connect = async (over: Partial<McpShellDeps> = {}): Promise<Client> => {
  const server = buildMcpShell({ commandPackage: fakeCommandPackage(), graph: fakeGraphClient(), fs: createFileSystemFake(), version: '1.2.3', ...over });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const textOf = (result: unknown): string => (result as { content: ReadonlyArray<{ text?: string }> }).content.map((c) => c.text ?? '').join('');
const isError = (result: unknown): boolean => (result as { isError?: boolean }).isError === true;

describe('the MCP gateway of a package', () => {
  it('introduces itself by the package bin and version', async () => {
    const client = await connect();
    expect(client.getServerVersion()).toMatchObject({ name: 'widget-cli', version: '1.2.3' });
  });

  it('offers the four gateway tools when the package adds none', async () => {
    const { tools } = await (await connect()).listTools();
    expect(tools.map((t) => t.name).toSorted((a, b) => a.localeCompare(b))).toEqual(['get-command-docs', 'list-commands', 'run-command', 'run-write-command']);
  });

  it('serves the tools a package adds beside the gateway tools', async () => {
    const client = await connect({
      addTools: (server) => {
        server.registerTool('sign-in', { title: 'Sign in', description: 'Sign in to the widget service.' }, async () => okText('signed in'));
      },
    });
    const { tools } = await client.listTools();
    expect(tools.map((t) => t.name)).toContain('sign-in');
    expect(textOf(await client.callTool({ name: 'sign-in', arguments: {} }))).toBe('signed in');
  });

  it('counts the commands of the package in the tool descriptions', async () => {
    const { tools } = await (await connect()).listTools();
    const descriptionOf = (name: string): string => tools.find((t) => t.name === name)?.description ?? '';
    expect(descriptionOf('list-commands')).toContain('for all 5 commands');
    expect(descriptionOf('run-command')).toContain('Run any of the 4 READ commands');
    expect(descriptionOf('run-write-command')).toContain('Run one of the 1 commands that WRITE');
  });

  it('lists the commands and the lifecycle entries of the package under its name', async () => {
    const text = textOf(await (await connect()).callTool({ name: 'list-commands', arguments: {} }));
    expect(text).toContain('@widgets/cli');
    expect(text).toContain('list-widgets');
    expect(text).toContain('sign-in');
    expect(text).not.toContain('list-mail-messages');
  });

  it('renders the docs of a lifecycle entry of the package', async () => {
    const text = textOf(await (await connect()).callTool({ name: 'get-command-docs', arguments: { command: 'sign-in' } }));
    expect(text).toContain('Sign in to the widget service.');
  });

  it('runs a read command of the package', async () => {
    const result = await (await connect()).callTool({ name: 'run-command', arguments: { command: 'list-widgets', params: { top: '2' } } });
    expect(isError(result)).toBe(false);
    expect(textOf(result)).toContain('top: 2');
  });

  it('refuses a write on run-command and names the write commands of the package', async () => {
    const result = await (await connect()).callTool({ name: 'run-command', arguments: { command: 'save-widget-draft' } });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('Write commands: save-widget-draft.');
  });

  it('answers an unknown name with a command of the package', async () => {
    const result = await (await connect()).callTool({ name: 'run-command', arguments: { command: 'list-widget' } });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('Did you mean `list-widgets`?');
  });
});
