import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { err, ok } from '../domain/result.ts';
import { fakeCommand, fakeReadPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import type { CommandPackage } from '../use-cases/commands/command-package.ts';
import { buildMcpShell } from './mcp-shell.ts';

// The MCP server of the read package of a split (widget-cli; auth is widget-auth,
// write is widget-write): its texts carry the bin placeholders, and a caller on
// this server is sent to the server that holds what it asked for.

const connect = async (commandPackage: CommandPackage = fakeReadPackage()): Promise<Client> => {
  const server = buildMcpShell({ commandPackage, graph: fakeGraphClient(), fs: createFileSystemFake(), version: '1.2.3' });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const textOf = (result: unknown): string => (result as { content: ReadonlyArray<{ text?: string }> }).content.map((c) => c.text ?? '').join('');

const call = async (name: string, args: Record<string, unknown>, commandPackage?: CommandPackage): Promise<string> =>
  textOf(await (await connect(commandPackage)).callTool({ name, arguments: args }));

const PLACEHOLDER = /\{(?:(?:auth|read|write)-)?bin\}/;

const failingWith = (code: string, message: string): CommandPackage =>
  fakeReadPackage({ registry: { 'list-widgets': { ...fakeCommand(), execute: async () => err({ type: 'auth_failed', message, code }) } } });

describe('the bins the MCP server of a package names', () => {
  it('fills the summaries of list-commands and the page of get-command-docs', async () => {
    const listed = await call('list-commands', {});
    expect(listed).toContain('List the widgets.');
    expect(listed).not.toMatch(PLACEHOLDER);
    const page = await call('get-command-docs', { command: 'list-widgets' });
    expect(page).toContain('How many widgets; `widget-auth status` shows the token first.');
    expect(page).toContain('widget-cli list-widgets --top 3');
    expect(page).not.toMatch(PLACEHOLDER);
  });

  it('fills the message a command fails with and the hint under it', async () => {
    const out = await call('run-command', { command: 'list-widgets' }, failingWith('InvalidAuthenticationToken', 'Token refused. Run `{auth-bin} login`.'));
    expect(out).toContain('error: Token refused. Run `widget-auth login`.');
    expect(out).toContain('hint: The cached access token is invalid or expired. Run `widget-auth login`');
    expect(out).not.toMatch(PLACEHOLDER);
  });
  it('pages a command answer with the next-page command of the package that holds it', async () => {
    const commandPackage = fakeReadPackage({
      registry: { 'list-widgets': { ...fakeCommand(), execute: async () => ok({ value: [{ id: 'w1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/widgets?$skip=1' }) } },
    });
    const out = await call('run-command', { command: 'list-widgets' }, commandPackage);
    expect(out).toContain("--- next: widget-cli next-page --url 'https://graph.microsoft.com/v1.0/widgets?$skip=1'");
    expect(out).not.toMatch(PLACEHOLDER);
  });

  it('fills the summaries of list-commands for one category', async () => {
    const commandPackage = fakeReadPackage({ registry: { 'list-widgets': fakeCommand({ summary: 'List the widgets of `{read-bin} list-drives`. Then more.' }) } });
    const listed = await call('list-commands', { category: 'drive' }, commandPackage);
    expect(listed).toContain('List the widgets of `widget-cli list-drives`.');
    expect(listed).not.toMatch(PLACEHOLDER);
  });
});
