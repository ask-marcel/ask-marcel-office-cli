import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { err, ok } from '../domain/result.ts';
import { fakeCommand, fakeCommandPackage, fakeReadPackage } from '../test-helpers/command-package-fake.ts';
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

describe('a caller sent to the server that holds what it asked for', () => {
  it('tells a caller of run-command that a write command lives on the write server, and to register it', async () => {
    expect(await call('run-command', { command: 'create-mail-draft' })).toContain(
      'error: Unknown command "create-mail-draft". `create-mail-draft` lives in @ask-marcel/office-write: register its MCP server (`widget-write mcp`) and call it there. Call list-commands to see every command.'
    );
  });

  it('sends the docs of an auth command to the auth server', async () => {
    expect(await call('get-command-docs', { command: 'status' })).toContain(
      '`status` lives in @ask-marcel/office-auth: register its MCP server (`widget-auth mcp`) and call its `status` tool there.'
    );
  });

  it('sends a caller of logout to a terminal, since the auth server has no logout tool', async () => {
    const out = await call('run-command', { command: 'logout' });
    expect(out).toContain('`logout` lives in @ask-marcel/office-auth and has no MCP tool: run `widget-auth logout` in a terminal.');
    expect(out).not.toContain('register');
  });

  it('tells a caller of token that it is the token helper, never a tool to call', async () => {
    const out = await call('get-command-docs', { command: 'token' });
    expect(out).toContain('`token` is the token helper of @ask-marcel/office-auth: the read and write packages run `widget-auth token` themselves.');
    expect(out).not.toContain('register');
  });

  it('names the status tool of the auth server as what replaced scopes-check', async () => {
    expect(await call('run-command', { command: 'scopes-check' })).toContain(
      'error: Unknown command "scopes-check". The `status` tool of the widget-auth server replaced `scopes-check`.'
    );
  });

  it('offers a write command for a near miss as a command of the write server', async () => {
    expect(await call('run-command', { command: 'create-mail-draf' })).toContain('Did you mean `create-mail-draft` on the widget-write server?');
  });

  it('offers only the auth tools for a near miss of an auth command', async () => {
    const out = await call('run-command', { command: 'logou' });
    expect(out).toContain('Did you mean the `login` tool of the widget-auth server?');
    expect(out).not.toContain('`logout`');
  });

  it('never offers the token helper for a near miss', async () => {
    expect(await call('run-command', { command: 'toke' })).not.toContain('token`');
  });
});

describe('a sign-in failure on a server that has no login tool', () => {
  it('names the login tool of the auth server under a sign-in failure, since this server has none', async () => {
    const out = await call('run-command', { command: 'list-widgets' }, failingWith('not_authenticated', 'Not signed in. Run `{auth-bin} login` first.'));
    expect(out).toContain('error: Not signed in. Run `widget-auth login` first. From an MCP client, call the `login` tool of the widget-auth server (`widget-auth mcp`).');
  });

  it('names it under a missing secondary token and a refused token too', async () => {
    for (const code of ['secondary_token_unavailable', 'InvalidAuthenticationToken', 'TokenExpired']) {
      expect(await call('run-command', { command: 'list-widgets' }, failingWith(code, 'No token.'))).toContain(
        'No token. From an MCP client, call the `login` tool of the widget-auth server (`widget-auth mcp`).'
      );
    }
  });

  it('adds nothing to a failure a sign-in does not cure', async () => {
    expect(await call('run-command', { command: 'list-widgets' }, failingWith('sign_in_in_progress', 'Another process is signing in.'))).toContain(
      'error: Another process is signing in.\n'
    );
  });

  it('adds nothing on the server that has the login tool itself, where auth and this package share one bin', async () => {
    const single = fakeCommandPackage({
      registry: { 'list-widgets': { ...fakeCommand(), execute: async () => err({ type: 'auth_failed', message: 'Not signed in.', code: 'not_authenticated' }) } },
    });
    expect(await call('run-command', { command: 'list-widgets' }, single)).toContain('error: Not signed in.\n');
  });
});
