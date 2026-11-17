import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { ok } from '../domain/result.ts';
import { fakeCommand, fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import type { Command } from '../use-cases/commands/command-types.ts';
import { buildMcpShell } from './mcp-shell.ts';

/*
 * The `params` of run-command: the docs and the errors name a parameter as a
 * flag (`--shelf-id`), so an agent sends that form. QA 2026-11-16 (P2-1) found
 * that only the camelCase key ran, and that a number value was refused.
 */

type Seen = Array<Record<string, string>>;

// A command that records the params it gets, so a test sees the keys and the values that arrive.
const recording = (seen: Seen): Command => ({
  ...fakeCommand(),
  execute: async (_graph, params) => {
    seen.push(params);
    return ok({ done: true });
  },
});

const connect = async (seen: Seen): Promise<Client> => {
  const commandPackage = fakeCommandPackage({ registry: { 'list-widgets': recording(seen) } });
  const server = buildMcpShell({ commandPackage, graph: fakeGraphClient(), fs: createFileSystemFake(), version: '1.2.3' });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const runWith = async (params: Record<string, unknown>): Promise<{ seen: Seen; result: unknown }> => {
  const seen: Seen = [];
  const client = await connect(seen);
  const result = await client.callTool({ name: 'run-command', arguments: { command: 'list-widgets', params } });
  return { seen, result };
};

const textOf = (result: unknown): string => (result as { content: ReadonlyArray<{ text?: string }> }).content.map((c) => c.text ?? '').join('');
const isError = (result: unknown): boolean => (result as { isError?: boolean }).isError === true;

describe('the keys of the params an MCP client sends', () => {
  it('reach the command when they have the flag form that the docs show', async () => {
    const { seen } = await runWith({ 'shelf-id': 's1' });
    expect(seen).toEqual([{ shelfId: 's1' }]);
  });

  it('reach the command when they have the `--` prefix that the error messages show', async () => {
    const { seen } = await runWith({ '--shelf-id': 's1', '--top': '2' });
    expect(seen).toEqual([{ shelfId: 's1', top: '2' }]);
  });

  it('reach the command unchanged when they are camelCase', async () => {
    const { seen } = await runWith({ shelfId: 's1' });
    expect(seen).toEqual([{ shelfId: 's1' }]);
  });

  it('are refused before the command runs when two of them are one parameter', async () => {
    const { seen, result } = await runWith({ 'shelf-id': 's1', shelfId: 's2' });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('`shelf-id` and `shelfId` are the same parameter. Give it one time.');
    expect(textOf(result)).toContain('duplicate_parameter');
    expect(seen).toEqual([]);
  });
});

describe('the values of the params an MCP client sends', () => {
  it('reach the command as strings when they are a number or a boolean', async () => {
    const { seen } = await runWith({ top: 2, deep: true, name: 'bolt' });
    expect(seen).toEqual([{ top: '2', deep: 'true', name: 'bolt' }]);
  });
});

describe('the description of the params', () => {
  it('gives an example in the flag form, names the value types, and does not name a removed alias', async () => {
    const { tools } = await (await connect([])).listTools();
    const params: { description?: string } | undefined = tools.find((t) => t.name === 'run-command')?.inputSchema.properties?.['params'];
    const description = params?.description ?? '';
    expect(description).toContain('{"message-id":"<id>","top":5}');
    expect(description).toContain('a string, a number or a boolean');
    expect(description).not.toContain('folderId');
  });
});
