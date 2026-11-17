import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { buildMcpServer } from './mcp.ts';

// QA 2026-11-16, P2-1: the docs of a read command name `--message-id`, and the
// read server refused that key with a pointer to the same key.
describe('a read command run over MCP with the keys that its docs show', () => {
  it('reads the mail message that `message-id` names', async () => {
    const urls: string[] = [];
    const graph = fakeGraphClient({
      get: async (url: string) => {
        urls.push(url);
        return ok({ id: 'AAMk-1', subject: 'Site survey' });
      },
    });
    const server = buildMcpServer({ graph, fs: createFileSystemFake(), version: '9.9.9' });
    const client = new Client({ name: 'test client', version: '1.0' });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);

    const result = await client.callTool({ name: 'run-command', arguments: { command: 'get-mail-message', params: { 'message-id': 'AAMk-1' } } });

    expect((result as { isError?: boolean }).isError).not.toBe(true);
    expect(urls).toHaveLength(1);
    expect(urls[0]).toContain('/me/messages/AAMk-1');
  });
});
