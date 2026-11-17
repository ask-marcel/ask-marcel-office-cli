import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { buildWriteMcpServer } from './write-server.ts';

// QA 2026-11-16, P2-1: the docs of a write command name `--body-content` and
// `--to-recipients`, and the write server refused these keys.
describe('a write command run over MCP with the keys that its docs show', () => {
  it('makes the mail draft from `body-content` and `to-recipients`', async () => {
    const posts: string[] = [];
    const graph = fakeGraphClient({
      post: async (url: string) => {
        posts.push(url);
        return ok({ id: 'draft-1', subject: 'Hello' });
      },
    });
    const server = buildWriteMcpServer({ graph, fs: createFileSystemFake(), version: '9.9.9' });
    const client = new Client({ name: 'test client', version: '1.0' });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);

    const result = await client.callTool({
      name: 'run-write-command',
      arguments: { command: 'create-mail-draft', params: { subject: 'Hello', 'body-content': 'Hi', 'to-recipients': 'robin.chen@contoso.com' } },
    });

    expect((result as { isError?: boolean }).isError).not.toBe(true);
    expect(posts).toHaveLength(1);
    expect((result as { content: ReadonlyArray<{ text?: string }> }).content[0]?.text).toContain('draft-1');
  });
});
