// The write cases of the MCP suite: a draft and a PDF conversion run through the
// run-write-command tool. Moved out of mcp.test.ts with their text unchanged (package
// split, step 16), so the write package takes this file whole; the helpers are copies of
// that file's.
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphClient } from '../infra/graph-client.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { buildMcpServer } from './mcp.ts';
import type { BuildMcpServerDeps } from './mcp.ts';

/**
 * Drives the REAL MCP protocol: a real Client and a real McpServer joined by
 * the SDK's in-memory transport pair. Only the secondary ports (Graph, auth,
 * filesystem) are fakes. So these tests exercise actual tool registration,
 * actual JSON-Schema input validation, and actual call dispatch — not our
 * assumption of them.
 */
const connect = async (overrides: Partial<BuildMcpServerDeps> = {}): Promise<Client> => {
  const server = buildMcpServer({
    auth: fakeAuthManager(),
    graph: fakeGraphClient(),
    fs: createFileSystemFake(),
    version: '9.9.9',
    ...overrides,
  });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const textOf = (result: unknown): string => {
  const content = (result as { content: ReadonlyArray<{ type: string; text?: string }> }).content;
  return content.map((c) => c.text ?? '').join('');
};
const isError = (result: unknown): boolean => (result as { isError?: boolean }).isError === true;

describe('the read/write boundary the readOnlyHint promises', () => {
  it('runs a write command through run-write-command', async () => {
    const graph: GraphClient = fakeGraphClient({ post: async () => ok({ id: 'draft-1', subject: 'Hello' }) });
    const client = await connect({ graph });
    const result = await client.callTool({
      name: 'run-write-command',
      arguments: { command: 'create-mail-draft', params: { subject: 'Hello', bodyContent: 'Hi', toRecipients: 'robin.chen@contoso.com' } },
    });
    expect(isError(result)).toBe(false);
    expect(textOf(result)).toContain('draft-1');
  });

  it('runs a PDF conversion through run-write-command', async () => {
    const graph: GraphClient = fakeGraphClient({
      get: async () => ok({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'report.pdf', contentBytes: btoa('%PDF-1.7') }),
    });
    const client = await connect({ graph });
    const result = await client.callTool({ name: 'run-write-command', arguments: { command: 'convert-mail-attachment-to-pdf', params: { messageId: 'm1', attachmentId: 'a1' } } });
    expect(isError(result)).toBe(false);
    expect(textOf(result)).toContain('application/pdf');
  });
});
