import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { fakeWriteGraph } from '@ask-marcel/office-core/test-helpers/write-graph-fake.ts';
import { commands } from '../use-cases/commands/write-commands.ts';
import { buildWriteMcpServer } from './write-server.ts';

// The MCP server of the write package: the core gateway on the write graph,
// with the write package. It takes over the write halves of mcp.test.ts. The
// SDK itself answers a call of a tool that the server does not have (-32602),
// before any code of the server runs (LESSONS 2026-10-18), so no case expects
// a pointer from a call of run-command here.

const connect = async (graph: WriteGraph = fakeWriteGraph()): Promise<Client> => {
  const server = buildWriteMcpServer({ graph, fs: createFileSystemFake(), version: '1.0.0' });
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

const toolDescription = async (client: Client, name: string): Promise<string> => (await client.listTools()).tools.find((tool) => tool.name === name)?.description ?? '';

describe('the MCP server of the write package', () => {
  it('offers exactly the three gateway tools of a package of writes', async () => {
    const { tools } = await (await connect()).listTools();
    expect(tools.map((tool) => tool.name).toSorted((a, b) => a.localeCompare(b))).toEqual(['get-command-docs', 'list-commands', 'run-write-command']);
  });

  it('names itself with the write bin and the version it is given', async () => {
    expect((await connect()).getServerVersion()).toEqual({ name: 'ask-marcel-office-write', version: '1.0.0' });
  });

  it('marks run-write-command as not read-only and not destructive: a draft is unsent, and the temporary upload is deleted', async () => {
    const { tools } = await (await connect()).listTools();
    const annotations = tools.find((tool) => tool.name === 'run-write-command')?.annotations;
    expect([annotations?.readOnlyHint, annotations?.destructiveHint]).toEqual([false, false]);
  });

  it('runs a draft through run-write-command on the write graph', async () => {
    const client = await connect(fakeWriteGraph({ post: async () => ok({ id: 'draft-1', subject: 'Agenda' }) }));
    const result = await client.callTool({
      name: 'run-write-command',
      arguments: { command: 'create-mail-draft', params: { subject: 'Agenda', bodyContent: 'Hi', toRecipients: 'robin.chen@contoso.com' } },
    });
    expect(isError(result)).toBe(false);
    expect(textOf(result)).toContain('draft-1');
  });

  it('refuses a read command on run-write-command, and names the MCP server of the read package', async () => {
    const result = await (await connect()).callTool({ name: 'run-write-command', arguments: { command: 'get-current-user' } });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('ask-marcel-office-read mcp');
  });

  it('sends a caller of get-command-docs for a read command to the MCP server of the read package', async () => {
    const result = await (await connect()).callTool({ name: 'get-command-docs', arguments: { command: 'list-drives' } });
    expect(isError(result)).toBe(true);
    expect(textOf(result)).toContain('ask-marcel-office-read mcp');
  });

  it('tells the agent what each write leaves behind, one sentence per write class', async () => {
    const description = await toolDescription(await connect(), 'run-write-command');
    expect(description).toContain('create-forward-draft, create-mail-draft, create-reply-draft, update-mail-draft: create or update an UNSENT mail draft');
    expect(description).toContain(
      'convert-calendar-event-attachment-to-pdf, convert-group-post-attachment-to-pdf, convert-mail-attachment-to-pdf: upload the attachment to a temporary file in the `.ask-marcel-temp` folder of your OneDrive to convert it, then try to delete that file (a failed cleanup can leave it there)'
    );
    expect(description).toContain('(this CLI cannot send mail). convert-calendar-event-attachment-to-pdf, ');
    expect(description).toContain('The read commands are on another MCP server (`ask-marcel-office-read mcp`).');
  });

  it('advertises every write command of the registry on run-write-command', async () => {
    const description = await toolDescription(await connect(), 'run-write-command');
    expect(Object.keys(commands).filter((name) => !description.includes(name))).toEqual([]);
    expect(description).toContain(`Run one of the ${Object.keys(commands).length} commands that WRITE to Microsoft 365.`);
  });

  it('points get-command-docs at run-write-command only, with the write lifecycle commands and example', async () => {
    const description = await toolDescription(await connect(), 'get-command-docs');
    expect(description).toContain('before run-write-command —');
    expect(description).toContain('Also covers the lifecycle commands (docs/help-json/mcp).');
    expect(description).not.toContain('run-command ');
  });
});
