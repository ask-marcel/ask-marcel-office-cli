import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { afterEach, describe, expect, it } from 'bun:test';
import { mkdtempSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createNodeFileSystem } from '@ask-marcel/office-core/infra/filesystem-node.ts';
import type { RunRecorder } from '@ask-marcel/office-core/use-cases/commands/record-run.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createRunJournalFake } from '@ask-marcel/office-core/test-helpers/run-journal-fake.ts';
import { createMailboxFake } from '../test-helpers/mailbox-fake.ts';
import type { MailboxFake } from '../test-helpers/mailbox-fake.ts';
import { buildWriteRuns } from './write-runs.ts';
import { buildWriteMcpServer } from './write-server.ts';

/*
 * The dry-run and the undo over MCP: `dryRun` on run-write-command, and the
 * tools `undo` (read-only) and `undo-apply` (destructive), on a fake mailbox.
 */

const RUN = '20261118T090507Z-0a1b2c3d';

const connect = async (mailbox: MailboxFake, runs: RunRecorder | undefined): Promise<Client> => {
  const server = buildWriteMcpServer({ graph: mailbox.graph, fs: createFileSystemFake(), version: '9.9.9', ...(runs === undefined ? {} : { runs }) });
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const recorder = (): RunRecorder => ({ journal: createRunJournalFake(), now: () => new Date('2026-11-18T09:05:07.000Z'), randomSuffix: () => '0a1b2c3d' });

const textOf = (result: unknown): string => (result as { content: ReadonlyArray<{ text?: string }> }).content.map((part) => part.text ?? '').join('');

const PARAMS = { subject: 'Agenda', 'body-content': 'Hi.', 'to-recipients': 'kim.lee@contoso.com' };

describe('the dry-run and the undo over MCP', () => {
  it('adds undo (read-only) and undo-apply (destructive) to a server with a run journal, and keeps run-write-command not destructive', async () => {
    const { tools } = await (await connect(createMailboxFake(), recorder())).listTools();
    const marks = Object.fromEntries(tools.map((tool) => [tool.name, [tool.annotations?.readOnlyHint, tool.annotations?.destructiveHint]]));
    expect(marks).toEqual({
      'list-commands': [true, undefined],
      'get-command-docs': [true, undefined],
      'run-write-command': [false, false],
      undo: [true, undefined],
      'undo-apply': [false, true],
    });
  });

  it('shows a draft on dryRun and makes nothing', async () => {
    const mailbox = createMailboxFake();
    const client = await connect(mailbox, recorder());
    const result = await client.callTool({ name: 'run-write-command', arguments: { command: 'create-mail-draft', params: PARAMS, dryRun: true } });
    expect(textOf(result)).toContain('dryRun: true');
    expect(mailbox.calls).toEqual([]);
  });

  it('makes a draft with a run id, shows its undo, and does it', async () => {
    const mailbox = createMailboxFake();
    const client = await connect(mailbox, recorder());
    expect(textOf(await client.callTool({ name: 'run-write-command', arguments: { command: 'create-mail-draft', params: PARAMS } }))).toContain(`runId: ${RUN}`);
    expect(textOf(await client.callTool({ name: 'undo', arguments: { runId: RUN } }))).toContain('status: will-undo');
    expect(mailbox.messages.has('draft-1')).toBe(true);
    expect(textOf(await client.callTool({ name: 'undo-apply', arguments: { runId: RUN } }))).toContain('status: undone');
    expect(mailbox.messages.has('draft-1')).toBe(false);
    const again = await client.callTool({ name: 'undo-apply', arguments: { runId: RUN } });
    expect([(again as { isError?: boolean }).isError, textOf(again)]).toEqual([true, expect.stringContaining('run_already_undone')]);
  });

  it('has no undo tool and no run id without a run journal', async () => {
    const client = await connect(createMailboxFake(), undefined);
    expect((await client.listTools()).tools.map((tool) => tool.name)).not.toContain('undo');
    expect(textOf(await client.callTool({ name: 'run-write-command', arguments: { command: 'create-mail-draft', params: PARAMS } }))).not.toContain('runId');
  });
});

describe('the run journal of the write bin', () => {
  const dirs: string[] = [];
  afterEach(() => {
    for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
  });

  it('keeps each run in .ask-marcel/runs under the home folder, with a random 8-hex suffix in its id', async () => {
    const home = mkdtempSync(join(tmpdir(), 'write-runs-'));
    dirs.push(home);
    const runs = buildWriteRuns({ home, disk: createNodeFileSystem() });
    const suffixes = [runs.randomSuffix(), runs.randomSuffix()];
    expect(suffixes.every((suffix) => /^[0-9a-f]{8}$/.test(suffix))).toBe(true);
    expect(suffixes[0]).not.toBe(suffixes[1]);
    expect(Math.abs(runs.now().getTime() - Date.now())).toBeLessThan(5000);
    const client = await connect(createMailboxFake(), runs);
    await client.callTool({ name: 'run-write-command', arguments: { command: 'create-mail-draft', params: PARAMS } });
    expect(readdirSync(join(home, '.ask-marcel', 'runs'))).toHaveLength(1);
  });

  it('takes the home folder and the disk of this process when none is given', () => {
    const runs = buildWriteRuns();
    expect(typeof runs.journal.load).toBe('function');
  });
});
