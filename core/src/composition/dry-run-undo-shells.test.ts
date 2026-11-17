import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { describe, expect, it } from 'bun:test';
import { Command as Program } from 'commander';
import { z } from 'zod';
import { err, ok } from '../domain/result.ts';
import type { WriteGraph } from '../infra/write-graph.ts';
import { writeCommandOf } from '../use-cases/commands/command-graph.ts';
import type { Command } from '../use-cases/commands/command-types.ts';
import { buildManifest, renderSingleCommand } from '../use-cases/commands/docs-builders.ts';
import type { RunRecorder } from '../use-cases/commands/record-run.ts';
import type { UndoHandlers } from '../use-cases/commands/undo-run.ts';
import type { PlanWrite } from '../use-cases/commands/write-plan.ts';
import { fakeCommand, fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { createRunJournalFake } from '../test-helpers/run-journal-fake.ts';
import { fakeWriteGraph } from '../test-helpers/write-graph-fake.ts';
import { buildCliShell } from './cli-shell.ts';
import type { CliOutput } from './cli-shell.ts';
import { buildMcpShell } from './mcp-shell.ts';
import { addUndoCommand, addUndoTools } from './undo-shells.ts';

/*
 * `--dry-run` and `undo` on both shells, for a package with one widget draft
 * that has a plan: the CLI flag and help, the MCP `dryRun` input, the run id
 * with a run journal, and the undo command and tools.
 */

const RUN = '20261118T090507Z-0a1b2c3d';

const plan: PlanWrite = async (reads, params) => {
  const shelf = await reads.getBinary('/shelves/1');
  if (!shelf.ok) return shelf;
  const fetched = await reads.fetchUrl('https://contoso.sharepoint.com/shelf');
  if (!fetched.ok) return fetched;
  return ok({
    changes: [{ action: 'save-widget', target: '/widgets', before: {} }],
    apply: async (graph) => {
      const made = await graph.post('/widgets', { name: params['name'] });
      if (!made.ok) return made;
      return ok({
        value: { id: 'w1' },
        changes: [{ action: 'save-widget', target: '/widgets', before: {}, after: { widgetId: 'w1' }, undo: { kind: 'drop-widget', data: { widgetId: 'w1' } } }],
      });
    },
  });
};

const saveWidget = writeCommandOf({
  schema: z.object({ name: z.string() }),
  execute: async () => ok({ id: 'w1' }),
  meta: {
    ...fakeCommand().meta,
    summary: 'Save a widget.',
    category: 'mail',
    graphMethod: 'POST',
    effect: 'draft',
    undo: 'drops the widget',
    options: [{ name: 'name', key: 'name', required: true, description: 'The name of the widget.' }],
  },
  plan,
});

const PACKAGE = fakeCommandPackage({ registry: { 'save-widget': saveWidget as Command, 'save-widget-draft': fakeCommand({ effect: 'draft', category: 'mail' }) } });

const HANDLERS: UndoHandlers = {
  'drop-widget': { check: async () => ok('ready'), apply: async (graph, data) => graph.delete(`/widgets/${String(data['widgetId'])}`) },
};

type Calls = { readonly graph: WriteGraph; readonly calls: string[] };

const loggedGraph = (): Calls => {
  const calls: string[] = [];
  const graph = fakeWriteGraph({
    post: async (path) => {
      calls.push(`POST ${path}`);
      return ok({});
    },
    delete: async (path) => {
      calls.push(`DELETE ${path}`);
      return ok({});
    },
  });
  return { graph, calls };
};

const recorder = (): RunRecorder => ({ journal: createRunJournalFake(), now: () => new Date('2026-11-18T09:05:07.000Z'), randomSuffix: () => '0a1b2c3d' });

const captureStdout = async (run: () => Promise<unknown>): Promise<string> => {
  const original = process.stdout.write.bind(process.stdout);
  let captured = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await run();
  } finally {
    process.stdout.write = original;
  }
  return captured;
};

const cli = async (graph: WriteGraph, runs: RunRecorder | undefined, ...args: string[]): Promise<string> => {
  const program = buildCliShell({ commandPackage: PACKAGE as never, graph, logger: createLoggerFake(), fs: createFileSystemFake(), ...(runs === undefined ? {} : { runs }) });
  return captureStdout(async () => program.parseAsync(['node', 'widget-cli', '--output', 'json', ...args]).catch(() => undefined));
};

describe('--dry-run on the CLI', () => {
  it('gives the flag to a command with a plan only, and its help says what the undo does', () => {
    const program = buildCliShell({ commandPackage: PACKAGE as never, graph: fakeWriteGraph(), logger: createLoggerFake(), fs: createFileSystemFake() });
    const flagged = program.commands.filter((command) => command.options.some((option) => option.long === '--dry-run')).map((command) => command.name());
    expect(flagged).toEqual(['save-widget']);
    const help = program.commands.find((command) => command.name() === 'save-widget')?.helpInformation() ?? '';
    expect(help).toContain('Show each change that the command would make');
  });

  it('shows the changes and makes none, then makes them with a run id when a journal is given', async () => {
    const { graph, calls } = loggedGraph();
    const runs = recorder();
    expect(JSON.parse(await cli(graph, runs, 'save-widget', '--name', 'Bolt', '--dry-run'))).toEqual({
      ok: true,
      data: { dryRun: true, command: 'save-widget', changes: [{ action: 'save-widget', target: '/widgets', before: {} }] },
    });
    expect(calls).toEqual([]);
    expect(JSON.parse(await cli(graph, runs, 'save-widget', '--name', 'Bolt')).data).toMatchObject({ id: 'w1', runId: RUN });
    expect(calls).toEqual(['POST /widgets']);
  });

  it('suggests --dry-run for a mistyped flag of a command with a plan', async () => {
    expect(await cli(fakeWriteGraph(), undefined, 'save-widget', '--name', 'Bolt', '--dry-rn')).toContain('--dry-run');
  });
});

const connect = async (server: McpServer): Promise<Client> => {
  const client = new Client({ name: 'test client', version: '1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
  return client;
};

const textOf = (result: unknown): string => (result as { content: ReadonlyArray<{ text?: string }> }).content.map((part) => part.text ?? '').join('');

describe('dryRun over MCP', () => {
  it('shows the changes of a command with a plan, refuses it on a command without one, and records a run with a journal', async () => {
    const { graph, calls } = loggedGraph();
    const client = await connect(buildMcpShell({ commandPackage: PACKAGE as never, graph, fs: createFileSystemFake(), runs: recorder() }));
    const shown = await client.callTool({ name: 'run-write-command', arguments: { command: 'save-widget', params: { name: 'Bolt' }, dryRun: true } });
    expect(textOf(shown)).toContain('dryRun: true');
    const refused = await client.callTool({ name: 'run-write-command', arguments: { command: 'save-widget-draft', params: {}, dryRun: true } });
    expect(textOf(refused)).toContain('dry_run_unsupported');
    expect(calls).toEqual([]);
    expect(textOf(await client.callTool({ name: 'run-write-command', arguments: { command: 'save-widget', params: { name: 'Bolt' } } }))).toContain(`runId: ${RUN}`);
  });
});

describe('the undo command and tools', () => {
  const seeded = async (): Promise<{ readonly runs: RunRecorder; readonly logged: Calls }> => {
    const logged = loggedGraph();
    const runs = recorder();
    await captureStdout(async () => cli(logged.graph, runs, 'save-widget', '--name', 'Bolt'));
    return { runs, logged };
  };

  it('shows the undo on the CLI, does it with --apply, and fails with the code of the error', async () => {
    const { runs, logged } = await seeded();
    const printed: unknown[] = [];
    const failed: Array<readonly [string, string | undefined, string | undefined]> = [];
    const output: CliOutput = { print: (data) => printed.push(data), fail: (message, code, source) => failed.push([message, code, source]) };
    const program = new Program();
    addUndoCommand(program, output, { runs, graph: logged.graph, handlers: HANDLERS });
    await program.parseAsync(['node', 'widget-cli', 'undo', RUN]);
    await program.parseAsync(['node', 'widget-cli', 'undo', RUN, '--apply']);
    await program.parseAsync(['node', 'widget-cli', 'undo', RUN, '--apply']);
    expect(printed).toMatchObject([
      { applied: false, changes: [{ status: 'will-undo' }] },
      { applied: true, changes: [{ status: 'undone' }] },
    ]);
    expect(failed).toEqual([[`Run ${RUN} is already fully undone: there is nothing left to undo.`, 'run_already_undone', 'cli']]);
    expect(logged.calls).toEqual(['POST /widgets', 'DELETE /widgets/w1']);
  });

  it('marks the undo tool read-only and the undo-apply tool destructive', async () => {
    const { runs, logged } = await seeded();
    const server = new McpServer({ name: 'widget-cli', version: '1.0.0' });
    addUndoTools(server, { runs, graph: logged.graph, handlers: HANDLERS });
    const client = await connect(server);
    const { tools } = await client.listTools();
    expect(tools.map((tool) => [tool.name, tool.annotations?.readOnlyHint, tool.annotations?.destructiveHint])).toEqual([
      ['undo', true, undefined],
      ['undo-apply', false, true],
    ]);
    expect(textOf(await client.callTool({ name: 'undo', arguments: { runId: RUN } }))).toContain('status: will-undo');
    expect(textOf(await client.callTool({ name: 'undo-apply', arguments: { runId: RUN } }))).toContain('status: undone');
    const unknown = await client.callTool({ name: 'undo', arguments: { runId: 'nope' } });
    expect([(unknown as { isError?: boolean }).isError, textOf(unknown)]).toEqual([true, expect.stringContaining('invalid_run_id')]);
  });
});

describe('the docs of a command with a dry-run and an undo', () => {
  it('carries the undo in the manifest and says on the docs page how to dry-run and undo it', () => {
    const entry = buildManifest(PACKAGE, '1.0.0').commands.find((command) => command.name === 'save-widget');
    expect(entry?.undo).toBe('drops the widget');
    const page = renderSingleCommand(PACKAGE, 'save-widget');
    expect(page.ok ? page.value : '').toContain('- **Dry-run and undo:** `--dry-run` shows each change');
    expect(page.ok ? page.value : '').toContain('the undo drops the widget.');
  });
});

describe('the reads of a plan', () => {
  it('passes on the failure of a read of bytes or of a URL', async () => {
    for (const failing of [
      { getBinary: async () => err({ type: 'network_error' as const, message: 'down' }) },
      { fetchUrl: async () => err({ type: 'network_error' as const, message: 'down' }) },
    ]) {
      const planned = await saveWidget.plan?.(fakeWriteGraph(failing), { name: 'Bolt' });
      expect(planned).toEqual(err({ type: 'network_error', message: 'down' }));
    }
  });
});
