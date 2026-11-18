import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import { ok } from '../domain/result.ts';
import type { WriteGraph } from '../infra/write-graph.ts';
import { writeCommandOf } from '../use-cases/commands/command-graph.ts';
import type { LocalFileReader } from '../use-cases/commands/read-local-file.ts';
import { readLocalFile } from '../use-cases/commands/read-local-file.ts';
import { withUnknownParamRejection } from '../use-cases/commands/reject-unknown-params.ts';
import { fakeCommand, fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { createLocalFilesFake } from '../test-helpers/local-files-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { fakeWriteGraph } from '../test-helpers/write-graph-fake.ts';
import { buildCliShell } from './cli-shell.ts';
import { buildMcpShell } from './mcp-shell.ts';
import { runRegistryCommand } from './run-registry-command.ts';

/*
 * A write command reads a local file that a param names (`--ids-file`) through
 * the file reader that the shell gives it: on the CLI, over MCP, in its plan
 * and in its execute. With no reader, the read is refused with a message that
 * names the cause, and an unknown param is still refused first.
 */

const textOf = async (files: LocalFileReader | undefined, path: string): Promise<string> => {
  const read = await readLocalFile(files, path, 1000);
  return read.ok ? new TextDecoder().decode(read.value.bytes) : `refused: ${read.error.code}`;
};

const tagWidgets = withUnknownParamRejection(
  writeCommandOf({
    schema: z.object({ idsFile: z.string() }),
    execute: async (_graph, params, files) => ok({ ids: await textOf(files, params['idsFile'] ?? '') }),
    meta: {
      ...fakeCommand().meta,
      summary: 'Tag the widgets of a file.',
      category: 'mail',
      graphMethod: 'PATCH',
      effect: 'mail-state',
      undo: 'removes the tags',
      options: [{ name: 'ids-file', key: 'idsFile', required: true, description: 'A file with one widget id on each line.' }],
    },
    plan: async (_reads, params, files) => {
      const ids = await textOf(files, params['idsFile'] ?? '');
      return ok({ changes: [{ action: 'tag-widgets', target: '/widgets', before: { ids } }], apply: async () => ok({ value: { ids }, changes: [] }) });
    },
  })
);

const PACKAGE = fakeCommandPackage({ registry: { 'tag-widgets': tagWidgets } });

const FILES = createLocalFilesFake({ 'ids.txt': 'w1\nw2\n' });

const run = async (params: Record<string, string>, files?: LocalFileReader, dryRun = false): Promise<unknown> => {
  const deps = { graph: fakeWriteGraph(), fs: createFileSystemFake(), commandPackage: PACKAGE, ...(files === undefined ? {} : { files }) };
  const result = await runRegistryCommand(deps, { name: 'tag-widgets', command: tagWidgets, params, surface: 'cli', dryRun });
  return result.ok ? result.value : result.error;
};

describe('the local file read of a write command', () => {
  it('reads the file in the execute and in the plan through the reader of the run', async () => {
    expect(await run({ idsFile: 'ids.txt' }, FILES)).toEqual({ ids: 'w1\nw2\n' });
    expect(await run({ idsFile: 'ids.txt' }, FILES, true)).toEqual({
      dryRun: true,
      command: 'tag-widgets',
      changes: [{ action: 'tag-widgets', target: '/widgets', before: { ids: 'w1\nw2\n' } }],
    });
  });

  it('is refused when the caller gives no reader', async () => {
    expect(await run({ idsFile: 'ids.txt' })).toEqual({ ids: 'refused: local_files_unavailable' });
    const read = await readLocalFile(undefined, 'ids.txt', 1000);
    expect(read.ok ? '' : read.error.message).toBe(
      'ids.txt cannot be read: the caller of this command gives it no access to local files. A library caller gives a file reader as the third argument of execute.'
    );
  });

  it('refuses an unknown param before the plan or the execute reads a file', async () => {
    expect(await run({ idsFile: 'ids.txt', idsFil: 'x' }, FILES)).toMatchObject({ code: 'unknown_parameter' });
    expect(await run({ idsFile: 'ids.txt', idsFil: 'x' }, FILES, true)).toMatchObject({ code: 'unknown_parameter' });
  });

  it('reaches the command on the CLI', async () => {
    const program = buildCliShell({
      commandPackage: PACKAGE as never,
      graph: fakeWriteGraph() as WriteGraph,
      logger: createLoggerFake(),
      fs: createFileSystemFake(),
      files: FILES,
    });
    const original = process.stdout.write.bind(process.stdout);
    let captured = '';
    process.stdout.write = (chunk: string | Uint8Array): boolean => {
      captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
      return true;
    };
    try {
      await program.parseAsync(['node', 'widget-cli', '--output', 'json', 'tag-widgets', '--ids-file', 'ids.txt']);
    } finally {
      process.stdout.write = original;
    }
    expect(JSON.parse(captured)).toEqual({ ok: true, data: { ids: 'w1\nw2\n' } });
  });

  it('reaches the command over MCP', async () => {
    const server = buildMcpShell({ commandPackage: PACKAGE as never, graph: fakeWriteGraph() as WriteGraph, fs: createFileSystemFake(), files: FILES });
    const client = new Client({ name: 'test client', version: '1.0' });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await Promise.all([client.connect(clientTransport), server.server.connect(serverTransport)]);
    const result = await client.callTool({ name: 'run-write-command', arguments: { command: 'tag-widgets', params: { 'ids-file': 'ids.txt' } } });
    expect((result as { content: ReadonlyArray<{ text: string }> }).content[0]?.text).toContain('w1\nw2');
  });
});
