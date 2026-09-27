import { describe, expect, it } from 'bun:test';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { createProcessRunnerFake } from '../test-helpers/process-runner-fake.ts';
import { buildCli } from './cli.ts';

const run = async (args: ReadonlyArray<string>): Promise<{ readonly out: string; readonly threw: boolean; readonly graphCalls: number }> => {
  let graphCalls = 0;
  const graph = fakeGraphClient({
    get: async () => {
      graphCalls += 1;
      return { ok: true, value: {} };
    },
  });
  const cli = buildCli({ auth: fakeAuthManager(), graph, logger: createLoggerFake(), processRunner: createProcessRunnerFake(), fs: createFileSystemFake(), version: '9.8.7' });
  const original = process.stdout.write.bind(process.stdout);
  let out = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    out += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  let threw = false;
  try {
    await cli.parseAsync(['node', 'ask-marcel-office', ...args]);
  } catch {
    threw = true;
  } finally {
    process.stdout.write = original;
  }
  return { out, threw, graphCalls };
};

describe('--version', () => {
  it('prints the version when it is the whole request, with or without an output format', async () => {
    expect((await run(['--version'])).out).toBe('9.8.7\n');
    expect((await run(['-V'])).out).toBe('9.8.7\n');
    expect((await run(['--output', 'json', '--version'])).out).toBe('9.8.7\n');
  });

  it('is an unknown flag on a command, so a mistyped --version-id is refused instead of printing the version and exiting 0', async () => {
    const r = await run(['--output', 'json', 'diff-drive-item-versions', '--drive-id', 'b!x', '--item-id', '01X', '--version', '3.0']);
    expect(r.threw).toBe(true);
    expect(r.graphCalls).toBe(0);
    const envelope = JSON.parse(r.out.trim()) as { ok: boolean; error: string; errorCode: string };
    expect(envelope).toMatchObject({ ok: false, errorCode: 'commander.unknownOption' });
    expect(envelope.error).toContain("unknown option '--version'");
    expect(envelope.error).toContain('`--version-id`');
  });

  it('suggests the command flag, not the version, for a --versio typo', async () => {
    const r = await run(['--output', 'json', 'download-drive-item-version', '--drive-id', 'b!x', '--item-id', '01X', '--versio', '3.0']);
    const envelope = JSON.parse(r.out.trim()) as { error: string };
    expect(envelope.error).toContain('Did you mean `--version-id`?');
  });
});
