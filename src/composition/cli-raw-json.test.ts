import { describe, expect, it } from 'bun:test';
import { ok } from '../domain/result.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { createProcessRunnerFake } from '../test-helpers/process-runner-fake.ts';
import { buildCli } from './cli.ts';

const captureStdout = async (run: () => Promise<unknown>): Promise<string> => {
  const original = process.stdout.write.bind(process.stdout);
  let captured = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await run();
  } catch {
    /* commander may exit */
  } finally {
    process.stdout.write = original;
  }
  return captured;
};

const cliServing = (body: unknown): ReturnType<typeof buildCli> =>
  buildCli({
    auth: fakeAuthManager(),
    graph: fakeGraphClient({ get: async () => ok(body) }),
    logger: createLoggerFake(),
    processRunner: createProcessRunnerFake(),
    fs: createFileSystemFake(),
  });

describe('--output raw-json', () => {
  it('prints the command payload alone', async () => {
    const out = await captureStdout(() =>
      cliServing({ value: [{ id: 'd1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/drives?$skiptoken=x' }).parseAsync([
        'node',
        'ask-marcel-office',
        '--output',
        'raw-json',
        'list-drives',
      ])
    );
    expect(out).toBe('{"value":[{"id":"d1"}]}\n');
  });

  it('names raw-json among the allowed formats when the format is unknown', async () => {
    const out = await captureStdout(() => cliServing({}).parseAsync(['node', 'ask-marcel-office', '--output', 'yaml', 'list-drives']));
    expect(out).toContain('Allowed choices are text, json, raw-json.');
  });
});
