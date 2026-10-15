// The auth case of the did-you-mean suite: a mistyped flag of `login`. Moved out of
// did-you-mean-quality.test.ts with its text unchanged (package split, step 16), so the auth
// package takes this file whole; the helper is a copy of that file's.
import { describe, expect, it } from 'bun:test';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { buildCli } from './cli.ts';

// The error line an agent reads after a mistyped name or flag.
const errorFor = async (args: ReadonlyArray<string>): Promise<string> => {
  const cli = buildCli({ auth: fakeAuthManager(), graph: fakeGraphClient(), logger: createLoggerFake(), fs: createFileSystemFake() });
  const original = process.stdout.write.bind(process.stdout);
  let out = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    out += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await cli.parseAsync(['node', 'ask-marcel-office', '--output', 'json', ...args]);
  } catch {
    /* commander exits on a parse error */
  } finally {
    process.stdout.write = original;
  }
  return (JSON.parse(out.trim()) as { error: string }).error;
};

describe('the did-you-mean line after a mistyped name or flag', () => {
  it('offers a lifecycle command its own flags', async () => {
    expect(await errorFor(['help-json', '--ters'])).toContain('Did you mean `--terse`?');
    expect(await errorFor(['login', '--forse'])).toContain('Did you mean `--force`?');
  });
});
