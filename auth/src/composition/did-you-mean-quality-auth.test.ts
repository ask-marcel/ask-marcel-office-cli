// The auth case of the did-you-mean suite: a mistyped flag of `login`. Moved out of
// did-you-mean-quality.test.ts with its text unchanged (package split, step 16), so the auth
// package takes this file whole; the helper is a copy of that file's.
import { describe, expect, it } from 'bun:test';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import { buildAuthCli } from './auth-program.ts';

// The error line an agent reads after a mistyped name or flag.
const errorFor = async (args: ReadonlyArray<string>): Promise<string> => {
  const cli = buildAuthCli({ auth: fakeAuthManager(), logger: createLoggerFake(), fs: createFileSystemFake() });
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
