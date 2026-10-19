import { describe, expect, it } from 'bun:test';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import { buildCli } from './cli.ts';

// The root fills each bin placeholder with the bin of the package that runs
// the command (single-package-names.ts): its own bin for a read command, the
// auth bin for a sign-in command and the write bin for a write command. The
// bins are pinned as literals, not read from a package.json: the Stryker
// sandbox copies no *.json.

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
    /* commander throws after exitOverride has rendered the error */
  } finally {
    process.stdout.write = original;
  }
  return captured;
};

const docsOf = (name: string): Promise<string> => {
  const cli = buildCli({ graph: fakeGraphClient(), logger: createLoggerFake(), fs: createFileSystemFake() });
  return captureStdout(() => cli.parseAsync(['node', 'ask-marcel-office', 'docs', name]));
};

describe('the bins that the root names in its docs', () => {
  it('names its own bin before a read command: `ask-marcel-office list-drives`', async () => {
    expect(await docsOf('list-drives')).toContain('ask-marcel-office list-drives');
  });

  it('names the auth bin before a sign-in command: `ask-marcel-office-auth login`', async () => {
    expect(await docsOf('login')).toContain('ask-marcel-office-auth login');
  });

  it('names the write bin before a write command: `ask-marcel-office-write create-mail-draft`', async () => {
    const page = await docsOf('create-mail-draft');
    expect(page).toContain('ask-marcel-office-write create-mail-draft');
    expect(page).not.toContain('ask-marcel-office create-mail-draft');
  });
});
