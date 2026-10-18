import { describe, expect, it } from 'bun:test';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import { fakeAuthManager } from '../test-helpers/auth-manager-fake.ts';
import { buildAuthCli } from './auth-program.ts';

// The CLI of the auth package `@ask-marcel/office-auth`: its six lifecycle
// commands under the bin `ask-marcel-office-auth`, and no Graph command.

const captureStdout = async (run: () => Promise<unknown>): Promise<string> => {
  const original = process.stdout.write.bind(process.stdout);
  let captured = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await run().catch(() => undefined);
  } finally {
    process.stdout.write = original;
  }
  return captured;
};

const run = (args: ReadonlyArray<string>): Promise<string> =>
  captureStdout(() =>
    buildAuthCli({ auth: fakeAuthManager(), logger: createLoggerFake(), fs: createFileSystemFake(), version: '1.0.0' }).parseAsync(['node', 'ask-marcel-office-auth', ...args])
  );

type Manifest = { readonly package: string; readonly commands: ReadonlyArray<{ readonly name: string }> };

describe('the CLI of the auth package', () => {
  it('answers --version with the version it is given', async () => {
    expect(await run(['--version'])).toBe('1.0.0\n');
  });

  it('describes the sign-in of the auth package, and no Graph command', async () => {
    const help = (await run([])).replaceAll(/\s+/g, ' ');
    expect(help).toContain('Usage: ask-marcel-office-auth');
    expect(help).toContain('Microsoft 365 sign-in for the @ask-marcel/office-read and @ask-marcel/office-write packages.');
    expect(help).not.toContain('Microsoft Graph CLI');
  });

  it('lists only the lifecycle commands in its help', async () => {
    const listed = [...(await run([])).matchAll(/^ {2}([a-z][a-z-]*)\b/gm)].map((match) => match[1]);
    expect(listed.toSorted((a, b) => (a ?? '').localeCompare(b ?? ''))).toEqual(['docs', 'help-json', 'login', 'logout', 'status']);
  });

  it('lists the six lifecycle commands of the auth package in help-json, and never the token helper', async () => {
    const manifest = JSON.parse(await run(['help-json'])) as Manifest;
    expect(manifest.package).toBe('@ask-marcel/office-auth');
    expect(manifest.commands.map((command) => command.name)).toEqual(['docs', 'help-json', 'login', 'logout', 'mcp', 'status']);
  });

  it('sends a caller of list-drives to the read package', async () => {
    expect(await run(['list-drives'])).toContain('`list-drives` lives in @ask-marcel/office-read: install it and run `ask-marcel-office-read list-drives`.');
  });

  it('names the auth bin in the report of status', async () => {
    expect(await run(['status'])).toContain('ask-marcel-office-auth login');
  });
});
