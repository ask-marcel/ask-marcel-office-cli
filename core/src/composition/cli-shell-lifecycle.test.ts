import { describe, expect, it } from 'bun:test';
import type { Command } from 'commander';
import { SIGN_IN_ENTRY, fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import type { CliOutput } from './cli-shell.ts';
import { buildLifecycleCli } from './cli-shell.ts';

// The CLI of a package with no Graph command, as the auth package of a split
// (widget-auth; read is widget-cli, write is widget-write): its own lifecycle
// commands, help-json, docs and --version, and no graph.

const DESCRIPTION = 'Sign in to the widget service for the widget packages.';

const AUTH_PACKAGE = fakeCommandPackage({
  registry: {},
  description: DESCRIPTION,
  lifecycle: [{ ...SIGN_IN_ENTRY, example: '{auth-bin} sign-in' }],
  binName: 'widget-auth',
  packageName: '@widgets/auth',
  bins: { auth: 'widget-auth', read: 'widget-cli', write: 'widget-write' },
});

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

const addSignIn = (program: Command, output: CliOutput): void => {
  program
    .command('sign-in')
    .description('Sign in to the widget service.')
    .action(() => output.print({ status: 'signed in' }));
};

const run = (args: ReadonlyArray<string>): Promise<string> =>
  captureStdout(() =>
    buildLifecycleCli({
      commandPackage: AUTH_PACKAGE,
      logger: createLoggerFake(),
      fs: createFileSystemFake(),
      version: '1.0.0',
      addLifecycleCommands: addSignIn,
    }).parseAsync(['node', 'widget-auth', ...args])
  );

const flat = (text: string): string => text.replaceAll(/\s+/g, ' ');

describe('the CLI of a package with no Graph command', () => {
  it('prints the description that the package gives, and no Graph wording', async () => {
    const help = flat(await run([]));
    expect(help).toContain(DESCRIPTION);
    expect(help).not.toContain('Microsoft Graph CLI');
    expect(help).not.toContain('GET endpoints');
  });

  it('lists its own lifecycle commands with help-json and docs, and no Graph command', async () => {
    const help = await run([]);
    expect(help).toContain('Usage: widget-auth');
    for (const name of ['sign-in', 'help-json', 'docs']) expect(help).toContain(`  ${name}`);
    expect(help).not.toContain('list-widgets');
  });

  it('runs a lifecycle command that the package adds', async () => {
    expect(await run(['sign-in', '--output', 'json'])).toBe('{"ok":true,"data":{"status":"signed in"}}\n');
  });

  it('answers --version with the version of the package', async () => {
    expect(await run(['--version'])).toBe('1.0.0\n');
  });

  it('lists only the lifecycle entries in help-json, filled with the bin of the package', async () => {
    const manifest = JSON.parse(await run(['help-json'])) as { readonly package: string; readonly commands: ReadonlyArray<{ readonly name: string; readonly example: string }> };
    expect(manifest.package).toBe('@widgets/auth');
    expect(manifest.commands.map((command) => [command.name, command.example])).toEqual([['sign-in', 'widget-auth sign-in']]);
  });

  it('sends a caller of a read command to the read package', async () => {
    expect(await run(['list-drives'])).toContain('`list-drives` lives in @ask-marcel/office-read: install it and run `widget-cli list-drives`.');
  });

  it('offers the closest lifecycle command for a mistyped one', async () => {
    expect(await run(['sign-ni'])).toContain('Did you mean `sign-in`?');
  });
});
