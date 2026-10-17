import { describe, expect, it } from 'bun:test';
import { fakeCommandPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import type { CliShellDeps } from './cli-shell.ts';
import { buildCliShell } from './cli-shell.ts';

// The CLI builder is core: it builds the CLI of whatever package it is given.
// These tests give it a made-up package, so a word of the single package in the
// output means the shell still reads it from somewhere.

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

const runWidgetCli = (args: ReadonlyArray<string>, over: Partial<CliShellDeps> = {}): Promise<string> => {
  const cli = buildCliShell({ commandPackage: fakeCommandPackage(), graph: fakeGraphClient(), logger: createLoggerFake(), fs: createFileSystemFake(), version: '1.2.3', ...over });
  return captureStdout(() => cli.parseAsync(['node', 'widget-cli', ...args]));
};

const flat = (text: string): string => text.replaceAll(/\s+/g, ' ');

// `<command> --help` ends in process.exit, so the help of one command is printed
// the way commander prints it, without the exit.
const widgetCommandHelp = (name: string): Promise<string> => {
  const cli = buildCliShell({ commandPackage: fakeCommandPackage(), graph: fakeGraphClient(), logger: createLoggerFake(), fs: createFileSystemFake() });
  const command = cli.commands.find((candidate) => candidate.name() === name);
  return captureStdout(async () => command?.outputHelp());
};

describe('the CLI of a package', () => {
  it('names the package bin in its help and lists the commands of the package only', async () => {
    const help = await runWidgetCli([]);
    expect(help).toContain('Usage: widget-cli');
    expect(help).toContain('widget-cli <command> --help');
    expect(help).toContain('list-widgets');
    expect(help).not.toContain('ask-marcel-office');
    expect(help).not.toContain('list-mail-messages');
  });

  it('counts the reads and the writes of the package in its description', async () => {
    const help = flat(await runWidgetCli([]));
    expect(help).toContain('the ONLY writes are 1 commands. save-widget-draft:');
    expect(help).toContain('3 GET endpoints + 1 search POST (find-widgets)');
  });

  it('gives help-json the package name and the lifecycle entries of the package', async () => {
    const manifest = JSON.parse(await runWidgetCli(['help-json'])) as { package: string; version: string; commands: ReadonlyArray<{ name: string }> };
    expect(manifest.package).toBe('@widgets/cli');
    expect(manifest.version).toBe('1.2.3');
    expect(manifest.commands.map((c) => c.name)).toEqual(['find-widgets', 'get-widget-file', 'get-widget-images', 'list-widgets', 'save-widget-draft', 'sign-in']);
  });

  it('prints the docs page of a lifecycle entry of the package', async () => {
    expect(await runWidgetCli(['docs', 'sign-in'])).toContain('Sign in to the widget service.');
  });

  it('answers a mistyped name with a command of the package and a remedy through its bin', async () => {
    const out = await runWidgetCli(['docs', 'list-widget']);
    expect(out).toContain('Did you mean `list-widgets`?');
    expect(out).toContain('Run `widget-cli --help` to list every command.');
  });

  it('points the help-json help at the help and the docs of the package bin', async () => {
    const help = flat(await widgetCommandHelp('help-json'));
    expect(help).toContain('well over 10× the size of `widget-cli --help`');
    expect(help).not.toContain('ask-marcel-office');
  });

  it('sends a text help-json request to the docs of the package bin', async () => {
    let failures = 0;
    const out = flat(await runWidgetCli(['--output', 'text', 'help-json'], { onCommandError: () => (failures += 1) }));
    expect(out).toContain('use `widget-cli docs <command>` instead.');
    expect(out).not.toContain('ask-marcel-office');
    expect(failures).toBe(1);
  });

  it('gives the docs help the package bin in its argument, its example and its lifecycle line', async () => {
    const help = flat(await widgetCommandHelp('docs'));
    expect(help).toContain('Command name to show docs for (run `widget-cli --help` to list every command).');
    expect(help).toContain('Example: widget-cli docs list-mail-messages');
    expect(help).toContain('Lifecycle: `widget-cli docs login`');
    expect(help).not.toContain('ask-marcel-office');
  });

  it('runs a command of the package on the graph and prints its answer', async () => {
    const out = await runWidgetCli(['--output', 'json', 'list-widgets', '--top', '3']);
    expect(JSON.parse(out)).toEqual({ ok: true, data: { ranWith: { top: '3' } } });
  });

  it('has no login when the package adds no lifecycle command', async () => {
    let failures = 0;
    const out = await runWidgetCli(['login'], { onCommandError: () => (failures += 1) });
    expect(out).toContain("unknown command 'login'");
    expect(failures).toBe(1);
  });

  it('lists the lifecycle commands a package adds under Lifecycle, and runs them through the same output', async () => {
    const deps: Partial<CliShellDeps> = {
      addLifecycleCommands: (program, output) => {
        program
          .command('sign-in')
          .description('Sign in to the widget service.')
          .action(() => output.print({ status: 'signed_in' }));
        program
          .command('sign-out')
          .description('Sign out of the widget service.')
          .action(() => output.fail('nobody is signed in', 'not_signed_in'));
      },
    };
    expect(flat(await runWidgetCli([], deps))).toMatch(/Lifecycle: sign-in Sign in to the widget service\. sign-out Sign out of the widget service\. docs /);
    expect(JSON.parse(await runWidgetCli(['--output', 'json', 'sign-in'], deps))).toEqual({ ok: true, data: { status: 'signed_in' } });
    let failures = 0;
    const out = await runWidgetCli(['--output', 'json', 'sign-out'], { ...deps, onCommandError: () => (failures += 1) });
    expect(JSON.parse(out)).toMatchObject({ ok: false, error: 'nobody is signed in', errorCode: 'not_signed_in' });
    expect(failures).toBe(1);
  });
});
