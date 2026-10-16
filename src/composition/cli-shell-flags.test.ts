import { describe, expect, it } from 'bun:test';
import { fakeCommand, fakeCommandPackage, fakeReadPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import type { CommandPackage } from '../use-cases/commands/command-package.ts';
import type { CliShellDeps } from './cli-shell.ts';
import { buildCliShell } from './cli-shell.ts';

/*
 * The CLI builder's own flags and lifecycle commands, on a made-up package
 * (core, package split phase 3): the global flags refuse a bad value, help and
 * help-json answer the package's commands, and a mistyped name or flag gets a
 * suggestion.
 */

type Run = { readonly out: string; readonly thrown?: string };

const captureStdout = async (run: () => Promise<unknown>): Promise<Run> => {
  const original = process.stdout.write.bind(process.stdout);
  let out = '';
  let thrown: string | undefined;
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    out += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await run().catch((e: unknown) => {
      thrown = (e as { readonly code?: string }).code;
    });
  } finally {
    process.stdout.write = original;
  }
  return thrown === undefined ? { out } : { out, thrown };
};

const runWidgetCli = (args: ReadonlyArray<string>, over: Partial<CliShellDeps> = {}): Promise<Run> => {
  const cli = buildCliShell({ commandPackage: fakeCommandPackage(), graph: fakeGraphClient(), logger: createLoggerFake(), fs: createFileSystemFake(), version: '1.2.3', ...over });
  return captureStdout(() => cli.parseAsync(['node', 'widget-cli', ...args]));
};

const commandHelp = async (commandPackage: CommandPackage, name: string): Promise<string> => {
  const cli = buildCliShell({ commandPackage, graph: fakeGraphClient(), logger: createLoggerFake(), fs: createFileSystemFake() });
  return (await captureStdout(async () => cli.commands.find((candidate) => candidate.name() === name)?.outputHelp())).out;
};

const INVALID_VALUE_HINT =
  'hint: A CLI flag value failed type validation (e.g. `--top abc` when an integer was expected). Run `widget-cli <command> --help` for the expected value shape per flag.';

describe('the global flags', () => {
  it('refuses a second --output, naming both values', async () => {
    const run = await runWidgetCli(['--output', 'json', '--output', 'text', 'list-widgets']);
    expect(JSON.parse(run.out)).toEqual(
      expect.objectContaining({
        ok: false,
        error: 'option \'--output <format>\' argument \'text\' is invalid. --output cannot be passed more than once (previous: "json", new: "text").',
        errorCode: 'commander.invalidArgument',
        source: 'cli',
      })
    );
    expect(run.thrown).toBe('commander.invalidArgument');
  });

  it('refuses an output format it does not have, naming the ones it has', async () => {
    const run = await runWidgetCli(['--output', 'yaml', 'list-widgets']);
    expect(run.out).toBe(
      `error: option '--output <format>' argument 'yaml' is invalid. Allowed choices are text, json, raw-json.\ncode: commander.invalidArgument\n${INVALID_VALUE_HINT}\nsource: cli\n`
    );
  });

  it('refuses a time zone that is not an IANA zone', async () => {
    const run = await runWidgetCli(['--tz', 'Mars/Base', 'list-widgets']);
    expect(run.out).toStartWith("error: option '--tz <zone>' argument 'Mars/Base' is invalid. Not a known IANA time zone (e.g. Europe/Amsterdam, Asia/Shanghai, UTC).\n");
  });

  it('takes an IANA time zone for the named days of a command', async () => {
    expect(JSON.parse((await runWidgetCli(['--output', 'json', '--tz', 'Asia/Shanghai', 'list-widgets'])).out)).toEqual({ ok: true, data: { ranWith: {} } });
  });

  it('prints the version alone or beside an output format, and 0.0.0 when the package gives none', async () => {
    expect((await runWidgetCli(['--version'])).out).toBe('1.2.3\n');
    expect((await runWidgetCli(['--output', 'json', '-V'], { version: undefined })).out).toBe('0.0.0\n');
    expect((await runWidgetCli(['--output=json', '--version'])).out).toBe('1.2.3\n');
  });

  it('refuses a repeated single-value flag of a command', async () => {
    const run = await runWidgetCli(['list-widgets', '--top', '1', '--top', '2']);
    expect(run.out).toStartWith(
      'error: option \'--top <value>\' argument \'2\' is invalid. --top cannot be passed more than once (previous: "1", new: "2"). Single-value flags reject duplicate occurrences.\n'
    );
    expect(run.thrown).toBe('commander.invalidArgument');
  });

  it('runs a command with its flags when the args come from a user rather than from node', async () => {
    const cli = buildCliShell({ commandPackage: fakeCommandPackage(), graph: fakeGraphClient(), logger: createLoggerFake(), fs: createFileSystemFake() });
    const run = await captureStdout(() => cli.parseAsync(['--output', 'json', 'list-widgets', '--top', '2'], { from: 'user' }));
    expect(JSON.parse(run.out)).toEqual({ ok: true, data: { ranWith: { top: '2' } } });
  });
});

describe('help and docs', () => {
  it('prints the global help for a bare help', async () => {
    expect((await runWidgetCli(['help'])).out).toStartWith('Usage: widget-cli [options] [command]\n');
  });

  it('prints the docs page of a command for help <command>', async () => {
    expect((await runWidgetCli(['help', 'list-widgets'])).out).toStartWith('# `list-widgets`\n\nDo a widget thing. Second sentence.\n');
  });

  it('answers an unknown name to help and to docs with the package remedy and a structured code', async () => {
    for (const command of ['help', 'docs']) {
      const run = await runWidgetCli([command, 'nope']);
      expect(run.out).toStartWith('error: Unknown command "nope". Run `widget-cli --help` to list every command.\ncode: cli_unknown_command\n');
    }
  });

  it('writes a docs page to --output-path and reports where', async () => {
    const fs = createFileSystemFake();
    const run = await runWidgetCli(['--output-path', 'docs/sign-in.md', 'docs', 'sign-in'], { fs });
    expect(run.out).toStartWith('contentType: text/markdown\n');
    expect(run.out).toContain('savedTo: docs/sign-in.md');
    expect(fs.snapshot('docs/sign-in.md')).toStartWith('# `sign-in`');
  });

  it('shows the stability, pagination and request body of a command in its help', async () => {
    const commandPackage = fakeCommandPackage({
      registry: { 'list-widgets': fakeCommand({ stability: 'experimental', pagination: true, bodyTemplate: '{ "requests": [] }' }) },
    });
    const help = await commandHelp(commandPackage, 'list-widgets');
    expect(help).toContain('Stability: experimental — rides a Microsoft-internal substrate');
    expect(help).toContain('Pagination: ');
    expect(help).toContain('Request body:\n  { "requests": [] }');
  });

  it('leaves the stability, pagination and request body out of a plain command help', async () => {
    const help = await commandHelp(fakeCommandPackage(), 'list-widgets');
    expect(help).not.toContain('Stability:');
    expect(help).not.toContain('Pagination:');
    expect(help).not.toContain('Request body:');
  });
});

describe('help-json', () => {
  it('refuses an explicit --output text, since the manifest is JSON', async () => {
    expect((await runWidgetCli(['--output', 'text', 'help-json'])).out).toBe(
      "error: help-json always emits JSON (that's the contract — the manifest is the LLM-consumable serialized form of every command's meta). Drop `--output text` for this command. To browse the manifest as Markdown, use `widget-cli docs <command>` instead.\n"
    );
  });

  it('refuses an unknown category, naming the categories', async () => {
    expect((await runWidgetCli(['help-json', '--category', 'nope'])).out).toBe(
      'error: Unknown --category "nope". Available categories: calendar, chats, drive, excel, lifecycle, mail, meta, notes, sharepoint, tasks, teams, user.\ncode: cli_unknown_category\n'
    );
  });

  it('gives the terse entries of one category', async () => {
    const manifest = JSON.parse((await runWidgetCli(['help-json', '--terse', '--category', 'drive'])).out) as { readonly commands: ReadonlyArray<unknown> };
    expect(manifest.commands).toEqual([
      { name: 'get-widget-file', summary: 'Download a widget file.', category: 'drive' },
      { name: 'get-widget-images', summary: 'Extract the widget images.', category: 'drive' },
      { name: 'list-widgets', summary: 'Do a widget thing.', category: 'drive' },
    ]);
  });

  it('writes the manifest to --output-path, and refuses an empty path', async () => {
    const fs = createFileSystemFake();
    const saved = await runWidgetCli(['--output-path', 'out/help.json', 'help-json', '--terse'], { fs });
    expect(saved.out).toContain('savedTo: out/help.json');
    expect(JSON.parse(fs.snapshot('out/help.json') ?? '')).toEqual(expect.objectContaining({ package: '@widgets/cli' }));
    const empty = await runWidgetCli(['--output-path', '', 'help-json', '--terse']);
    expect(empty.out).toStartWith('error: --output-path: path argument is empty (likely a shell-quoting mistake — pass a real filesystem path)\n');
  });
});

describe('a mistyped name or flag', () => {
  it('suggests a flag of a lifecycle command', async () => {
    const run = await runWidgetCli(['help-json', '--terse', '--categry', 'x']);
    expect(run.out).toStartWith("error: unknown option '--categry'. Did you mean `--category`?\ncode: commander.unknownOption\n");
  });

  it('names the flag a missing required one was mistyped as', async () => {
    const run = await runWidgetCli(['list-widgets', '--shelf', 'x'], { commandPackage: fakeReadPackage() });
    expect(run.out).toStartWith("error: required option '--shelf-id <value>' not specified. `--shelf` is not a flag of this command. Did you mean `--shelf-id`?\n");
    expect(run.thrown).toBe('commander.missingMandatoryOptionValue');
  });

  it('says only that a required flag is missing when no mistyped flag was given', async () => {
    const run = await runWidgetCli(['list-widgets', '--top', '2'], { commandPackage: fakeReadPackage() });
    expect(run.out).toStartWith("error: required option '--shelf-id <value>' not specified\ncode: commander.missingMandatoryOptionValue\n");
  });

  it('reports an unknown command with the package hint', async () => {
    const run = await runWidgetCli(['nopecmd']);
    expect(run.out).toStartWith("error: unknown command 'nopecmd'\ncode: commander.unknownCommand\n");
    expect(run.thrown).toBe('commander.unknownCommand');
  });
});

describe('the lifecycle commands a package adds', () => {
  it('print through the CLI output, and fail with the bin names filled', async () => {
    const run = (args: ReadonlyArray<string>): Promise<Run> =>
      runWidgetCli(args, {
        addLifecycleCommands: (program, output) => {
          program.command('sign-in').action(() => output.print({ signedIn: true }));
          program.command('sign-out').action(() => output.fail('Not signed in. Run `{bin} sign-in` first.', 'not_signed_in'));
        },
      });
    expect((await run(['sign-in'])).out).toBe('signedIn: true\n');
    expect((await run(['sign-out'])).out).toStartWith('error: Not signed in. Run `widget-cli sign-in` first.\ncode: not_signed_in\n');
  });
});

describe('the top-level description', () => {
  it('names the search POSTs of the package in alphabetical order', async () => {
    const commandPackage = fakeCommandPackage({
      registry: {
        'search-widgets': fakeCommand({ graphMethod: 'POST', category: 'meta' }),
        'find-widgets': fakeCommand({ graphMethod: 'POST', category: 'meta' }),
        'list-widgets': fakeCommand(),
      },
    });
    const help = (await runWidgetCli([], { commandPackage })).out.replaceAll(/\s+/g, ' ');
    expect(help).toContain('1 GET endpoints + 2 search POST (find-widgets, search-widgets).');
  });
});
