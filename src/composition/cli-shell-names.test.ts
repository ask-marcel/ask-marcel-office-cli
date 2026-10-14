import { describe, expect, it } from 'bun:test';
import { err, ok } from '../domain/result.ts';
import { fakeCommand, fakeCommandPackage, fakeReadPackage } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import type { CommandPackage } from '../use-cases/commands/command-package.ts';
import { buildCliShell } from './cli-shell.ts';

// The CLI of the read package of a split (widget-cli; auth is widget-auth, write
// is widget-write). Its texts carry the bin placeholders, and every one the CLI
// prints must come out filled with the bin of the package that holds the command.

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

const cliOf = (commandPackage: CommandPackage): ReturnType<typeof buildCliShell> =>
  buildCliShell({ commandPackage, graph: fakeGraphClient(), logger: createLoggerFake(), fs: createFileSystemFake(), version: '1.2.3' });

const run = (args: ReadonlyArray<string>, commandPackage: CommandPackage = fakeReadPackage()): Promise<string> =>
  captureStdout(() => cliOf(commandPackage).parseAsync(['node', 'widget-cli', ...args]));

const commandHelp = (name: string): Promise<string> =>
  captureStdout(async () =>
    cliOf(fakeReadPackage())
      .commands.find((c) => c.name() === name)
      ?.outputHelp()
  );

const flat = (text: string): string => text.replaceAll(/\s+/g, ' ');

const PLACEHOLDER = /\{(?:(?:auth|read|write)-)?bin\}/;

describe('the bins the CLI of a package names', () => {
  it('fills the summary, the options and the example of a command in its help', async () => {
    const help = flat(await commandHelp('list-widgets'));
    expect(help).toContain('List the widgets. Page with `widget-cli next-page`.');
    expect(help).toContain('How many widgets; `widget-auth status` shows the token first.');
    expect(help).toContain('Example: widget-cli list-widgets --top 3');
    expect(help).not.toMatch(PLACEHOLDER);
  });

  it('fills the description of a required option in the help of its command', async () => {
    expect(flat(await commandHelp('list-widgets'))).toContain('The shelf, as `widget-cli list-shelves` returns it.');
  });

  it('fills the page of help <command>', async () => {
    const page = await run(['help', 'list-widgets']);
    expect(page).toContain('widget-cli list-widgets --top 3');
    expect(page).not.toMatch(PLACEHOLDER);
  });

  it('fills the full and the terse manifest of one category', async () => {
    const full = await run(['help-json', '--category', 'drive']);
    expect(full).toContain('"example":"widget-cli list-widgets --top 3"');
    expect(full).not.toMatch(PLACEHOLDER);
    const terse = fakeReadPackage({ registry: { 'list-widgets': fakeCommand({ summary: 'List the widgets of `{read-bin} list-drives`. Then more.' }) } });
    expect(await run(['help-json', '--terse', '--category', 'drive'], terse)).toContain('"summary":"List the widgets of `widget-cli list-drives`."');
  });

  it('fills the summaries of the top-level listing', async () => {
    const help = flat(await run([]));
    expect(help).toContain('List the widgets.');
    expect(help).not.toMatch(PLACEHOLDER);
  });

  it('fills the full and the terse manifest of help-json', async () => {
    const full = await run(['help-json']);
    expect(full).toContain('"example":"widget-cli list-widgets --top 3"');
    expect(full).toContain('"example":"widget-auth sign-in"');
    expect(full).not.toMatch(PLACEHOLDER);
    expect(await run(['help-json', '--terse'])).toContain('"summary":"Download a widget file."');
  });

  it('fills the docs page of a command and of a lifecycle entry', async () => {
    const page = await run(['docs', 'list-widgets']);
    expect(page).toContain('| `--top` | How many widgets; `widget-auth status` shows the token first. |');
    expect(page).toContain('widget-cli list-widgets --top 3');
    expect(await run(['docs', 'sign-in'])).toContain('Sign in to the widget service. Check with `widget-auth status`.');
  });

  it('fills the message a command fails with and the hint under it', async () => {
    const commandPackage = fakeReadPackage({
      registry: {
        'list-widgets': {
          ...fakeCommand(),
          execute: async () => err({ type: 'auth_failed', message: 'Not signed in. Run `{auth-bin} login`.', code: 'InvalidAuthenticationToken' }),
        },
      },
    });
    const out = await run(['list-widgets'], commandPackage);
    expect(out).toContain('error: Not signed in. Run `widget-auth login`.');
    expect(out).toContain('hint: The cached access token is invalid or expired. Run `widget-auth login` to re-authenticate');
    expect(out).not.toMatch(PLACEHOLDER);
  });

  it('prints the next page of an answer as a command of the package that holds next-page', async () => {
    const commandPackage = fakeReadPackage({
      registry: { 'list-widgets': { ...fakeCommand(), execute: async () => ok({ value: [{ id: 'w1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/widgets?$skip=1' }) } },
    });
    expect(await run(['list-widgets'], commandPackage)).toContain("--- next: widget-cli next-page --url 'https://graph.microsoft.com/v1.0/widgets?$skip=1'");
  });
});

describe('a command the package does not hold', () => {
  it('tells the caller which package holds a write command, and the command to run there', async () => {
    const out = await run(['create-mail-draft']);
    expect(out).toContain(
      "error: unknown command 'create-mail-draft'. `create-mail-draft` lives in @ask-marcel/office-write: install it and run `widget-write create-mail-draft`."
    );
  });

  it('tells the caller of a write package which package holds a read command', async () => {
    const writePackage = fakeReadPackage({ bins: { auth: 'widget-auth', read: 'widget-read', write: 'widget-cli' } });
    expect(await run(['list-drives'], writePackage)).toContain(
      "error: unknown command 'list-drives'. `list-drives` lives in @ask-marcel/office-read: install it and run `widget-read list-drives`."
    );
  });

  it('sends the docs of an auth command to the auth package', async () => {
    const out = await run(['docs', 'login']);
    expect(out).toContain(
      'error: Unknown command "login". `login` lives in @ask-marcel/office-auth: install it and run `widget-auth login`. Run `widget-cli --help` to list every command.'
    );
  });

  it('offers a write command for a near miss, under the bin that holds it', async () => {
    expect(await run(['create-mail-draf'])).toContain('Did you mean `widget-write create-mail-draft`?');
  });

  it('offers no command of the directory that a package with one bin for all three does not have', async () => {
    const out = await run(['create-mail-draf'], fakeCommandPackage());
    expect(out).toContain("error: unknown command 'create-mail-draf'");
    expect(out).not.toContain('create-mail-draft');
  });

  it('never offers the token helper for a near miss', async () => {
    expect(await run(['toke'])).not.toContain('token`');
  });

  it('tells the caller of token that it is the token helper of the auth package', async () => {
    expect(await run(['token'])).toContain(
      "error: unknown command 'token'. `token` is the token helper of @ask-marcel/office-auth: the read and write packages run `widget-auth token` themselves."
    );
  });

  it('sends nobody elsewhere for a command of the directory that a package with one bin for all three does not list', async () => {
    expect(await run(['docs', 'token'], fakeCommandPackage())).toContain('error: Unknown command "token". Run `widget-cli --help` to list every command.');
  });

  it('names status under the auth bin as what replaced scopes-check', async () => {
    expect(await run(['scopes-check'])).toContain("error: unknown command 'scopes-check'. `widget-auth status` replaced `scopes-check`.");
  });

  it('answers the removed update command with the install command of the package', async () => {
    expect(await run(['update'])).toContain(
      "error: unknown command 'update'. `update` was removed: to upgrade, install the package again with `npm i -g @widgets/cli@latest` (or `bun add -g @widgets/cli@latest`)."
    );
  });
});
