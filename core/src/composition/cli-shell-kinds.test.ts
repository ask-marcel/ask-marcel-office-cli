import { describe, expect, it } from 'bun:test';
import type { CommandPackage } from '../use-cases/commands/command-package.ts';
import { fakeCommand, fakeCommandPackage, SIGN_IN_ENTRY } from '../test-helpers/command-package-fake.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '../test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { buildCliShell } from './cli-shell.ts';

// A package of the split holds one kind of command (package split, phase 5):
// the read package only reads, the write package only writes. Its CLI says what
// the package is in its own words, and adds only the facts its registry gives.

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

const flat = (text: string): string => text.replaceAll(/\s+/g, ' ');

const LIFECYCLE = [
  { ...SIGN_IN_ENTRY, name: 'docs' },
  { ...SIGN_IN_ENTRY, name: 'help-json' },
  { ...SIGN_IN_ENTRY, name: 'mcp' },
];

const READS_ONLY = fakeCommandPackage({
  registry: {
    'list-widgets': fakeCommand(),
    'get-widget-file': fakeCommand({ summary: 'Download a widget file.', producesBytes: true }),
    'find-widgets': fakeCommand({ summary: 'Search the widgets.', category: 'meta', graphMethod: 'POST' }),
  },
  lifecycle: LIFECYCLE,
  description: 'Widget reader. It changes no widget.',
  exampleCommand: 'get-widget-file',
});

const WRITES_ONLY = fakeCommandPackage({
  registry: {
    'save-widget-draft': fakeCommand({ summary: 'Save a widget draft.', category: 'mail', graphMethod: 'POST', effect: 'draft' }),
    'update-widget-draft': fakeCommand({ summary: 'Update a widget draft.', category: 'mail', graphMethod: 'PATCH', effect: 'draft' }),
  },
  lifecycle: LIFECYCLE,
  description: 'Widget writer. It sends nothing.',
  exampleCommand: 'save-widget-draft',
});

const topHelp = (commandPackage: CommandPackage): Promise<string> => {
  const cli = buildCliShell({ commandPackage, graph: fakeGraphClient(), logger: createLoggerFake(), fs: createFileSystemFake() });
  return captureStdout(async () => cli.outputHelp());
};

const docsHelp = (commandPackage: CommandPackage): Promise<string> => {
  const cli = buildCliShell({ commandPackage, graph: fakeGraphClient(), logger: createLoggerFake(), fs: createFileSystemFake() });
  const docs = cli.commands.find((candidate) => candidate.name() === 'docs');
  return captureStdout(async () => docs?.outputHelp());
};

describe('the CLI of a package that holds only reads', () => {
  it('starts its description with the words of the package, then counts its GET and search POST reads, and names no write', async () => {
    const help = flat(await topHelp(READS_ONLY));
    expect(help).toContain('Widget reader. It changes no widget. 2 GET endpoints + 1 search POST (find-widgets).');
    expect(help).not.toContain('Read-mostly');
    expect(help).not.toContain('commands write.');
  });

  it("gives the docs help the package's example and its own lifecycle commands", async () => {
    const help = flat(await docsHelp(READS_ONLY));
    expect(help).toContain('Example: widget-cli docs get-widget-file');
    expect(help).toContain('Lifecycle: `widget-cli docs help-json` (or mcp / docs) prints the same --help');
    expect(help).toContain('Lifecycle commands (docs/help-json/mcp) are also covered');
    expect(help).not.toContain('login');
  });
});

describe('the CLI of a package that holds only writes', () => {
  it('starts its description with the words of the package, then says that every command writes and what each leaves behind, with no read count', async () => {
    const help = flat(await topHelp(WRITES_ONLY));
    expect(help).toContain('Widget writer. It sends nothing. All 2 commands write. save-widget-draft, update-widget-draft: create or update an UNSENT');
    expect(help).not.toContain('Read-mostly');
    expect(help).not.toContain('GET endpoints');
  });

  it("gives the docs help the package's example", async () => {
    expect(flat(await docsHelp(WRITES_ONLY))).toContain('Example: widget-cli docs save-widget-draft');
  });
});

describe('the CLI of a package whose only lifecycle entry is its own command', () => {
  it('offers docs as the other lifecycle command of the docs help', async () => {
    const help = flat(await docsHelp(fakeCommandPackage({ lifecycle: [SIGN_IN_ENTRY] })));
    expect(help).toContain('Lifecycle: `widget-cli docs sign-in` (or docs) prints the same --help');
  });
});

describe('the CLI of a package that holds reads and writes', () => {
  it('names its writes and what each leaves behind, then counts its reads, one sentence after the other', async () => {
    const help = flat(await topHelp(fakeCommandPackage()));
    expect(help).toContain(
      'Widget CLI. Read-mostly by design — the ONLY writes are 1 commands. save-widget-draft: create or update an UNSENT mail draft (this CLI cannot send mail). 3 GET endpoints + 1 search POST (find-widgets).'
    );
  });
});
