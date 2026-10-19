import { describe, expect, it } from 'bun:test';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import { fakeWriteGraph } from '@ask-marcel/office-core/test-helpers/write-graph-fake.ts';
import { buildWriteCli } from './write-program.ts';

// The CLI of the write package `@ask-marcel/office-write`: its seven commands
// and its lifecycle commands under the bin `ask-marcel-office-write`, on the
// write graph alone. The help cases call the bare command: commander ends the
// process on `--help`, and the test runner then stops with no report.

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
    buildWriteCli({ graph: fakeWriteGraph(), logger: createLoggerFake(), fs: createFileSystemFake(), version: '1.0.0', onCommandError: () => undefined }).parseAsync([
      'node',
      'ask-marcel-office-write',
      ...args,
    ])
  );

// The commands that the help lists under one group heading, e.g. "Mail:".
const listedUnder = (help: string, heading: string): ReadonlyArray<string> => {
  const group = help.split(/\n(?=\S)/).find((block) => block.startsWith(heading)) ?? '';
  return [...group.matchAll(/^ {2}([a-z][a-z-]*)\b/gm)].map((match) => match[1] ?? '');
};

type Manifest = { readonly package: string; readonly commands: ReadonlyArray<{ readonly name: string }> };

describe('the CLI of the write package', () => {
  it('answers --version with the version it is given', async () => {
    expect(await run(['--version'])).toBe('1.0.0\n');
  });

  it('lists the seven write commands, help-json and docs in its help, and not mcp, which the bin serves before commander', async () => {
    const help = await run([]);
    expect(help).toContain('Usage: ask-marcel-office-write');
    expect([...listedUnder(help, 'Mail:'), ...listedUnder(help, 'Calendar:')].toSorted((a, b) => a.localeCompare(b))).toEqual([
      'convert-calendar-event-attachment-to-pdf',
      'convert-group-post-attachment-to-pdf',
      'convert-mail-attachment-to-pdf',
      'create-forward-draft',
      'create-mail-draft',
      'create-reply-draft',
      'update-mail-draft',
    ]);
    expect(listedUnder(help, 'Commands:')).toEqual(['help-json']);
    expect(listedUnder(help, 'Lifecycle:')).toEqual(['docs']);
    expect(help).not.toMatch(/^ {2}mcp\b/m);
  });

  it('says that all seven commands write, and gives no sentence on GET endpoints', async () => {
    const help = (await run([])).replaceAll(/\s+/g, ' ');
    expect(help).toContain('Microsoft Graph write commands for agents');
    expect(help).toContain('All 7 commands write.');
    expect(help).not.toContain('GET endpoints');
    expect(help).not.toContain('ONLY writes');
  });

  it('lists the seven commands and docs, help-json and mcp in help-json, under the write package', async () => {
    const manifest = JSON.parse(await run(['help-json'])) as Manifest;
    expect(manifest.package).toBe('@ask-marcel/office-write');
    expect(manifest.commands).toHaveLength(10);
    expect(manifest.commands.map((command) => command.name).filter((name) => ['docs', 'help-json', 'mcp'].includes(name))).toEqual(['docs', 'help-json', 'mcp']);
  });

  it('sends a caller of list-drives to the read package', async () => {
    expect(await run(['list-drives'])).toContain('`list-drives` lives in @ask-marcel/office-read: install it and run `ask-marcel-office-read list-drives`.');
  });

  it('shows the docs of create-mail-draft with what it writes, and the write bin in its example', async () => {
    const page = await run(['docs', 'create-mail-draft']);
    expect(page).toContain('Writes to Microsoft 365');
    expect(page).toContain('ask-marcel-office-write create-mail-draft');
  });
});
