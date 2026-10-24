import { describe, expect, it } from 'bun:test';
import type { Command } from 'commander';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import { createLoggerFake } from '@ask-marcel/office-core/test-helpers/logger-fake.ts';
import { COMMAND_DIRECTORY } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';
import { buildCli } from './cli.ts';

// The root package gives the core CLI its own texts (package split, phase 5,
// E5): the description that leads the top-level help, and the command that the
// docs help shows as its example. Whichever package holds the write commands, a
// person who reads the root help can find each of them.

const captureStdout = async (run: () => unknown): Promise<string> => {
  const original = process.stdout.write.bind(process.stdout);
  let captured = '';
  process.stdout.write = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  try {
    await run();
  } finally {
    process.stdout.write = original;
  }
  return captured;
};

const flat = (text: string): string => text.replaceAll(/\s+/g, ' ');

const rootCli = (): Command => buildCli({ graph: fakeGraphClient(), logger: createLoggerFake(), fs: createFileSystemFake() });

const WRITE_NAMES = [...COMMAND_DIRECTORY].filter(([, role]) => role === 'write').map(([name]) => name);

describe('the texts of the read package', () => {
  it('gives the docs help an example from its own registry: `ask-marcel-office-read docs list-mail-messages`', async () => {
    const docs = rootCli().commands.find((command) => command.name() === 'docs');
    expect(flat(await captureStdout(() => docs?.outputHelp()))).toContain('Example: ask-marcel-office-read docs list-mail-messages');
  });

  it('shows where each of the 7 write commands lives: in the list of the root commands, or in @ask-marcel/office-write', () => {
    const cli = rootCli();
    const listed = new Set(cli.commands.map((command) => command.name()));
    const description = cli.description();
    const nowhere = WRITE_NAMES.filter((name) => !listed.has(name) && !description.includes('@ask-marcel/office-write'));
    expect(WRITE_NAMES).toHaveLength(7);
    expect(nowhere).toEqual([]);
  });
});
