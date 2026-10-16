import { describe, expect, it } from 'bun:test';
import { join } from 'node:path';
import { commandRole } from './command-directory.ts';

// Core's half of the bin-placeholder guard (package split, phase 3). The root's
// src/use-cases/commands/bin-placeholders.test.ts reads the root src/; the hints,
// remedies and docs text that moved to core/src/ are read here, so a command
// reference in core that names the wrong package, or a bin typed by hand, fails
// in core's own test run.

// What `{bin}` may stand before: the commands and help every package has, the
// generic words the hints use for "any command", and the `--` of
// `claude mcp add … {bin} -- {bin} mcp`, where the bin also names the server.
const EVERY_PACKAGE = new Set(['docs', 'help', 'help-json', 'mcp', '--help', '<command>', '<cmd>', 'subcommand', '--']);

// Words after a role placeholder that are not commands: the auth lock's holder.
const NOT_COMMANDS: Readonly<Record<string, ReadonlySet<string>>> = { auth: new Set(['process']) };

// The core folder, from this file rather than the working directory: a run
// from another directory would otherwise scan nothing and pass.
const CORE = join(import.meta.dir, '../../..');

type Line = { readonly where: string; readonly text: string };

const isComment = (text: string): boolean => /^\s*(\/\/|\/\*|\*)/.test(text);

const sourceLines = async (): Promise<ReadonlyArray<Line>> => {
  const lines: Line[] = [];
  for await (const path of new Bun.Glob('src/**/*.ts').scan(CORE)) {
    if (path.endsWith('.test.ts') || path.startsWith('src/test-helpers/')) continue;
    const text = await Bun.file(join(CORE, path)).text();
    text.split('\n').forEach((line, at) => {
      if (!isComment(line)) lines.push({ where: `core/${path}:${at + 1}`, text: line });
    });
  }
  return lines;
};

const misnamed = (role: string, word: string): boolean => {
  if (role === '') return !EVERY_PACKAGE.has(word) && !word.startsWith('${');
  return commandRole(word) !== role && !(NOT_COMMANDS[role]?.has(word) ?? false);
};

describe('the bin placeholders in core’s source', () => {
  it('are read from the whole of core, the hint text included, wherever the tests run from', async () => {
    const lines = await sourceLines();

    expect(lines.length).toBeGreaterThan(4_000);
    expect(lines.some(({ where }) => where.startsWith('core/src/presenter/error-hints.ts:'))).toBe(true);
  });

  it('name the package that holds each command they stand before', async () => {
    const wrong: string[] = [];
    for (const { where, text } of await sourceLines()) {
      for (const match of text.matchAll(/\{(?:(auth|read|write)-)?bin\}\s+([^\s`'"),.;\\]+)/g)) {
        const role = match[1] ?? '';
        const word = match[2] ?? '';
        if (misnamed(role, word)) wrong.push(`${where}: ${match[0]}`);
      }
    }
    expect(wrong).toEqual([]);
  });

  it('come in the four forms only', async () => {
    const odd = (await sourceLines()).flatMap(({ where, text }) =>
      [...text.matchAll(/\{[a-z-]*bin\}/g)].filter((m) => !/^\{((auth|read|write)-)?bin\}$/.test(m[0])).map((m) => `${where}: ${m[0]}`)
    );
    expect(odd).toEqual([]);
  });

  it('leave no bin typed by hand: every core line goes through a renderer that fills the placeholder', async () => {
    const byHand = (await sourceLines()).filter(({ text }) => text.includes('ask-marcel-office ')).map(({ where }) => where);
    expect(byHand).toEqual([]);
  });
});
