import { describe, expect, it } from 'bun:test';
import { join } from 'node:path';
import { commandRole } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';

// Write's half of the bin-placeholder guard (package split, phase 5). Every
// command reference in the CLI's own text names its bin through a placeholder
// (bin-names.ts): `{auth-bin}`, `{read-bin}` or `{write-bin}` before a command
// of that package (command-directory.ts), `{bin}` before what every package
// has. It finds its folder from its own path, so it reads write/src/ alone.
// Write's source has no line that keeps a bin by hand.

// What `{bin}` may stand before: the commands and help every package has, the
// generic words the hints use for "any command", and the `--` of
// `claude mcp add … {bin} -- {bin} mcp`, where the bin also names the server.
const EVERY_PACKAGE = new Set(['docs', 'help', 'help-json', 'mcp', '--help', '<command>', '<cmd>', 'subcommand', '--']);

// The folder that holds src/, from this file rather than the working directory:
// a run from another directory would otherwise scan nothing and pass.
const PACKAGE_ROOT = join(import.meta.dir, '../../..');

type Line = { readonly where: string; readonly text: string };

const isComment = (text: string): boolean => /^\s*(\/\/|\/\*|\*)/.test(text);

// A bin typed by hand before a word: the bin of the single package or of a
// package of the split.
const BIN_BY_HAND = /ask-marcel-office(-(auth|read|write))?\s/;

const sourceLines = async (): Promise<ReadonlyArray<Line>> => {
  const lines: Line[] = [];
  for await (const path of new Bun.Glob('src/**/*.ts').scan(PACKAGE_ROOT)) {
    if (path.endsWith('.test.ts') || path.startsWith('src/test-helpers/')) continue;
    const text = await Bun.file(join(PACKAGE_ROOT, path)).text();
    text.split('\n').forEach((line, at) => {
      if (!isComment(line)) lines.push({ where: `${path}:${at + 1}`, text: line });
    });
  }
  return lines;
};

const misnamed = (role: string, word: string): boolean => {
  if (role === '') return !EVERY_PACKAGE.has(word) && !word.startsWith('${');
  return commandRole(word) !== role;
};

describe('the bin placeholders in write’s source', () => {
  it('are read from the whole of write, the write registry included, wherever the tests run from', async () => {
    const lines = await sourceLines();
    expect(lines.length).toBeGreaterThan(800);
    expect(lines.some(({ where }) => where.startsWith('src/use-cases/commands/write-commands.ts:'))).toBe(true);
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

  it('leave no bin typed by hand', async () => {
    const byHand = (await sourceLines()).filter(({ text }) => BIN_BY_HAND.test(text)).map(({ where }) => where);
    expect(byHand).toEqual([]);
  });
});
