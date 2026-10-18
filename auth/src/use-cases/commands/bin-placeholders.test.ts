import { describe, expect, it } from 'bun:test';
import { join } from 'node:path';
import { commandRole } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';

// Auth's half of the bin-placeholder guard (package split, phase 4). The root's
// src/use-cases/commands/bin-placeholders.test.ts reads the root src/ and core's
// reads core/src/; the sign-in texts, remedies and docs that moved to auth/src/
// are read here, so a command reference in auth that names the wrong package, or
// a bin typed by hand, fails in auth's own test run.

// What `{bin}` may stand before: the commands and help every package has, the
// generic words the hints use for "any command", and the `--` of
// `claude mcp add … {bin} -- {bin} mcp`, where the bin also names the server.
const EVERY_PACKAGE = new Set(['docs', 'help', 'help-json', 'mcp', '--help', '<command>', '<cmd>', 'subcommand', '--']);

// Words after a role placeholder that are not commands: the auth lock's holder.
const NOT_COMMANDS: Readonly<Record<string, ReadonlySet<string>>> = { auth: new Set(['process']) };

// The one file that may keep a bin by hand, on one line: the browser sign-in
// writes it to stderr itself, with no renderer between it and the terminal.
const UNRENDERED_FILE = 'src/infra/browser-auth.ts';
const UNRENDERED_LINE = 'Clearing the browser session so the sign-in can be re-captured';

// The auth folder, from this file rather than the working directory: a run
// from another directory would otherwise scan nothing and pass.
const AUTH = join(import.meta.dir, '../../..');

type Line = { readonly where: string; readonly text: string };

const isComment = (text: string): boolean => /^\s*(\/\/|\/\*|\*)/.test(text);

const sourceLines = async (): Promise<ReadonlyArray<Line>> => {
  const lines: Line[] = [];
  for await (const path of new Bun.Glob('src/**/*.ts').scan(AUTH)) {
    if (path.endsWith('.test.ts') || path.startsWith('src/test-helpers/')) continue;
    const text = await Bun.file(join(AUTH, path)).text();
    text.split('\n').forEach((line, at) => {
      if (!isComment(line)) lines.push({ where: `auth/${path}:${at + 1}`, text: line });
    });
  }
  return lines;
};

const misnamed = (role: string, word: string): boolean => {
  if (role === '') return !EVERY_PACKAGE.has(word) && !word.startsWith('${');
  return commandRole(word) !== role && !(NOT_COMMANDS[role]?.has(word) ?? false);
};

const isUnrendered = ({ where, text }: Line): boolean => where.startsWith(`auth/${UNRENDERED_FILE}:`) && text.includes(UNRENDERED_LINE);

describe('the bin placeholders in auth’s source', () => {
  it('are read from the whole of auth, the browser sign-in included, wherever the tests run from', async () => {
    const lines = await sourceLines();

    expect(lines.length).toBeGreaterThan(2_000);
    expect(lines.some(({ where }) => where.startsWith(`auth/${UNRENDERED_FILE}:`))).toBe(true);
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

  it('leave no bin typed by hand outside the one line of the browser sign-in that no renderer fills', async () => {
    const byHand = (await sourceLines())
      .filter(({ text }) => text.includes('ask-marcel-office '))
      .filter((line) => !isUnrendered(line))
      .map(({ where }) => where);
    expect(byHand).toEqual([]);
  });
});
