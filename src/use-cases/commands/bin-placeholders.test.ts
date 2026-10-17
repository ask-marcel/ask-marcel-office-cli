import { describe, expect, it } from 'bun:test';
import { join } from 'node:path';
import { commandRole } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';

// Every command reference in the CLI's own text names its bin through a
// placeholder (bin-names.ts): `{auth-bin}`, `{read-bin}` or `{write-bin}` before
// a command of that package (command-directory.ts), `{bin}` before what every
// package has. This guard reads the source, so a hint, a remedy or an example
// that names a command of the wrong package, or a bin by hand, fails here.

// What `{bin}` may stand before: the commands and help every package has, the
// generic words the hints use for "any command", and the `--` of
// `claude mcp add … {bin} -- {bin} mcp`, where the bin also names the server.
const EVERY_PACKAGE = new Set(['docs', 'help', 'help-json', 'mcp', '--help', '<command>', '<cmd>', 'subcommand', '--']);

// Words after a role placeholder that are not commands: the auth lock's holder.
const NOT_COMMANDS: Readonly<Record<string, ReadonlySet<string>>> = { auth: new Set(['process']) };

// The one line that keeps the bin by hand: written to stderr by the browser
// sign-in itself, with no renderer between it and the terminal to fill it.
const UNRENDERED_LINES = ['src/infra/browser-auth.ts: Clearing the browser session so the sign-in can be re-captured'];

// The repository root, from this file rather than the working directory: a run
// from another directory would otherwise scan nothing and pass.
const ROOT = join(import.meta.dir, '../../..');

const PLACEHOLDER = /\{(?:(?:auth|read|write)-)?bin\}/;

type Line = { readonly where: string; readonly text: string };

const isComment = (text: string): boolean => /^\s*(\/\/|\/\*|\*)/.test(text);

const sourceLines = async (): Promise<ReadonlyArray<Line>> => {
  const lines: Line[] = [];
  for await (const path of new Bun.Glob('src/**/*.ts').scan(ROOT)) {
    if (path.endsWith('.test.ts') || path.startsWith('src/test-helpers/')) continue;
    const text = await Bun.file(join(ROOT, path)).text();
    text.split('\n').forEach((line, at) => {
      if (!isComment(line)) lines.push({ where: `${path}:${at + 1}`, text: line });
    });
  }
  return lines;
};

const misnamed = (role: string, word: string): boolean => {
  if (role === '') return !EVERY_PACKAGE.has(word) && !word.startsWith('${');
  return commandRole(word) !== role && !(NOT_COMMANDS[role]?.has(word) ?? false);
};

describe('the bin placeholders in the source', () => {
  it('are read from the whole source, wherever the tests run from', async () => {
    expect((await sourceLines()).length).toBeGreaterThan(10_000);
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

  it('leave no bin typed by hand outside the one line no renderer fills', async () => {
    const byHand = (await sourceLines())
      .filter(({ text }) => text.includes('ask-marcel-office '))
      .filter(({ where, text }) => !UNRENDERED_LINES.some((allowed) => allowed.startsWith(where.split(':')[0] ?? '') && text.includes(allowed.split(': ')[1] ?? '')))
      .map(({ where }) => where);
    expect(byHand).toEqual([]);
  });
});

describe('the generated command reference', () => {
  it('carries the bins filled in, never a placeholder: gen-docs fills them', async () => {
    const manifest = await Bun.file(join(ROOT, 'docs/commands.json')).text();
    expect(manifest).toContain('"example": "ask-marcel-office list-drives');
    for (const doc of ['docs/commands.json', 'docs/COMMANDS.md']) {
      const text = await Bun.file(join(ROOT, doc)).text();
      expect({ doc, placeholder: PLACEHOLDER.exec(text)?.[0] }).toEqual({ doc, placeholder: undefined });
    }
  });
});
