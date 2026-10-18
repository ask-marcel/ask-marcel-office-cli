import { describe, expect, it } from 'bun:test';
import { commandRole } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';
import { buildManifest, renderSingleCommand } from '@ask-marcel/office-core/use-cases/commands/docs-builders.ts';
import type { CommandManifestEntry } from '@ask-marcel/office-core/use-cases/commands/docs-render.ts';
import { commands } from './write-commands.ts';
import { WRITE_LIFECYCLE_ENTRIES } from './write-lifecycle-entries.ts';

/*
 * The lifecycle commands of the write package (package split, phase 5): its own
 * `docs`, `help-json` and `mcp`. The write package has no sign-in command: the
 * auth package signs in.
 */

const LIFECYCLE_ONLY = { registry: {}, lifecycle: WRITE_LIFECYCLE_ENTRIES, packageName: '@ask-marcel/office-write' };
const WRITE_PACKAGE = { registry: commands, lifecycle: WRITE_LIFECYCLE_ENTRIES };
const LIFECYCLE_COMMANDS = ['docs', 'help-json', 'mcp'];

const entry = (name: string): CommandManifestEntry => {
  const found = WRITE_LIFECYCLE_ENTRIES.find((candidate) => candidate.name === name);
  if (found === undefined) throw new Error(`no lifecycle entry ${name}`);
  return found;
};

// Every `{bin} <name>` and `{write-bin} <name>` that a text of the entry names.
const namedCommands = (text: string): ReadonlyArray<string> => [...text.matchAll(/\{(?:write-)?bin\} ([a-z][a-z-]*)/g)].map((match) => match[1] ?? '');

const textsOf = (lifecycle: CommandManifestEntry): ReadonlyArray<string> => [
  lifecycle.summary,
  lifecycle.example,
  lifecycle.responseShape ?? '',
  ...lifecycle.options.map((option) => option.description),
  ...(lifecycle.positionalArguments ?? []).map((argument) => argument.description),
];

describe('the lifecycle entries of the write package', () => {
  it('lists docs, help-json and mcp, and no sign-in command', () => {
    expect(buildManifest(LIFECYCLE_ONLY, '1.0.0').commands.map((command) => command.name)).toEqual(LIFECYCLE_COMMANDS);
  });

  for (const name of LIFECYCLE_COMMANDS) {
    it(`gives \`${name}\` the meta that the docs and help-json render`, () => {
      const lifecycle = entry(name);
      expect(lifecycle.category).toBe('lifecycle');
      expect(lifecycle.summary.trim().length).toBeGreaterThan(40);
      expect(lifecycle.graphMethod).toBe('GET');
      expect(lifecycle.graphPathTemplate.startsWith('(lifecycle) ')).toBe(true);
      expect(lifecycle.graphDocsUrl.startsWith('https://')).toBe(true);
      expect(commandRole(name)).toBeUndefined();
      expect(lifecycle.example.startsWith(`{bin} ${name}`)).toBe(true);
      expect(lifecycle.responseShape?.trim().length ?? 0).toBeGreaterThan(0);
      expect(lifecycle.options).toEqual([]);
      for (const argument of lifecycle.positionalArguments ?? []) expect([argument.name, argument.description].map((text) => text.trim().length > 0)).toEqual([true, true]);
    });
  }

  it('names only commands that the write package holds', () => {
    const named = WRITE_LIFECYCLE_ENTRIES.flatMap((lifecycle) => textsOf(lifecycle).flatMap(namedCommands));
    expect(named.filter((name) => !LIFECYCLE_COMMANDS.includes(name) && !Object.hasOwn(commands, name))).toEqual([]);
    expect(named.length).toBeGreaterThan(0);
  });

  it('asks `docs` for the name of a write command as a required positional argument', () => {
    expect(entry('docs').positionalArguments?.map((argument) => [argument.name, argument.required])).toEqual([['command', true]]);
    expect(entry('docs').example).toBe('{bin} docs create-mail-draft');
  });

  it('describes an MCP server with three gateway tools, of which only run-write-command runs a command', () => {
    const mcp = entry('mcp');
    expect(mcp.summary).toContain('three gateway tools (list-commands, get-command-docs, run-write-command)');
    expect(mcp.summary).toContain('claude mcp add --transport stdio --scope user {bin} -- {bin} mcp');
    expect(mcp.summary).not.toMatch(/run-command|login|token/);
  });

  it('gives help-json the two categories of the write commands', () => {
    expect(entry('help-json').summary).toContain('Categories: lifecycle, mail, calendar.');
    const categories = new Set(Object.values(commands).map((command) => command.meta.category));
    expect([...categories].toSorted((a, b) => a.localeCompare(b))).toEqual(['calendar', 'mail']);
  });

  it('renders the docs page of `docs` with the example of a write command', () => {
    const page = renderSingleCommand(WRITE_PACKAGE, 'docs');
    expect(page.ok).toBe(true);
    if (page.ok) expect(page.value).toContain('{bin} docs create-mail-draft');
  });
});
