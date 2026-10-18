import { describe, expect, it } from 'bun:test';
import { commandRole } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';
import { buildManifest, renderSingleCommand } from '@ask-marcel/office-core/use-cases/commands/docs-builders.ts';
import type { CommandManifestEntry } from '@ask-marcel/office-core/use-cases/commands/docs-render.ts';
import { AUTH_LIFECYCLE_ENTRIES } from './auth-lifecycle-entries.ts';

/*
 * The lifecycle commands of the auth package (package split, phase 4): its
 * three sign-in commands and its own `docs`, `help-json` and `mcp`. The auth
 * package has no registry, so these six entries are its whole manifest.
 */

const AUTH_PACKAGE = { registry: {}, lifecycle: AUTH_LIFECYCLE_ENTRIES, packageName: '@ask-marcel/office-auth' };
const AUTH_COMMANDS = ['docs', 'help-json', 'login', 'logout', 'mcp', 'status'];

const entry = (name: string): CommandManifestEntry => {
  const found = AUTH_LIFECYCLE_ENTRIES.find((candidate) => candidate.name === name);
  if (found === undefined) throw new Error(`no lifecycle entry ${name}`);
  return found;
};

// Every `{bin} <name>` and `{auth-bin} <name>` that a text of the entry names.
const namedCommands = (text: string): ReadonlyArray<string> => [...text.matchAll(/\{(?:auth-)?bin\} ([a-z][a-z-]*)/g)].map((match) => match[1] ?? '');

const textsOf = (lifecycle: CommandManifestEntry): ReadonlyArray<string> => [
  lifecycle.summary,
  lifecycle.example,
  lifecycle.responseShape ?? '',
  ...lifecycle.options.map((option) => option.description),
  ...(lifecycle.positionalArguments ?? []).map((argument) => argument.description),
];

describe('the lifecycle entries of the auth package', () => {
  it('lists the six commands of the auth package, and no Graph command', () => {
    expect(buildManifest(AUTH_PACKAGE, '1.0.0').commands.map((command) => command.name)).toEqual(AUTH_COMMANDS);
  });

  for (const name of AUTH_COMMANDS) {
    it(`gives \`${name}\` the meta that the docs and help-json render`, () => {
      const lifecycle = entry(name);
      expect(lifecycle.category).toBe('lifecycle');
      expect(lifecycle.summary.trim().length).toBeGreaterThan(40);
      expect(lifecycle.graphMethod).toBe('GET');
      expect(lifecycle.graphPathTemplate.startsWith('(lifecycle) ')).toBe(true);
      expect(lifecycle.graphDocsUrl.startsWith('https://')).toBe(true);
      const role = commandRole(name);
      const bin = role === undefined ? '{bin}' : `{${role}-bin}`;
      expect(lifecycle.example.startsWith(`${bin} ${name}`)).toBe(true);
      expect(lifecycle.responseShape?.trim().length ?? 0).toBeGreaterThan(0);
      for (const option of lifecycle.options) expect([option.name, option.key, option.description].map((text) => text.trim().length > 0)).toEqual([true, true, true]);
      for (const argument of lifecycle.positionalArguments ?? []) expect([argument.name, argument.description].map((text) => text.trim().length > 0)).toEqual([true, true]);
    });
  }

  it('names only commands that the auth package holds', () => {
    const named = AUTH_LIFECYCLE_ENTRIES.flatMap((lifecycle) => textsOf(lifecycle).flatMap(namedCommands));
    expect(named.filter((name) => !AUTH_COMMANDS.includes(name))).toEqual([]);
    expect(named.length).toBeGreaterThan(0);
  });

  it('keeps `login --force` optional, so a bare `login` still runs', () => {
    expect(entry('login').options.map((option) => [option.name, option.key, option.required])).toEqual([['force', 'force', false]]);
  });

  it('asks `docs` for the name of a command as a required positional argument', () => {
    expect(entry('docs').positionalArguments?.map((argument) => [argument.name, argument.required])).toEqual([['command', true]]);
    expect(entry('docs').example).toBe('{bin} docs login');
  });

  it('describes an MCP server with the login and status tools, and never offers the token helper', () => {
    const mcp = entry('mcp');
    expect(mcp.summary).toContain('two tools: `login` and `status`');
    expect(mcp.summary).toContain('claude mcp add --transport stdio --scope user {bin} -- {bin} mcp');
    expect(mcp.summary).not.toMatch(/token|run-command|list-commands|get-command-docs/);
  });

  it('renders the docs page of `help-json` with the one category of the auth package', () => {
    const page = renderSingleCommand(AUTH_PACKAGE, 'help-json');
    expect(page.ok).toBe(true);
    if (page.ok) expect(page.value).toContain('Every command of this package is in the category `lifecycle`');
  });
});
