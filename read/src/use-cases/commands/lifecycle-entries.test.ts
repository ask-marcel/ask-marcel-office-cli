import { describe, expect, it } from 'bun:test';
import { buildManifest } from '@ask-marcel/office-core/use-cases/commands/docs-builders.ts';
import type { CommandManifestEntry } from '@ask-marcel/office-core/use-cases/commands/docs-render.ts';
import { LIFECYCLE_ENTRIES } from './lifecycle-entries.ts';

/*
 * The lifecycle commands of the read package: `docs`, `help-json` and `mcp`.
 * The sign-in commands are the auth package's (`auth-lifecycle-entries.ts`),
 * so read lists none of them.
 */

const ROOT_LIFECYCLE = { registry: {}, lifecycle: LIFECYCLE_ENTRIES, packageName: '@ask-marcel/office-read' };
const ROOT_COMMANDS = ['docs', 'help-json', 'mcp'];

const entry = (name: string): CommandManifestEntry => {
  const found = LIFECYCLE_ENTRIES.find((candidate) => candidate.name === name);
  if (found === undefined) throw new Error(`no lifecycle entry ${name}`);
  return found;
};

describe('the lifecycle entries of the read package', () => {
  it('lists docs, help-json and mcp, and no sign-in command', () => {
    expect(buildManifest(ROOT_LIFECYCLE, '1.0.0').commands.map((command) => command.name)).toEqual(ROOT_COMMANDS);
  });

  for (const name of ROOT_COMMANDS) {
    it(`gives \`${name}\` the meta that the docs and help-json render`, () => {
      const lifecycle = entry(name);
      expect(lifecycle.category).toBe('lifecycle');
      expect(lifecycle.summary.trim().length).toBeGreaterThan(40);
      expect(lifecycle.graphMethod).toBe('GET');
      expect(lifecycle.graphPathTemplate.startsWith('(lifecycle) ')).toBe(true);
      expect(lifecycle.graphDocsUrl.startsWith('https://')).toBe(true);
      expect(lifecycle.example.startsWith(`{bin} ${name}`)).toBe(true);
      expect(lifecycle.responseShape?.trim().length ?? 0).toBeGreaterThan(0);
      expect(lifecycle.options).toEqual([]);
    });
  }

  it('asks `docs` for the name of a command as a required positional argument', () => {
    const docs = entry('docs');
    expect(docs.positionalArguments?.map((argument) => [argument.name, argument.required])).toEqual([['command', true]]);
    expect(docs.positionalArguments?.[0]?.description.trim().length ?? 0).toBeGreaterThan(0);
  });

  it('describes an MCP server with the three gateway tools, and no sign-in tool', () => {
    const mcp = entry('mcp');
    expect(mcp.summary).toContain('three gateway tools (list-commands, get-command-docs, run-command)');
    expect(mcp.summary).not.toMatch(/login|status/);
  });
});
