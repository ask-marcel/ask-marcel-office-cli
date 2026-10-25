import { describe, expect, it } from 'bun:test';
import type { CommandManifestEntry } from './docs-render.ts';
import { renderLifecycleTable } from './docs-render.ts';

// The lifecycle table of the COMMANDS.md of a package. gen-docs writes it
// between the AUTO-GENERATED-LIFECYCLE markers, so each package lists its own
// lifecycle commands with its own bin.

const lifecycleEntry = (name: string, summary: string): CommandManifestEntry => ({
  name,
  summary,
  category: 'lifecycle',
  graphMethod: 'GET',
  graphPathTemplate: '(lifecycle) not a Graph endpoint',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/',
  options: [],
  example: `{bin} ${name}`,
});

const withReadBin = (text: string): string => text.replaceAll('{bin}', 'ask-marcel-office-read');

const HEAD = ['| Command | What it does |', '|:--|:--|'];

describe('the lifecycle table of a package', () => {
  it('gives only the table head when the package has no lifecycle command', () => {
    expect(renderLifecycleTable([], withReadBin)).toBe(HEAD.join('\n'));
  });

  it('lists the lifecycle commands by name, in any order of the entries: docs, help-json, mcp', () => {
    const entries = [lifecycleEntry('mcp', 'Start the server.'), lifecycleEntry('docs', 'Print the docs.'), lifecycleEntry('help-json', 'Print the manifest.')];
    expect(renderLifecycleTable(entries, withReadBin)).toBe(
      [
        ...HEAD,
        '| `ask-marcel-office-read docs` | Print the docs. |',
        '| `ask-marcel-office-read help-json` | Print the manifest. |',
        '| `ask-marcel-office-read mcp` | Start the server. |',
      ].join('\n')
    );
  });

  it('names each command with the bin of the package: `ask-marcel-office-write docs` in the write package', () => {
    const withWriteBin = (text: string): string => text.replaceAll('{bin}', 'ask-marcel-office-write');
    expect(renderLifecycleTable([lifecycleEntry('docs', 'Print the docs.')], withWriteBin)).toBe([...HEAD, '| `ask-marcel-office-write docs` | Print the docs. |'].join('\n'));
  });

  it('gives the first sentence of a summary of two sentences, with its bins filled', () => {
    const mcp = lifecycleEntry('mcp', 'Start the server of {bin} over stdio. Register it with `claude mcp add`.');
    expect(renderLifecycleTable([mcp], withReadBin)).toBe([...HEAD, '| `ask-marcel-office-read mcp` | Start the server of ask-marcel-office-read over stdio. |'].join('\n'));
  });
});
