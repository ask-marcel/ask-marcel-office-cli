import { describe, expect, it } from 'bun:test';
import { commandRole } from '@ask-marcel/office-core/use-cases/commands/command-directory.ts';
import { renderSingleCommand } from '@ask-marcel/office-core/use-cases/commands/docs-builders.ts';
import { commands } from './write-commands.ts';
import { WRITE_LIFECYCLE_ENTRIES } from './write-lifecycle-entries.ts';

/*
 * The `undo` lifecycle entry of the write package: `help-json`, `docs undo` and
 * the MCP `get-command-docs undo` read it.
 */

const UNDO = WRITE_LIFECYCLE_ENTRIES.find((entry) => entry.name === 'undo');

describe('the undo entry of the write package', () => {
  it('asks for the run id as a required positional argument and takes --apply as an option', () => {
    expect(UNDO?.positionalArguments?.map((argument) => [argument.name, argument.required])).toEqual([['runId', true]]);
    expect(UNDO?.options.map((option) => [option.name, option.key, option.required])).toEqual([['apply', 'apply', false]]);
    expect(UNDO?.example).toBe('undo 20261118T090507Z-0a1b2c3d --apply');
  });

  it('is a lifecycle entry of the write package, with no command of the registry under its name', () => {
    expect([UNDO?.category, UNDO?.graphMethod, commandRole('undo'), Object.hasOwn(commands, 'undo')]).toEqual(['lifecycle', 'DELETE', undefined, false]);
    expect(UNDO?.graphPathTemplate.startsWith('(lifecycle) ')).toBe(true);
    expect(UNDO?.graphDocsUrl).toBe('https://learn.microsoft.com/en-us/graph/api/message-delete');
  });

  it('says what the undo does, its MCP tools, its statuses and its error codes', () => {
    expect(UNDO?.summary).toContain('moves it to Deleted Items');
    expect(UNDO?.summary).toContain('`undo` (read-only) and `undo-apply` (destructive)');
    expect(UNDO?.responseShape).toContain('`will-undo`, `undone`, `already-undone`, `changed`, `missing` or `failed`');
    expect(UNDO?.responseShape).toContain('`run_already_undone`');
  });

  it('renders a docs page with the run id and --apply', () => {
    const page = renderSingleCommand({ registry: commands, lifecycle: WRITE_LIFECYCLE_ENTRIES }, 'undo');
    expect(page.ok).toBe(true);
    if (!page.ok) return;
    expect(page.value).toContain('| `<runId>` | yes |');
    expect(page.value).toContain('| `--apply` |');
  });

  it('names the undo tools in the summary of mcp', () => {
    expect(WRITE_LIFECYCLE_ENTRIES.find((entry) => entry.name === 'mcp')?.summary).toContain('and the undo tools undo and undo-apply');
  });
});
