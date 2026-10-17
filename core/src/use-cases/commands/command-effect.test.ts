import { describe, expect, it } from 'bun:test';
import { fakeCommand } from '../../test-helpers/command-package-fake.ts';
import { describeWrites, WRITE_EFFECTS, writeGroups } from './command-effect.ts';

/*
 * Each write class is worded once (core, package split phase 3): the help, the
 * MCP tool descriptions and the docs read the groups and the sentence here.
 */

const registry = {
  'save-widget-draft': fakeCommand({ effect: 'draft' }),
  'list-widgets': fakeCommand(),
  'convert-widget-to-pdf': fakeCommand({ effect: 'transient-upload' }),
  'reply-widget-draft': fakeCommand({ effect: 'draft' }),
};

describe('the writes of a registry', () => {
  it('fall in one group per write class, in the order the registry first names it, with sorted names', () => {
    expect(writeGroups(registry)).toEqual([
      { effect: 'draft', names: ['reply-widget-draft', 'save-widget-draft'] },
      { effect: 'transient-upload', names: ['convert-widget-to-pdf'] },
    ]);
  });

  it('are described as one sentence per write class', () => {
    expect(describeWrites(registry)).toBe(
      'reply-widget-draft, save-widget-draft: create or update an UNSENT mail draft (this CLI cannot send mail). convert-widget-to-pdf: upload the attachment to a temporary file in the `.ask-marcel-temp` folder of your OneDrive to convert it, then try to delete that file (a failed cleanup can leave it there)'
    );
  });

  it('are none in a registry that only reads', () => {
    expect(writeGroups({ 'list-widgets': fakeCommand() })).toEqual([]);
    expect(describeWrites({ 'list-widgets': fakeCommand() })).toBe('');
  });

  it('change nothing destructively today', () => {
    expect(Object.values(WRITE_EFFECTS).map((effect) => effect.destructive)).toEqual([false, false]);
  });
});
