import { describe, expect, it } from 'bun:test';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err } from '@ask-marcel/office-core/domain/result.ts';
import { checklistLinesOf, checklistPatchOf, checklistValue } from './planner-checklist.ts';

/*
 * A checklist file sets the checklist of a Planner task: one item on each
 * line, `[x] ` checked, `[ ] ` not checked. The patch keeps the items of the
 * same title, removes the others and adds the new ones.
 */

const refusal = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

describe('a checklist file', () => {
  it('gives one item for each line that is not empty, with the state of its mark', () => {
    expect(checklistLinesOf('Book the room\r\n[x] Send the agenda\n\n  [ ] Print the slides  \n[X] Call the team\n')).toEqual({
      ok: true,
      value: [
        { title: 'Book the room' },
        { title: 'Send the agenda', isChecked: true },
        { title: 'Print the slides', isChecked: false },
        { title: 'Call the team', isChecked: true },
      ],
    });
  });

  it('refuses a file with no item', () => {
    expect(checklistLinesOf(' \n\n')).toEqual(refusal('checklist_file_empty', '--checklist-file: the file has no item. To remove an item, use remove-planner-checklist-item.'));
  });

  it('takes 20 items and refuses 21', () => {
    const lines = (count: number): string => Array.from({ length: count }, (_, index) => `Item ${index}`).join('\n');
    expect(checklistLinesOf(lines(20)).ok).toBe(true);
    expect(checklistLinesOf(lines(21))).toEqual(refusal('checklist_too_long', '--checklist-file: the file has 21 items. Planner keeps at most 20.'));
  });

  it('refuses an item with no title or a title of more than 100 characters, and names the item, not its text', () => {
    const message = '--checklist-file: item 2 has no title, or a title of more than 100 characters.';
    expect(checklistLinesOf('One\n[x] \nThree')).toEqual(refusal('checklist_item_invalid', message));
    expect(checklistLinesOf(`One\n${'a'.repeat(101)}`)).toEqual(refusal('checklist_item_invalid', message));
    expect(checklistLinesOf(`One\n${'a'.repeat(100)}`).ok).toBe(true);
  });

  it('refuses a title twice and names the second item', () => {
    expect(checklistLinesOf('One\nTwo\n[x] One')).toEqual(refusal('checklist_item_twice', '--checklist-file: item 3 has the same title as an item before it.'));
  });
});

describe('the patch of a checklist', () => {
  const current = { k1: { title: 'Keep', isChecked: true }, k2: { title: 'Remove', isChecked: false }, k3: { title: 'Check', isChecked: false } };
  const keys = (): (() => string) => {
    let count = 0;
    return () => `new-${++count}`;
  };

  it('keeps an item of the same title, removes the others, adds the new ones, and changes a state that a mark gives', () => {
    const lines = [{ title: 'Keep' }, { title: 'Check', isChecked: true }, { title: 'Add' }, { title: 'Add checked', isChecked: true }];
    expect(checklistPatchOf(current, lines, keys())).toEqual({
      k2: null,
      k3: { title: 'Check', isChecked: true },
      'new-1': { title: 'Add', isChecked: false },
      'new-2': { title: 'Add checked', isChecked: true },
    });
  });

  it('is empty when the file gives the checklist as it is', () => {
    expect(checklistPatchOf(current, [{ title: 'Keep', isChecked: true }, { title: 'Remove' }, { title: 'Check', isChecked: false }], keys())).toEqual({});
  });

  it('gives the Graph value of an item', () => {
    expect(checklistValue({ title: 'Keep', isChecked: true })).toEqual({ '@odata.type': '#microsoft.graph.plannerChecklistItem', title: 'Keep', isChecked: true });
  });
});
