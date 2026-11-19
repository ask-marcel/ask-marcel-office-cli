import { describe, expect, it } from 'bun:test';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { checklistLinesOf } from './planner-checklist.ts';
import { snapshotOf, withNewKeys } from './planner-snapshot.ts';

/*
 * The edges of the Planner domain: a mark only at the start of a line and
 * before a space, the title after the mark trimmed, a bad first line, the
 * counts of a new task, and the keys of a restored checklist.
 */

describe('the edges of a checklist file', () => {
  it('reads a mark only at the start of a line and before a space', () => {
    expect(checklistLinesOf('Do [x] this\n[x]title\n[x]  Spaced out  ')).toEqual(ok([{ title: 'Do [x] this' }, { title: '[x]title' }, { title: 'Spaced out', isChecked: true }]));
  });

  it('names the first item when it is the bad one', () => {
    const refusal: Result<never, GraphError> = err({
      type: 'validation_error',
      code: 'checklist_item_invalid',
      message: '--checklist-file: item 1 has no title, or a title of more than 100 characters.',
    });
    expect(checklistLinesOf('[ ]\nTwo')).toEqual(refusal);
  });
});

describe('the edges of a snapshot', () => {
  it('gives a new task no checklist item, no reference and no description when Graph gives no count', () => {
    expect(snapshotOf({}, { checklistItemCount: 1, referenceCount: 1, hasDescription: true })).toEqual({ checklistItemCount: 0, referenceCount: 0, hasDescription: false });
  });

  it('keeps the key of each item that the checklist still has, and of each item to remove', () => {
    expect(withNewKeys({ c1: { title: 'A' }, c2: null }, { c1: { title: 'B' } }, () => 'new')).toEqual({ c1: { title: 'A' }, c2: null });
    expect(withNewKeys({ c1: { title: 'A' } }, { c1: 'not an item' }, () => 'new')).toEqual({ new: { title: 'A' } });
  });
});
