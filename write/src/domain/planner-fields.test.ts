import { describe, expect, it } from 'bun:test';
import { byHint, hintBetween, isLabelKey, labelPatch, labelsOn, plannerDayOf, priorityOf } from './planner-fields.ts';

/*
 * The fields of a Planner task in the Graph form, and the order hints of
 * Planner: a day at 10:00 UTC, a priority from 0 to 10, 25 labels, and the
 * hint "<previous> <next>!" of the Graph docs.
 */

describe('the fields of a Planner task', () => {
  it('keeps a day at 10:00 UTC, removes it with none, and refuses a day that is not in the calendar', () => {
    expect(plannerDayOf('2026-11-20')).toBe('2026-11-20T10:00:00Z');
    expect(plannerDayOf('none')).toBeNull();
    for (const bad of ['2026-13-01', '2026-02-30', '20-11-2026', '2026-11-20T10:00', '']) expect(plannerDayOf(bad)).toBeUndefined();
  });

  it('gives the number of a priority name or of a number from 0 to 10', () => {
    expect([priorityOf('urgent'), priorityOf('important'), priorityOf('medium'), priorityOf('low')]).toEqual([1, 3, 5, 9]);
    expect([priorityOf('0'), priorityOf('7'), priorityOf('10')]).toEqual([0, 7, 10]);
    for (const bad of ['11', '-1', '1.5', 'high', 'toString', '']) expect(priorityOf(bad)).toBeUndefined();
  });

  it('knows the 25 label keys', () => {
    expect(['category1', 'category9', 'category10', 'category19', 'category20', 'category25'].every(isLabelKey)).toBe(true);
    expect(['category0', 'category26', 'category', 'Category1', 'category1 ', 'xcategory1'].some(isLabelKey)).toBe(false);
  });

  it('gives the labels that are on, sorted, and the patch from one set of labels to another', () => {
    expect(labelsOn({ category5: true, category2: true, category3: false })).toEqual(['category2', 'category5']);
    expect(labelPatch({ category1: true, category2: true, category4: false }, ['category2', 'category4'])).toEqual({ category1: false, category4: true });
    expect(labelPatch({ category1: true }, ['category1'])).toEqual({});
  });
});

describe('the order hints of Planner', () => {
  it('composes the hint between two items, before the first and after the last', () => {
    expect(hintBetween('5637', 'adhg')).toBe('5637 adhg!');
    expect(hintBetween(undefined, '5637')).toBe(' 5637!');
    expect(hintBetween('adhg', undefined)).toBe('adhg !');
    expect(hintBetween(undefined, undefined)).toBe(' !');
  });

  it('sorts hints by the ordinal of their characters, a shorter hint first', () => {
    expect(['adhg', '5637', 'a', 'ab', ' 5637!', 'Z'].toSorted(byHint)).toEqual([' 5637!', '5637', 'Z', 'a', 'ab', 'adhg']);
    expect(byHint('a', 'a')).toBe(0);
  });
});
