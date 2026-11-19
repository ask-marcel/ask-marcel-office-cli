import { describe, expect, it } from 'bun:test';
import { checklistItemOf, referenceOf, referenceValue, restoreOf, snapshotOf, withNewKeys } from './planner-snapshot.ts';

/*
 * What the undo of a Planner change compares and writes back: the fields of
 * the PATCH body, and in a map only the entries that the body names.
 */

const ASSIGNED = { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: ' !' };

const task = {
  title: 'Write the report',
  dueDateTime: '2026-11-20T10:00:00Z',
  assignments: { me: { assignedDateTime: '2026-11-19T10:00:00Z' }, ann: { orderHint: 'x' } },
  appliedCategories: { category1: true, category2: false },
  other: 'not in the body',
};

const details = {
  description: 'Text',
  checklist: { c1: { title: 'Draft', isChecked: false, orderHint: 'h' }, c2: { title: 'Send', isChecked: true } },
  references: { 'https%3A//a%2Eb': { alias: 'Spec', type: 'Other', previewPriority: 'p' }, 'https%3A//c%2Ed': { alias: '' } },
};

describe('the snapshot of a Planner item', () => {
  it('gives the plain fields of the body, null for a field that the item does not have', () => {
    expect(snapshotOf(task, { title: 'x', startDateTime: 'y', percentComplete: 50 })).toEqual({ title: 'Write the report', startDateTime: null, percentComplete: null });
  });

  it('gives only the entries of a map that the body names', () => {
    expect(snapshotOf(task, { assignments: { me: null, bob: {} }, appliedCategories: { category1: false, category2: true, category3: true } })).toEqual({
      assignments: { me: true, bob: false },
      appliedCategories: { category1: true, category2: false, category3: false },
    });
    expect(snapshotOf(details, { checklist: { c1: null, c9: {} }, references: { 'https%3A//a%2Eb': null, 'https%3A//c%2Ed': null, gone: null } })).toEqual({
      checklist: { c1: { title: 'Draft', isChecked: false }, c9: null },
      references: { 'https%3A//a%2Eb': { alias: 'Spec', type: 'Other' }, 'https%3A//c%2Ed': { alias: null, type: null }, gone: null },
    });
  });

  it('compares HTML notes by their plain text description, and gives the defaults of the description and the preview', () => {
    expect(snapshotOf(details, { notes: { contentType: 'html', content: '<p>x</p>' }, previewType: 'noPreview' })).toEqual({ description: 'Text', previewType: 'automatic' });
    expect(snapshotOf({}, { description: 'x' })).toEqual({ description: '' });
    expect(snapshotOf({ categoryDescriptions: { category1: 'Red', category2: '' } }, { categoryDescriptions: { category1: 'x', category2: 'y', category3: 'z' } })).toEqual({
      categoryDescriptions: { category1: 'Red', category2: null, category3: null },
    });
  });

  it('compares a snapshot with the snapshot of a later read of the same shape', () => {
    const set = snapshotOf(task, { title: 'x', assignments: { ann: null } });
    expect(snapshotOf(task, set)).toEqual(set);
  });
});

describe('the restore of a Planner item', () => {
  it('gives the PATCH that writes back the values of before', () => {
    expect(restoreOf(task, { title: 'New', dueDateTime: null, assignments: { me: null, bob: ASSIGNED }, appliedCategories: { category1: false, category3: true } })).toEqual({
      title: 'Write the report',
      dueDateTime: '2026-11-20T10:00:00Z',
      assignments: { me: ASSIGNED, bob: null },
      appliedCategories: { category1: true, category3: false },
    });
  });

  it('writes back the checklist items, the references and the label names of before, and the plain text of HTML notes', () => {
    expect(restoreOf(details, { checklist: { c1: null, new1: {} }, references: { 'https%3A//a%2Eb': null, 'https%3A//x': {} }, notes: {} })).toEqual({
      checklist: { c1: { '@odata.type': '#microsoft.graph.plannerChecklistItem', title: 'Draft', isChecked: false }, new1: null },
      references: { 'https%3A//a%2Eb': { '@odata.type': '#microsoft.graph.plannerExternalReference', alias: 'Spec', type: 'Other' }, 'https%3A//x': null },
      description: 'Text',
    });
    expect(restoreOf({ categoryDescriptions: { category1: 'Red' } }, { categoryDescriptions: { category1: 'Blue', category2: 'Green' } })).toEqual({
      categoryDescriptions: { category1: 'Red', category2: null },
    });
  });

  it('gives a new key to an item that comes back but is no longer in the checklist', () => {
    let count = 0;
    const restore = { c1: { title: 'Draft' }, c2: null, c3: { title: 'Back' } };
    expect(withNewKeys(restore, { c1: { title: 'Changed' }, c2: { title: 'Added' } }, () => `k${++count}`)).toEqual({ c1: { title: 'Draft' }, c2: null, k1: { title: 'Back' } });
  });
});

describe('the values that Graph gives', () => {
  it('reads a checklist item and a reference, and refuses another shape', () => {
    expect(checklistItemOf({ title: 'A', isChecked: true, orderHint: 'x' })).toEqual({ title: 'A', isChecked: true });
    expect([checklistItemOf(null), checklistItemOf({ title: 'A' }), checklistItemOf({ isChecked: true }), checklistItemOf('A')]).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(referenceOf({ alias: 'Spec', type: 'Word' })).toEqual({ alias: 'Spec', type: 'Word' });
    expect([referenceOf(null), referenceOf([]), referenceOf('x')]).toEqual([undefined, undefined, undefined]);
    expect(referenceValue({ alias: null, type: null })).toEqual({ '@odata.type': '#microsoft.graph.plannerExternalReference' });
  });
});
