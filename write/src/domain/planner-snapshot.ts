import { checklistValue } from './planner-checklist.ts';
import type { ChecklistItem } from './planner-checklist.ts';
import { ASSIGNMENT } from './planner-fields.ts';

/*
 * What the undo of a Planner change compares and writes back. A PATCH body
 * names the fields that it changes; a field that is a map (assignments,
 * labels, checklist, references, label names) names only the entries that it
 * changes. `snapshotOf` gives the values of these fields and entries of an
 * item, so the undo check compares the item of a later read with the item
 * that the write gave. `restoreOf` gives the PATCH that writes the values of
 * before back. The HTML `notes` are compared and written back by their plain
 * text `description`: Graph v1.0 reads no `notes`, and Planner keeps the two
 * the same.
 */

type Fields = Readonly<Record<string, unknown>>;

const isRecord = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value);

const mapOf = (value: unknown): Record<string, unknown> => (isRecord(value) ? value : {});

const textOrNull = (value: unknown): string | null => (typeof value === 'string' && value !== '' ? value : null);

/** The title and the state of a checklist item that Graph gives, or undefined. */
export const checklistItemOf = (value: unknown): ChecklistItem | undefined => {
  const item = mapOf(value);
  return typeof item['title'] === 'string' && typeof item['isChecked'] === 'boolean' ? { title: item['title'], isChecked: item['isChecked'] } : undefined;
};

/** The name and the type of a reference that Graph gives, or undefined when there is none. */
export type ReferenceValue = { readonly alias: string | null; readonly type: string | null };

export const referenceOf = (value: unknown): ReferenceValue | undefined => (isRecord(value) ? { alias: textOrNull(value['alias']), type: textOrNull(value['type']) } : undefined);

/** The PATCH value of a reference. */
export const referenceValue = (reference: ReferenceValue): Record<string, unknown> => ({
  '@odata.type': '#microsoft.graph.plannerExternalReference',
  ...(reference.alias === null ? {} : { alias: reference.alias }),
  ...(reference.type === null ? {} : { type: reference.type }),
});

type Entry = { readonly compared: (value: unknown) => unknown; readonly restored: (value: unknown) => unknown };

const ENTRIES: Readonly<Record<string, Entry>> = {
  assignments: { compared: isRecord, restored: (value) => (isRecord(value) ? ASSIGNMENT : null) },
  appliedCategories: { compared: (value) => value === true, restored: (value) => value === true },
  checklist: {
    compared: (value) => checklistItemOf(value) ?? null,
    restored: (value) => {
      const item = checklistItemOf(value);
      return item === undefined ? null : checklistValue(item);
    },
  },
  references: {
    compared: (value) => referenceOf(value) ?? null,
    restored: (value) => {
      const reference = referenceOf(value);
      return reference === undefined ? null : referenceValue(reference);
    },
  },
  categoryDescriptions: { compared: textOrNull, restored: textOrNull },
};

// The value of a plain field when the item has none: an empty description, the automatic preview.
const DEFAULTS: Fields = { description: '', previewType: 'automatic' };

const plainOf = (fields: Fields, key: string): unknown => fields[key] ?? DEFAULTS[key] ?? null;

const valuesOf = (fields: Fields, shape: Fields, pick: (entry: Entry) => (value: unknown) => unknown): Record<string, unknown> =>
  Object.fromEntries(
    Object.keys(shape).map((key) => {
      if (key === 'notes') return ['description', plainOf(fields, 'description')];
      const entry = ENTRIES[key];
      if (entry === undefined) return [key, plainOf(fields, key)];
      const now = mapOf(fields[key]);
      return [key, Object.fromEntries(Object.keys(mapOf(shape[key])).map((sub) => [sub, pick(entry)(now[sub])]))];
    })
  );

/** The values of the item for the fields and entries of `shape` (a PATCH body, or an earlier snapshot). */
export const snapshotOf = (fields: Fields, shape: Fields): Record<string, unknown> => valuesOf(fields, shape, (entry) => entry.compared);

/** The PATCH that gives the fields and entries of `body` the values that the item `fields` has. */
export const restoreOf = (fields: Fields, body: Fields): Record<string, unknown> => valuesOf(fields, body, (entry) => entry.restored);

/**
 * The checklist of a restore PATCH, with a new key for each item that is to
 * come back but is no longer in the checklist: Planner does not take the key
 * of a removed item again in a sure way, so the item comes back with a new key.
 */
export const withNewKeys = (checklist: Fields, now: Fields, newKey: () => string): Record<string, unknown> =>
  Object.fromEntries(Object.entries(checklist).map(([key, value]) => [value !== null && !isRecord(now[key]) ? newKey() : key, value]));
