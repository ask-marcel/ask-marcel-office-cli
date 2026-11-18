import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads, PlannedChange } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { isGraphId } from '../../domain/graph-id.ts';
import type { TodoOp } from './todo-batch.ts';
import { outcomeOf } from './todo-batch.ts';
import type { TodoList } from './todo-reader.ts';
import { unreadable } from './todo-reader.ts';

/*
 * The ops on a To Do list: create one, and change its name. No command of this
 * CLI deletes a list, so the undo of a create keeps the list (see the item undo handlers).
 */

const listSchema = z.object({ id: z.string(), displayName: z.string() });

/** The name of the list as Graph gives it now. */
export const readListName = async (reads: GraphReads, listId: string): Promise<Result<string, GraphError>> => {
  const read = await reads.get(`/me/todo/lists/${listId}`);
  if (!read.ok) return read;
  const list = listSchema.safeParse(read.value);
  return list.success ? ok(list.data.displayName) : err(unreadable('a To Do list'));
};

/** Creates a list with the name `name`. */
export const createListOp = (name: string): TodoOp => {
  const change: PlannedChange = { action: 'create-list', target: '/me/todo/lists', before: {} };
  return {
    change,
    write: async (graph) => {
      const created = await graph.post(change.target, { displayName: name });
      if (!created.ok) return outcomeOf(created);
      const list = listSchema.safeParse(created.value);
      if (!list.success || !isGraphId(list.data.id)) return outcomeOf(err(unreadable('the list it created')));
      return outcomeOf(ok({ ...change, after: { listId: list.data.id, list: list.data.displayName }, undo: { kind: 'todo-list-created', data: { listId: list.data.id } } }));
    },
  };
};

/** Gives the list the name `name`. */
export const renameListOp = (list: TodoList, name: string): TodoOp => {
  const change: PlannedChange = { action: 'rename-list', target: `/me/todo/lists/${list.id}`, before: { listId: list.id, list: list.displayName } };
  return {
    change,
    recheck: async (reads) => {
      const now = await readListName(reads, list.id);
      return now.ok ? ok(now.value === list.displayName) : now;
    },
    write: async (graph) => {
      const renamed = await graph.patch(change.target, { displayName: name });
      if (!renamed.ok) return outcomeOf(renamed);
      const now = listSchema.safeParse(renamed.value);
      if (!now.success) return outcomeOf(err(unreadable('the list it changed')));
      const data = { listId: list.id, before: list.displayName, set: now.data.displayName };
      return outcomeOf(ok({ ...change, after: { listId: list.id, list: now.data.displayName }, undo: { kind: 'restore-todo-list-name', data } }));
    },
  };
};
