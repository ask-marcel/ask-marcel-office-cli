import type { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { UndoCheck, UndoHandler } from '@ask-marcel/office-core/use-cases/commands/undo-run.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { WriteGraph } from '@ask-marcel/office-core/use-cases/ports/graph.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { isNotThere } from './mail-folder-reader.ts';
import type { TodoTask } from './todo-reader.ts';
import { readTask } from './todo-reader.ts';

/*
 * What the undo handlers of To Do share: the check of the undo data that the
 * run journal keeps (a person can edit that file), the compare of two values,
 * and the reads of a task and of a list that a check makes.
 */

/** The refusal for undo data that this CLI did not write. */
export const badJournal: GraphError = {
  type: 'validation_error',
  message: 'The run journal names an id that is not a Graph id, or data that this CLI did not write, so this undo does not run.',
};

/** The JSON of a value with the keys of each object in order, so two equal values compare equal. */
export const canonical = (value: unknown): string =>
  JSON.stringify(value, (_key, item: unknown) =>
    item !== null && typeof item === 'object' && !Array.isArray(item) ? Object.fromEntries(Object.entries(item).toSorted(([a], [b]) => a.localeCompare(b))) : item
  );

type Check<T> = (reads: GraphReads, data: T) => Promise<Result<UndoCheck, GraphError>>;
type Apply<T> = (graph: WriteGraph, data: T) => Promise<Result<unknown, GraphError>>;

/** An undo handler whose data has the shape `schema`: data of another shape is refused, never sent to Graph. */
export const createTodoUndoHandler = <T>(schema: z.ZodType<T>, check: Check<T>, apply: Apply<T>): UndoHandler => ({
  check: async (reads, data) => {
    const step = schema.safeParse(data);
    return step.success ? check(reads, step.data) : err(badJournal);
  },
  apply: async (graph, data) => {
    const step = schema.safeParse(data);
    return step.success ? apply(graph, step.data) : err(badJournal);
  },
});

/** The check of an item that a read did not give: `missing` when it is not there, else the error. */
export const notThere = (error: GraphError): Result<UndoCheck, GraphError> => (isNotThere(error) ? ok('missing') : err(error));

/** The check of an item that a read gives: `missing` when it is not there, else the outcome of `compare`. */
export const checkRead = async <T>(read: Result<T, GraphError>, compare: (item: T) => boolean): Promise<Result<UndoCheck, GraphError>> => {
  if (!read.ok) return notThere(read.error);
  return ok(compare(read.value) ? 'ready' : 'changed');
};

/** The check of the task: missing, or the outcome of `compare` on it. */
export const checkTask = async (reads: GraphReads, listId: GraphId, taskId: GraphId, compare: (task: TodoTask) => boolean): Promise<Result<UndoCheck, GraphError>> =>
  checkRead(await readTask(reads, listId, taskId), compare);

/** `ready` when the list is there, else `missing`. */
export const listThere = async (reads: GraphReads, listId: GraphId): Promise<Result<UndoCheck, GraphError>> => checkRead(await reads.get(`/me/todo/lists/${listId}`), () => true);
