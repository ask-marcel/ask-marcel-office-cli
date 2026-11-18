import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { GraphReads } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { isGraphId, parseGraphId } from '../../domain/graph-id.ts';
import { isNotThere } from './mail-folder-reader.ts';

/*
 * The reads of To Do that the To Do commands need: a list by its id or its
 * name, a task, and the time zone of the mailbox. A list name never goes into
 * a Graph path: the plan reads the lists and keeps the id that Graph gave.
 */

export type TodoList = { readonly id: GraphId; readonly displayName: string };

/** A task as the plan read it. `stamp` is its `lastModifiedDateTime`: a change of the task changes it. */
export type TodoTask = {
  readonly listId: GraphId;
  readonly id: GraphId;
  readonly title: string;
  readonly status: string;
  readonly stamp: string;
  readonly fields: Readonly<Record<string, unknown>>;
};

/** The Graph path of a task. */
export const taskPath = (listId: string, taskId: string): string => `/me/todo/lists/${listId}/tasks/${taskId}`;

/** The most lists that one read gives; a name in a later page is not found. */
export const LIST_PAGE = 200;

const listSchema = z.object({ id: z.string(), displayName: z.string() });
const listsSchema = z.object({ value: z.array(listSchema), '@odata.nextLink': z.string().optional() });
const taskSchema = z.object({ id: z.string(), title: z.string(), status: z.string(), lastModifiedDateTime: z.string() }).loose();
const timeZoneSchema = z.object({ value: z.string().min(1) });

/** The refusal for an answer of Graph that the CLI cannot read. */
export const unreadable = (what: string): GraphError => ({ type: 'validation_error', message: `Graph gave ${what} in a shape that this CLI cannot read.` });

const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

const noList = (what: string): Result<never, GraphError> =>
  refused('todo_list_not_found', `${what}: no To Do list has this name or id. list-todo-task-lists gives the name and the id of each list.`);

const listOf = (raw: unknown): Result<TodoList, GraphError> => {
  const list = listSchema.safeParse(raw);
  if (!list.success || !isGraphId(list.data.id)) return err(unreadable('a To Do list'));
  return ok({ id: list.data.id, displayName: list.data.displayName });
};

// The list that has the id `spec` and is not in the first page of the lists.
const listById = async (reads: GraphReads, spec: string, what: string): Promise<Result<TodoList, GraphError>> => {
  if (!isGraphId(spec)) return noList(what);
  const read = await reads.get(`/me/todo/lists/${spec}`);
  if (!read.ok) return isNotThere(read.error) ? noList(what) : read;
  return listOf(read.value);
};

/** The list with the id or the name (compared with no case) `spec`. `what` names the param in a refusal, for example `--list`. */
export const resolveList = async (reads: GraphReads, spec: string, what: string): Promise<Result<TodoList, GraphError>> => {
  const read = await reads.get(`/me/todo/lists?$top=${LIST_PAGE}`);
  if (!read.ok) return read;
  const page = listsSchema.safeParse(read.value);
  if (!page.success) return err(unreadable('the To Do lists'));
  const byId = page.data.value.find((list) => list.id === spec);
  if (byId !== undefined) return listOf(byId);
  const named = page.data.value.filter((list) => list.displayName.toLowerCase() === spec.toLowerCase());
  if (named.length > 1) return refused('todo_list_ambiguous', `${what}: ${named.length} To Do lists have this name. Give the id of the list, as list-todo-task-lists returns it.`);
  return named.length === 1 ? listOf(named[0]) : listById(reads, spec, what);
};

/** The task `taskId` of the list, as Graph gives it now. */
export const readTask = async (reads: GraphReads, listId: GraphId, taskId: GraphId): Promise<Result<TodoTask, GraphError>> => {
  const read = await reads.get(taskPath(listId, taskId));
  if (!read.ok) return read;
  const task = taskSchema.safeParse(read.value);
  if (!task.success) return err(unreadable('a To Do task'));
  return ok({ listId, id: taskId, title: task.data.title, status: task.data.status, stamp: task.data.lastModifiedDateTime, fields: task.data });
};

/** The task named by `raw` in the list, after the check of the id. */
export const readTaskById = async (reads: GraphReads, list: TodoList, raw: string, what: string): Promise<Result<TodoTask, GraphError>> => {
  const id = parseGraphId(raw, what);
  return id.ok ? readTask(reads, list.id, id.value) : id;
};

/** The time zone of the mailbox (`Singapore Standard Time`, `Europe/Paris`), in which To Do shows the dates. */
export const readTimeZone = async (reads: GraphReads): Promise<Result<string, GraphError>> => {
  const read = await reads.get('/me/mailboxSettings/timeZone');
  if (!read.ok) return read;
  const zone = timeZoneSchema.safeParse(read.value);
  return zone.success ? ok(zone.data.value) : err(unreadable('the time zone of the mailbox'));
};
