import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { GraphReads, PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';
import { isGraphId } from '../../domain/graph-id.ts';
import { needsTimeZone, taskFieldsOf } from '../../domain/todo-fields.ts';
import { createExecute, parseParams } from './mail-state-command.ts';
import type { TodoOp } from './todo-batch.ts';
import { todoPlan } from './todo-batch.ts';
import { TODO_RESPONSE } from './todo-command.ts';
import { readChildren } from './todo-copy.ts';
import type { TodoOpInput } from './todo-ops-file.ts';
import { fileInputOf, MAX_OPS, readOpsFile } from './todo-ops-file.ts';
import type { TodoList, TodoTask } from './todo-reader.ts';
import { readTask, readTimeZone, resolveList } from './todo-reader.ts';
import { createTaskOp, deleteTaskOp, moveTaskOp, updateTaskOp } from './todo-task-ops.ts';

/*
 * Applies the ops of a local JSON file: add, update (fields and new steps),
 * move and delete of tasks, in the order of the file. The plan reads each list
 * and task first; one wrong op refuses the whole file, and nothing changes. At
 * the write, an op whose task changed since the plan is passed over, and an op
 * that fails does not stop the next ones.
 */

const schema = z.object({ file: z.string().min(1) });

type Context = { readonly reads: GraphReads; readonly zone: string | undefined; readonly lists: Map<string, Result<TodoList, GraphError>> };

const refused = (place: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code: 'invalid_ops_file', message: `--file: ${place} ${message}` });

const listOf = async (context: Context, spec: string, place: string): Promise<Result<TodoList, GraphError>> => {
  const known = context.lists.get(spec);
  if (known !== undefined) return known;
  const list = await resolveList(context.reads, spec, place);
  context.lists.set(spec, list);
  return list;
};

// The task of an op. The id of the file is checked, and never quoted.
const taskOf = async (context: Context, list: TodoList, id: string, place: string): Promise<Result<TodoTask, GraphError>> => {
  if (!isGraphId(id)) return refused(place, 'id is not a Graph id. Give the id of a task, as list-todo-tasks returns it.');
  return readTask(context.reads, list.id, id);
};

const OP_PLANS: {
  readonly [K in TodoOpInput['op']]: (context: Context, op: Extract<TodoOpInput, { op: K }>, list: TodoList, place: string) => Promise<Result<TodoOp, GraphError>>;
} = {
  add: async (context, op, list) => ok(createTaskOp(list, taskFieldsOf(fileInputOf(op), context.zone))),
  update: async (context, op, list, place) => {
    const fields = taskFieldsOf(fileInputOf(op.set ?? {}), context.zone);
    if (Object.keys(fields).length === 0 && (op.addChecklist ?? []).length === 0) return refused(place, 'needs "set", "addChecklist", or both.');
    const task = await taskOf(context, list, op.id, place);
    return task.ok ? ok(updateTaskOp('update-task', list, task.value, fields, op.addChecklist)) : task;
  },
  move: async (context, op, list, place) => {
    const to = await listOf(context, op.to, `${place} to`);
    if (!to.ok) return to;
    if (to.value.id === list.id) return refused(place, 'moves the task to the list it is in.');
    const task = await taskOf(context, list, op.id, place);
    if (!task.ok) return task;
    const children = await readChildren(context.reads, task.value);
    return children.ok ? ok(moveTaskOp(list, task.value, to.value, children.value)) : children;
  },
  delete: async (context, op, list, place) => {
    const task = await taskOf(context, list, op.id, place);
    if (!task.ok) return task;
    const children = await readChildren(context.reads, task.value);
    return children.ok ? ok(deleteTaskOp(list, task.value, children.value)) : children;
  },
};

const planOp = async (context: Context, op: TodoOpInput, place: string): Promise<Result<TodoOp, GraphError>> => {
  const list = await listOf(context, op.list, `${place} list`);
  if (!list.ok) return list;
  return (OP_PLANS[op.op] as (context: Context, op: TodoOpInput, list: TodoList, place: string) => Promise<Result<TodoOp, GraphError>>)(context, op, list.value, place);
};

// A task in two ops would be changed on a state that the plan did not read.
const twice = (ops: ReadonlyArray<TodoOpInput>): Result<undefined, GraphError> => {
  const seen = new Map<string, number>();
  for (const [index, op] of ops.entries()) {
    if (op.op === 'add') continue;
    const first = seen.get(op.id);
    if (first !== undefined) return refused(`op ${index + 1}`, `names the task of op ${first + 1}. Give each task in one op only.`);
    seen.set(op.id, index);
  }
  return ok(undefined);
};

// The fields that an op sets: those of an add, or the `set` of an update.
const fieldsInOp = (op: TodoOpInput): Parameters<typeof fileInputOf>[0] => {
  if (op.op === 'add') return op;
  return op.op === 'update' ? (op.set ?? {}) : {};
};

const zoneFor = async (reads: GraphReads, ops: ReadonlyArray<TodoOpInput>): Promise<Result<string | undefined, GraphError>> =>
  ops.some((op) => needsTimeZone(fileInputOf(fieldsInOp(op)))) ? readTimeZone(reads) : ok(undefined);

const plan: PlanWrite = async (reads, params, files) => {
  const parsed = parseParams(schema, params);
  if (!parsed.ok) return parsed;
  const ops = await readOpsFile(files, parsed.value.file);
  if (!ops.ok) return ops;
  const once = twice(ops.value);
  if (!once.ok) return once;
  const zone = await zoneFor(reads, ops.value);
  if (!zone.ok) return zone;
  const context: Context = { reads, zone: zone.value, lists: new Map() };
  const planned: TodoOp[] = [];
  for (const [index, op] of ops.value.entries()) {
    const one = await planOp(context, op, `op ${index + 1}`);
    if (!one.ok) return one;
    planned.push(one.value);
  }
  return ok(todoPlan({ ops: planned }));
};

const execute = createExecute(plan);

const meta: WriteCommandMeta = {
  summary: `Apply a batch of To Do ops from a local JSON file: add, update, move and delete tasks, at most ${MAX_OPS} ops. One wrong op refuses the whole file before any change. Over MCP it runs on run-destructive-command, because an op can move or delete a task.`,
  category: 'tasks',
  graphMethod: 'POST',
  graphPathTemplate: '/me/todo/lists/<list id>/tasks (and PATCH, DELETE on /me/todo/lists/<list id>/tasks/<task id>)',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/resources/todotask',
  options: [
    {
      name: 'file',
      key: 'file',
      required: true,
      description:
        'A local JSON file, UTF-8, at most 1 MB: `{ "ops": [ { "op": "add", "list": "Tasks", "title": "...", "due": "2026-11-20" }, { "op": "update", "list": "Tasks", "id": "<task id>", "set": { "status": "completed" }, "addChecklist": ["..."] }, { "op": "move", "list": "Tasks", "id": "<task id>", "to": "Later" }, { "op": "delete", "list": "Tasks", "id": "<task id>" } ] }`. The fields of add and set: title, body, status, importance, due and start (YYYY-MM-DD or null), reminder (YYYY-MM-DDTHH:mm or null), categories (a list). A task can be in one op only.',
    },
  ],
  example: '{write-bin} apply-todo-ops --file ./ops.json --dry-run',
  bodyTemplate: 'for each op of {file}: the body of create-todo-task, update-todo-task, add-todo-task-step, move-todo-task or delete-todo-task',
  effect: 'delete',
  undo: 'reverses each change of the run as the single commands do: it deletes an added task, writes the fields of before back, deletes the new steps, moves a task back, and creates a deleted task again (with a new id)',
  scopesRequired: ['Tasks.ReadWrite', 'MailboxSettings.Read'],
  responseShape: TODO_RESPONSE,
};

export { execute, meta, plan, schema };
