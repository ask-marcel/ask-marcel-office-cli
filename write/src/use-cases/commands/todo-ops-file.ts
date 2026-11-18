import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import { parseJson } from '@ask-marcel/office-core/domain/json.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { LocalFileReader } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import { readLocalFile } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import type { TaskInput } from '../../domain/todo-fields.ts';
import { isDay, isMinute, TASK_IMPORTANCES, TASK_STATUSES } from '../../domain/todo-fields.ts';

/*
 * The ops file of the batch command: `{ "ops": [ ... ] }`, the form of the edit
 * script. An op adds, updates, moves or deletes a task. A refusal names the op
 * and the field, and never quotes the file: over MCP the file can be any
 * local file, and a quote would show its content.
 */

/** The largest ops file. */
export const OPS_FILE_MAX_BYTES = 1_000_000;

/** The most ops in one file: each op is one Graph call or more. */
export const MAX_OPS = 100;

/** The most new steps in one update op. */
export const MAX_NEW_STEPS = 20;

const day = z.union([z.string().refine(isDay, 'must be a day in the form YYYY-MM-DD, or null'), z.null()]);
const minute = z.union([z.string().refine(isMinute, 'must be a day and a time in the form YYYY-MM-DDTHH:mm, or null'), z.null()]);
const text = z.string().trim().min(1, 'must not be empty');

const fields = {
  title: text,
  body: z.string(),
  status: z.enum(TASK_STATUSES),
  importance: z.enum(TASK_IMPORTANCES),
  due: day,
  start: day,
  reminder: minute,
  categories: z.array(z.string()),
};

const setSchema = z.strictObject(fields).partial();

const opSchema = z.discriminatedUnion('op', [
  z.strictObject({ ...fields, op: z.literal('add'), list: text }).partial({ body: true, status: true, importance: true, due: true, start: true, reminder: true, categories: true }),
  z.strictObject({
    op: z.literal('update'),
    list: text,
    id: z.string(),
    set: setSchema.optional(),
    addChecklist: z.array(text.max(255, 'must have 255 characters or fewer')).max(MAX_NEW_STEPS).optional(),
  }),
  z.strictObject({ op: z.literal('move'), list: text, id: z.string(), to: text }),
  z.strictObject({ op: z.literal('delete'), list: text, id: z.string() }),
]);

const fileSchema = z.object({ ops: z.array(opSchema).min(1, 'must have one op or more').max(MAX_OPS, `must have ${MAX_OPS} ops or fewer`) });

export type TodoOpInput = z.infer<typeof opSchema>;

const refused = (message: string): Result<never, GraphError> => err({ type: 'validation_error', code: 'invalid_ops_file', message: `--file: ${message}` });

// The place of an issue as a person reads it: `op 3 due` for ops[2].due.
const placeOf = (path: ReadonlyArray<PropertyKey>): string => {
  const parts = path.map((part, index) => (typeof part === 'number' && path[index - 1] === 'ops' ? `op ${part + 1}` : String(part)));
  const words = parts.filter((part, index) => !(part === 'ops' && typeof path[index + 1] === 'number'));
  return words.length === 0 ? 'the file' : words.join(' ');
};

// The message of an issue, with no value or key of the file in it.
const messageOf = (issue: z.core.$ZodIssue): string => {
  if (issue.code === 'unrecognized_keys') return 'has a field that apply-todo-ops does not know';
  if (issue.code === 'invalid_union') return 'has no "op" that apply-todo-ops knows: add, update, move or delete';
  if (issue.code === 'invalid_type') return issue.message.endsWith('received undefined') ? 'is missing' : `must be of the type ${issue.expected}`;
  return issue.message;
};

/** The ops of the file at `path`, checked. */
export const readOpsFile = async (files: LocalFileReader | undefined, path: string): Promise<Result<ReadonlyArray<TodoOpInput>, GraphError>> => {
  const read = await readLocalFile(files, path, OPS_FILE_MAX_BYTES);
  if (!read.ok) return read;
  const json = parseJson(new TextDecoder().decode(read.value.bytes).replace(/^\uFEFF/, ''));
  if (!json.ok) return refused('the file is not valid JSON. Write it as UTF-8, for example with a file tool or Out-File -Encoding utf8.');
  const parsed = fileSchema.safeParse(json.value);
  if (parsed.success) return ok(parsed.data.ops);
  const issue = parsed.error.issues[0];
  return refused(issue === undefined ? 'the file is not an ops file.' : `${placeOf(issue.path)} ${messageOf(issue)}.`);
};

/** The input of the fields of an add op or of the `set` of an update op. */
export const fileInputOf = (op: Partial<Record<keyof typeof fields, unknown>>): TaskInput =>
  Object.fromEntries(Object.keys(fields).flatMap((key) => (op[key as keyof typeof fields] === undefined ? [] : [[key, op[key as keyof typeof fields]]])));
