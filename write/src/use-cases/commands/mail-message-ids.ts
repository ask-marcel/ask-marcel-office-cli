import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import type { LocalFileReader } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import { readLocalFile } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import type { GraphId } from '../../domain/graph-id.ts';
import { parseGraphId } from '../../domain/graph-id.ts';

/*
 * The message ids of a batch: from `--message-id` (one id, or ids separated by
 * commas) and from `--ids-file` (a text file, one id on each line, empty lines
 * passed over). Each id is checked before it goes into a Graph path, a repeated
 * id counts once, and a batch has at most 1000 messages.
 */

/** The largest `--ids-file`: about 6000 ids. */
export const IDS_FILE_MAX_BYTES = 1_000_000;

/** The most messages that one run changes: each is one Graph call or more. */
export const MAX_BATCH_MESSAGES = 1000;

const refused = (message: string): Result<never, GraphError> => err({ type: 'validation_error', code: 'invalid_batch', message });

type RawId = { readonly raw: string; readonly what: string; readonly fromFile: boolean };

// A line of a file is not quoted in the refusal: over MCP, the file can be any local file, and a quote would show its content.
const badLine = (what: string): Result<never, GraphError> =>
  err({ type: 'validation_error', code: 'invalid_id', message: `${what} is not a Graph id. Give one message id on each line, as list-mail-folder-messages returns it.` });

const parseAll = (raws: ReadonlyArray<RawId>): Result<ReadonlyArray<GraphId>, GraphError> => {
  const ids: GraphId[] = [];
  for (const { raw, what, fromFile } of raws) {
    const id = parseGraphId(raw, what);
    if (!id.ok) return fromFile ? badLine(what) : id;
    ids.push(id.value);
  }
  return ok(ids);
};

const fromParam = (value: string | undefined): ReadonlyArray<RawId> =>
  (value ?? '')
    .split(',')
    .map((raw) => raw.trim())
    .filter((raw) => raw !== '')
    .map((raw) => ({ raw, what: '--message-id', fromFile: false }));

const fromFile = async (path: string | undefined, files: LocalFileReader | undefined): Promise<Result<ReadonlyArray<RawId>, GraphError>> => {
  if (path === undefined) return ok([]);
  const read = await readLocalFile(files, path, IDS_FILE_MAX_BYTES);
  if (!read.ok) return read;
  const lines = new TextDecoder().decode(read.value.bytes).split(/\r?\n/);
  return ok(lines.flatMap((line, index) => (line.trim() === '' ? [] : [{ raw: line.trim(), what: `--ids-file line ${index + 1}`, fromFile: true }])));
};

/** The ids of `--message-id` and `--ids-file`, in their order, each once. */
export const messageIdsOf = async (
  params: { readonly messageId?: string; readonly idsFile?: string },
  files: LocalFileReader | undefined
): Promise<Result<ReadonlyArray<GraphId>, GraphError>> => {
  const listed = await fromFile(params.idsFile, files);
  if (!listed.ok) return listed;
  const ids = parseAll([...fromParam(params.messageId), ...listed.value]);
  if (!ids.ok) return ids;
  const unique = [...new Set(ids.value)];
  if (unique.length === 0) return refused('Give the messages: --message-id <id> (ids separated by commas) or --ids-file <file> (one id on each line).');
  if (unique.length > MAX_BATCH_MESSAGES) return refused(`${unique.length} messages: one run changes at most ${MAX_BATCH_MESSAGES}. Divide the ids into more runs.`);
  return ok(unique);
};
