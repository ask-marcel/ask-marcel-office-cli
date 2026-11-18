import { z } from 'zod';
import type { GraphError } from '@ask-marcel/office-core/domain/graph-error.ts';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import { err, ok } from '@ask-marcel/office-core/domain/result.ts';
import { decodeUtf8 } from '@ask-marcel/office-core/domain/utf8.ts';
import type { CommandOptionMeta, WriteCommand } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import type { LocalFileReader } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import { readLocalFile } from '@ask-marcel/office-core/use-cases/commands/read-local-file.ts';
import type { PlanWrite } from '@ask-marcel/office-core/use-cases/commands/write-plan.ts';

/*
 * `--body-file` on the four draft commands: the body text comes from a local
 * file, for a body too long for a command line. The file is read before the
 * checks of the command, in the dry-run too, and its text goes into the text
 * option of the command (`--body-content` or `--comment`). So the checks, the
 * writes and the undo of each command stay as they are. The text is UTF-8: a
 * byte order mark is removed, and a file that is not UTF-8 is refused.
 */

/** The largest body file, in bytes. Graph takes a request of at most 4 MB, and JSON makes the text longer. */
export const BODY_FILE_MAX_BYTES = 1_000_000;

/** The text options that `--body-file` can stand in for. */
export type BodyKey = 'bodyContent' | 'comment';

/** How `--body-file` goes into a command. */
export type BodyFileSpec = {
  /** The option that gets the text of the file. */
  readonly key: BodyKey;
  /** True when the command needs a body: then one of `key` and `--body-file` must be given. */
  readonly required: boolean;
  /** The text options that cannot go with `--body-file`. */
  readonly exclusive: ReadonlyArray<BodyKey>;
};

const FLAGS: Readonly<Record<BodyKey, string>> = { bodyContent: '--body-content', comment: '--comment' };

/** The field of `--body-file` in the schema of a draft command: the exported schema names the params, and the file read checks the path. */
export const BODY_FILE_FIELD = z.string().optional();

/** The option `--body-file` of a draft command; `stands` names the option it stands in for. */
export const bodyFileOption = (stands: string): CommandOptionMeta => ({
  name: 'body-file',
  key: 'bodyFile',
  required: false,
  description: `A local UTF-8 text file with the body text, in place of ${stands}, for a long body. At most ${BODY_FILE_MAX_BYTES} bytes. A byte order mark is removed. A relative path starts from the folder the CLI runs in. Mutually exclusive with ${stands}.`,
});

const refused = (code: string, message: string): Result<never, GraphError> => err({ type: 'validation_error', code, message });

const conflict = (key: BodyKey): Result<never, GraphError> =>
  refused('body_file_conflict', `--body-file and ${FLAGS[key]} are mutually exclusive: give the body text in one of them.`);

// The params with no body file: refused when the command needs a body and has none.
const withoutFile = (params: Record<string, string>, spec: BodyFileSpec): Result<Record<string, string>, GraphError> =>
  spec.required && params[spec.key] === undefined ? refused('body_missing', `Give the body text with ${FLAGS[spec.key]} or --body-file.`) : ok(params);

// The text of the file, never quoted in a refusal: over MCP the path can name any local file.
const textOf = (path: string, bytes: Uint8Array): Result<string, GraphError> => {
  const text = decodeUtf8(bytes);
  if (!text.ok) return refused('body_file_not_utf8', `${path} is not UTF-8 text. Save the body file as UTF-8 (a byte order mark is removed).`);
  return text.value.length === 0 ? refused('body_file_empty', `${path} has no text. Write the body text in the file.`) : ok(text.value);
};

/** The params with the text of `--body-file` in the text option of the command. The schema of the command drops `bodyFile`. */
export const resolveBodyFile = async (
  params: Record<string, string>,
  files: LocalFileReader | undefined,
  spec: BodyFileSpec
): Promise<Result<Record<string, string>, GraphError>> => {
  const path = params['bodyFile'];
  if (path === undefined) return withoutFile(params, spec);
  const other = spec.exclusive.find((key) => params[key] !== undefined);
  if (other !== undefined) return conflict(other);
  const file = await readLocalFile(files, path, BODY_FILE_MAX_BYTES);
  if (!file.ok) return file;
  const text = textOf(path, file.value.bytes);
  if (!text.ok) return text;
  return ok({ ...params, [spec.key]: text.value });
};

/** The `execute` of a draft command that takes `--body-file`, from the `execute` of its text option. */
export const createBodyFileExecute =
  (execute: WriteCommand['execute'], spec: BodyFileSpec): WriteCommand['execute'] =>
  async (graph, params, files) => {
    const resolved = await resolveBodyFile(params, files, spec);
    return resolved.ok ? execute(graph, resolved.value) : resolved;
  };

/** The plan of a draft command that takes `--body-file`: the dry-run reads the file too. */
export const createBodyFilePlan =
  (plan: PlanWrite, spec: BodyFileSpec): PlanWrite =>
  async (reads, params, files) => {
    const resolved = await resolveBodyFile(params, files, spec);
    return resolved.ok ? plan(reads, resolved.value) : resolved;
  };
