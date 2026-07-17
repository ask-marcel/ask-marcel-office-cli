/*
 * The one place a registry command actually runs.
 *
 * Both front ends call this: `cli.ts` (commander) and `mcp.ts` (the MCP
 * gateway's run-command / run-write-command tools). Everything between "I have
 * a resolved command and a bag of params" and "I have a value or a failure to
 * render" lives here, so a fix reaches both surfaces at once.
 *
 * That sharing is the point. LESSONS 2026-06-13 records an additive envelope
 * field that had to be threaded through three layers and reached library
 * consumers but not CLI ones; LESSONS 2026-07-16 records a param threaded
 * through N call sites that silently missed one, with a green suite the whole
 * time. A second front end re-implementing this block would repeat both.
 *
 * Lives in `composition` rather than `use-cases`: it depends on `ErrorSource`
 * (a presenter type), and there are zero use-cases -> presenter imports today.
 *
 * Deliberately NOT here: rendering. This returns a value; the caller decides
 * whether it becomes stdout (CLI) or MCP tool content.
 */
import type { GraphClient, GraphError } from '../infra/graph-client.ts';
import type { ErrorSource } from '../presenter/error-hints.ts';
import type { Command } from '../use-cases/commands/command-types.ts';
import { commands as cmdRegistry } from '../use-cases/commands/index.ts';
import type { OutputDirError, OutputPathError } from '../use-cases/commands/output-path.ts';
import { persistIfRequested, persistMediaIfRequested } from '../use-cases/commands/output-path.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';

// map the discriminated GraphError type to
// the presenter's ErrorSource. `api_error` and `network_error` both surface
// Graph endpoint failures (the latter is "couldn't even reach Graph"), so
// both classify as `graph`. `validation_error` is Zod / use-case-side
// schema rejection. `auth_failed` is a CLI-side concern (the remedy is
// `ask-marcel-office login`, not a Graph operation). When the hint table matches
// a more specific source (e.g. `substrate` for chatsvcagg-coded errors),
// the hint's source wins — this explicit source is just the fallback so
// bare envelopes still carry the classifier.
const sourceFromGraphError = (error: GraphError): ErrorSource => {
  if (error.type === 'validation_error') return 'validation';
  if (error.type === 'auth_failed') return 'cli';
  return 'graph';
};

// manifest-driven --output-path-supporting list.
const bytesProducingCommands = Object.entries(cmdRegistry)
  .filter(([, c]) => c.meta.producesBytes === true)
  .map(([n]) => n)
  .toSorted((a, b) => a.localeCompare(b));

// Parallel manifest-driven list for the --output-dir flag.
const mediaProducingCommands = Object.entries(cmdRegistry)
  .filter(([, c]) => c.meta.producesMedia === true)
  .map(([n]) => n)
  .toSorted((a, b) => a.localeCompare(b));

const formatOutputPathError = (error: OutputPathError, commandName: string): string => {
  if (error.type === 'no_inlined_bytes')
    return `--output-path: ${commandName} did not return inlined bytes — this flag works only with commands that produce a body to write. Supported: ${bytesProducingCommands.join(', ')}. Plain JSON commands (list-*, get-*-user, get-organization, etc.) don't have a body to write — drop the flag and use a shell redirect instead: \`ask-marcel-office ${commandName} ... > out.json\`.`;
  if (error.type === 'empty_path') return '--output-path: path argument is empty (likely a shell-quoting mistake — pass a real filesystem path)';
  // paths ending in `/` or `\` look like a directory; reject upfront instead of Node's `EISDIR`.
  if (error.type === 'is_directory') return '--output-path: must be a file path, not a directory (paths ending in `/` or `\\` look like a directory).';
  // *-as-pdf fallbacks return source bytes with `passthrough:true`; refuse `.pdf` to avoid a corrupt save.
  if (error.type === 'passthrough_extension_mismatch')
    return `--output-path: response is passthrough source bytes (contentType: \`${error.contentType}\`), NOT a converted PDF. Save with the source extension matching that contentType, not \`${error.requestedExtension}\` — see the response's \`note\` field.`;
  // Large-payload guard: refuse to dump a multi-MB base64 blob to stdout (context bomb). Point at --output-path.
  if (error.type === 'inline_too_large')
    return `--output-path: ${commandName} returned a ~${(error.base64Length * 0.75e-6).toFixed(1)} MB inline payload — too large to print to stdout (it would flood the context). Re-run with \`--output-path <file>\` to write the bytes to disk; the envelope then carries \`savedTo\` instead of \`base64\`.`;
  // humanise the ENOENT/mkdir shape; preserve EACCES/ENOSPC verbatim — they're already actionable.
  const enoent = /^ENOENT:.*'([^']+)'/.exec(error.message);
  if (enoent !== null) return `--output-path: parent directory missing or not writable: ${enoent[1]}`;
  return `--output-path: write failed: ${error.message}`;
};

const formatOutputDirError = (error: OutputDirError, commandName: string): string => {
  if (error.type === 'no_media')
    return `--output-dir: ${commandName} did not return a media array — this flag works only with the image-extraction commands. Supported: ${mediaProducingCommands.join(', ')}.`;
  if (error.type === 'empty_path') return '--output-dir: directory argument is empty (likely a shell-quoting mistake — pass a real directory path)';
  return `--output-dir: write failed: ${error.message}`;
};

export type RunRegistryCommandDeps = {
  readonly graph: GraphClient;
  readonly fs: FileSystem;
};

export type RunRegistryCommandRequest = {
  /** Canonical registry name — drives the manifest-derived error messages. */
  readonly name: string;
  readonly command: Command;
  /** Raw params, still possibly keyed by an option ALIAS. */
  readonly params: Record<string, string>;
  readonly outputPath?: string;
  readonly outputDir?: string;
};

/**
 * `source` / `code` / `retryAfterSeconds` are optional on purpose: the
 * --output-path and --output-dir rejections carry none of them, and stamping a
 * source they never had would change the shipped error envelope.
 */
export type RunRegistryCommandFailure = {
  readonly message: string;
  readonly code?: string;
  readonly source?: ErrorSource;
  readonly retryAfterSeconds?: number;
};

/**
 * Track which alias the user actually typed so a post-validation error message
 * references the flag they typed rather than the canonical schema name (e.g.
 * `--query is empty` not `--title-substring is empty`).
 */
const normalizeAliases = (command: Command, params: Record<string, string>): { readonly normalized: Record<string, string>; readonly aliasUsedFor: Record<string, string> } => {
  const normalized: Record<string, string> = { ...params };
  const aliasUsedFor: Record<string, string> = {};
  for (const opt of command.meta.options) {
    for (const alias of opt.aliases ?? []) {
      const aliasValue = params[alias.key];
      if (typeof aliasValue === 'string' && !Object.hasOwn(params, opt.key)) {
        normalized[opt.key] = aliasValue;
        aliasUsedFor[opt.name] = alias.name;
      }
    }
  }
  return { normalized, aliasUsedFor };
};

const toFailure = (command: Command, error: GraphError, aliasUsedFor: Record<string, string>): RunRegistryCommandFailure => {
  let message = error.message;
  for (const [canonical, alias] of Object.entries(aliasUsedFor)) {
    message = message.replaceAll(`--${canonical}`, `--${alias}`);
  }
  // 2026-06-15 (F-02): a local-filesystem command (convert-local-file-to-markdown,
  // extract-local-file-images) never touches Graph, so its runtime
  // failures must not be stamped `source: graph` (misleading — an LLM
  // might treat a missing local file as a transient Graph error and
  // retry). Validation errors still flow through the normal classifier.
  const isLocalCommand = command.executeLocal !== undefined;
  return {
    message,
    ...(error.code === undefined ? {} : { code: error.code }),
    source: isLocalCommand && error.type !== 'validation_error' ? 'cli' : sourceFromGraphError(error),
    ...(error.type === 'api_error' && error.retryAfterSeconds !== undefined ? { retryAfterSeconds: error.retryAfterSeconds } : {}),
  };
};

export const runRegistryCommand = async (deps: RunRegistryCommandDeps, request: RunRegistryCommandRequest): Promise<Result<unknown, RunRegistryCommandFailure>> => {
  const { command, name } = request;
  const { normalized, aliasUsedFor } = normalizeAliases(command, request.params);
  // `convert-local-file-to-markdown` / `extract-local-file-images` are the commands whose
  // input is the local filesystem, not Graph — route them to executeLocal with the
  // composition-selected FileSystem (the same instance --output-path uses).
  const result = command.executeLocal !== undefined ? await command.executeLocal(deps.fs, normalized) : await command.execute(deps.graph, normalized);
  if (!result.ok) return err(toFailure(command, result.error, aliasUsedFor));

  // --output-dir is checked BEFORE --output-path: a media-producing command
  // paired with both should land the media array, not fall into the
  // single-file writer. Order is behaviour, not style.
  if (request.outputDir !== undefined) {
    const persistedMedia = await persistMediaIfRequested(deps.fs, request.outputDir, result.value);
    if (persistedMedia.ok) return ok(persistedMedia.value);
    return err({ message: formatOutputDirError(persistedMedia.error, name) });
  }
  const persisted = await persistIfRequested(deps.fs, request.outputPath, result.value);
  if (persisted.ok) return ok(persisted.value);
  return err({ message: formatOutputPathError(persisted.error, name) });
};

export { formatOutputDirError, formatOutputPathError };
