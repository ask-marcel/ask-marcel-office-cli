/*
 * The one place a registry command actually runs.
 *
 * Both front ends call this: `cli-shell.ts` (commander) and `mcp-shell.ts` (the
 * MCP gateway's run-command / run-write-command tools). Everything between "I have
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
 *
 * Core (package split, step 13): the commands an error names and the bin a
 * shell remedy types come from the `CommandPackage` in the deps, never from a
 * registry import.
 */
import type { GraphError } from '../infra/graph-request.ts';
import type { ReadGraph } from '../infra/read-graph.ts';
import type { WriteGraph } from '../infra/write-graph.ts';
import type { ErrorSource } from '../presenter/error-hints.ts';
import type { RenderContext, RenderSurface } from '../presenter/render-to-string.ts';
import type { CommandPackage } from '../use-cases/commands/command-package.ts';
import { isArchiveResult } from '../use-cases/commands/archive-result.ts';
import type { Command } from '../use-cases/commands/command-types.ts';
import type { OutputDirError, OutputPathError } from '../use-cases/commands/output-path.ts';
import { persistIfRequested, persistMediaIfRequested } from '../use-cases/commands/output-path.ts';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';
import type { LocalFileReader } from '../use-cases/commands/read-local-file.ts';
import type { RunRecorder } from '../use-cases/commands/record-run.ts';
import { noDryRun, previewPlan, recordRun } from '../use-cases/commands/record-run.ts';

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

// The package's own commands, not a module-level registry: this module is core
// and serves whichever package hands it its registry.
const commandNamesWhere = (registry: ErrorWordingSource['registry'], matches: (command: Command<never>) => boolean): ReadonlyArray<string> =>
  Object.entries(registry)
    .filter(([, c]) => matches(c))
    .map(([n]) => n)
    .toSorted((a, b) => a.localeCompare(b));

// manifest-driven --output-path-supporting list.
const bytesProducingCommands = (registry: ErrorWordingSource['registry']): ReadonlyArray<string> => commandNamesWhere(registry, (c) => c.meta.producesBytes === true);

// Parallel manifest-driven list for the --output-dir flag.
const mediaProducingCommands = (registry: ErrorWordingSource['registry']): ReadonlyArray<string> => commandNamesWhere(registry, (c) => c.meta.producesMedia === true);

const hasOption = (command: Command<never>, name: string): boolean => command.meta.options.some((option) => option.name === name);

/**
 * The facts the oversized-response banner needs to name a remedy this caller
 * can actually use. Derived from the manifest, so a new command gets an honest
 * banner for free. See `RenderContext` in the presenter for the why.
 */
// `tenantId` is the one context field that comes from the INVOCATION rather
// than the manifest: a Graph cursor carries no tenant, so the footer is the
// only place the next call can learn which identity signed this page.
//
// Taken as-is, deliberately. It reaches the footer UNQUOTED, which would matter
// if it were arbitrary caller text — but it cannot be. A param no command
// declares is refused before it runs, and all twelve commands that DO declare
// `--tenant-id` brand it through `routeGet` before their first Graph call, so a
// value only survives to a rendered success by already being a valid GUID.
// Re-branding here would be unreachable code guarding an invariant the command
// layer already holds; `mcp.test.ts` pins both halves of that.
const buildRenderContext = (commandName: string, command: Command<never>, surface: RenderSurface, params: Record<string, unknown>): RenderContext => {
  const raw = params['tenantId'];
  return {
    commandName,
    producesBytes: command.meta.producesBytes === true,
    supportsSelect: hasOption(command, 'select'),
    supportsTop: hasOption(command, 'top'),
    surface,
    ...(typeof raw === 'string' ? { tenantId: raw } : {}),
  };
};

// A shell redirect is the CLI's answer to "this command has no body to write".
// An MCP client has no shell, so it gets the honest version instead: there is
// no remedy at this layer, narrow the request.
const noBodyRemedy = (commandName: string, surface: RenderSurface, binName: string): string =>
  surface === 'mcp'
    ? 'Narrow the request instead (fewer items, a tighter query); there is no body to write.'
    : `Plain JSON commands (list-*, get-*-user, get-organization, etc.) don't have a body to write — drop the flag and use a shell redirect instead: \`${binName} ${commandName} ... > out.json\`.`;

/** What the error wording needs from the package: the commands it can name, and the bin a shell remedy types. */
type ErrorWordingSource = Pick<CommandPackage<Command<never>>, 'registry' | 'binName'>;

/** The --output-path refusals that name nothing of the package: all of them but the no-body refusal. */
type PackageFreeOutputPathError = Exclude<OutputPathError, { readonly type: 'no_inlined_bytes' }>;

// Large-payload guard: refuse to dump a multi-MB base64 blob to stdout (context bomb). Point at the flag
// that takes it: --output-dir for the entry images of an archive result, --output-path for a single body.
const describeInlineTooLarge = (error: Extract<OutputPathError, { readonly type: 'inline_too_large' }>, commandName: string): string => {
  const megabytes = (error.base64Length * 0.75e-6).toFixed(1);
  return error.inImages === true
    ? `--output-dir: ${commandName} returned ~${megabytes} MB of inline images — too large to print to stdout (it would flood the context). Re-run with \`--output-dir <dir>\` to write the images to disk; each image then carries \`savedTo\` instead of \`base64\`.`
    : `--output-path: ${commandName} returned a ~${megabytes} MB inline payload — too large to print to stdout (it would flood the context). Re-run with \`--output-path <file>\` to write the bytes to disk; the envelope then carries \`savedTo\` instead of \`base64\`.`;
};

const formatOutputPathError = (error: PackageFreeOutputPathError, commandName: string): string => {
  if (error.type === 'empty_path') return '--output-path: path argument is empty (likely a shell-quoting mistake — pass a real filesystem path)';
  // paths ending in `/` or `\` look like a directory; reject upfront instead of Node's `EISDIR`.
  if (error.type === 'is_directory') return '--output-path: must be a file path, not a directory (paths ending in `/` or `\\` look like a directory).';
  // A text answer (converted markdown, a PDF text layer) under a binary name is a file no reader opens.
  if (error.type === 'text_under_binary_extension')
    return `--output-path: ${commandName} returned text (\`${error.contentType}\`), not the bytes of a \`${error.requestedExtension}\`; saving it under that name would make a file no reader can open. Save it as \`.md\` or \`.txt\`, or fetch the original bytes with \`get-mail-attachment\` (a mail attachment), \`download-drive-item-content\` (a drive file) or \`download-drive-item-version --format original\` (a version).`;
  // *-as-pdf fallbacks return source bytes with `passthrough:true`; refuse `.pdf` to avoid a corrupt save.
  if (error.type === 'passthrough_extension_mismatch')
    return `--output-path: response is passthrough source bytes (contentType: \`${error.contentType}\`), NOT a converted PDF. Save with the source extension matching that contentType, not \`${error.requestedExtension}\` — see the response's \`note\` field.`;
  if (error.type === 'inline_too_large') return describeInlineTooLarge(error, commandName);
  // humanise the ENOENT/mkdir shape; preserve EACCES/ENOSPC verbatim — they're already actionable.
  const enoent = /^ENOENT:.*'([^']+)'/.exec(error.message);
  if (enoent !== null) return `--output-path: parent directory missing or not writable: ${enoent[1]}`;
  return `--output-path: write failed: ${error.message}`;
};

/** Every --output-path refusal: the no-body one names the package's bytes commands and, on the CLI, its bin. */
const describeOutputPathError = (error: OutputPathError, commandName: string, surface: RenderSurface, commandPackage: ErrorWordingSource): string =>
  error.type === 'no_inlined_bytes'
    ? `--output-path: ${commandName} did not return inlined bytes — this flag works only with commands that produce a body to write. Supported: ${bytesProducingCommands(commandPackage.registry).join(', ')}. ${noBodyRemedy(commandName, surface, commandPackage.binName)}`
    : formatOutputPathError(error, commandName);

const formatOutputDirError = (error: OutputDirError, commandName: string, registry: ErrorWordingSource['registry']): string => {
  if (error.type === 'no_media')
    return `--output-dir: ${commandName} did not return a media array — this flag works only with the image-extraction commands. Supported: ${mediaProducingCommands(registry).join(', ')}.`;
  if (error.type === 'empty_path') return '--output-dir: directory argument is empty (likely a shell-quoting mistake — pass a real directory path)';
  return `--output-dir: write failed: ${error.message}`;
};

export type RunRegistryCommandDeps<G = ReadGraph & WriteGraph> = {
  readonly graph: G;
  readonly fs: FileSystem;
  /** The package the command belongs to: its errors name the package's commands and bin. */
  readonly commandPackage: ErrorWordingSource;
  /**
   * The run journal of the write package. With it, a write that has a plan
   * gets a run id and a record that `undo` reads (record-run.ts). Without it (a
   * library caller, the read package), each command runs as it is.
   */
  readonly runs?: RunRecorder;
  /**
   * The reader of a local file that a param of a write command names
   * (`--ids-file`). Without it, such a read is refused (read-local-file.ts).
   */
  readonly files?: LocalFileReader;
};

export type RunRegistryCommandRequest<G = ReadGraph & WriteGraph> = {
  /** Canonical registry name — drives the manifest-derived error messages. */
  readonly name: string;
  readonly command: Command<G>;
  /** Raw params, keyed by the canonical camelCase option keys. */
  readonly params: Record<string, string>;
  readonly outputPath?: string;
  readonly outputDir?: string;
  /** Which front end asked. Decides whether a remedy may mention a shell. */
  readonly surface: RenderSurface;
  /** Show the changes of a write with a plan, and make none. */
  readonly dryRun?: boolean;
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

// 2026-07-24: one name per flag. The alias normalizer that used to run here
// (mapping alias keys onto canonical ones and rewriting error messages to the
// spelling the caller typed) went with the alias system — params arrive keyed
// by canonical names on every surface, or they are rejected.
const toFailure = (command: Command<never>, error: GraphError): RunRegistryCommandFailure => {
  // 2026-06-15 (F-02): a local-filesystem command (convert-local-file-to-markdown,
  // extract-local-file-images) never touches Graph, so its runtime
  // failures must not be stamped `source: graph` (misleading — an LLM
  // might treat a missing local file as a transient Graph error and
  // retry). Validation errors still flow through the normal classifier.
  const isLocalCommand = command.executeLocal !== undefined;
  return {
    message: error.message,
    ...(error.code === undefined ? {} : { code: error.code }),
    source: isLocalCommand && error.type !== 'validation_error' ? 'cli' : sourceFromGraphError(error),
    ...(error.type === 'api_error' && error.retryAfterSeconds !== undefined ? { retryAfterSeconds: error.retryAfterSeconds } : {}),
  };
};

// --output-path, or with no path the size guard on what goes to stdout.
const persistToPath = async <G>(deps: RunRegistryCommandDeps<G>, request: RunRegistryCommandRequest<G>, value: unknown): Promise<Result<unknown, RunRegistryCommandFailure>> => {
  const persisted = await persistIfRequested(deps.fs, request.outputPath, value, { sourceText: request.command.meta.returnsSourceText === true });
  if (persisted.ok) return ok(persisted.value);
  // Discriminant as errorCode: `no_inlined_bytes` (this flag on a plain-JSON command), `is_directory`,
  // `passthrough_extension_mismatch`, `text_under_binary_extension`, `inline_too_large`, `empty_path`, `write_failed`.
  return err({ message: describeOutputPathError(persisted.error, request.name, request.surface, deps.commandPackage), code: persisted.error.type });
};

// A dry-run shows the plan; a write with a plan and a journal is recorded; any other runs as it is.
const runCommand = async <G>(deps: RunRegistryCommandDeps<G>, request: RunRegistryCommandRequest<G>): Promise<Result<unknown, GraphError>> => {
  const { command, name, params } = request;
  if (request.dryRun === true) return command.plan === undefined ? err(noDryRun(name)) : previewPlan(name, await command.plan(deps.graph, params, deps.files));
  if (command.plan !== undefined && deps.runs !== undefined) return recordRun(deps.runs, name, await command.plan(deps.graph, params, deps.files));
  // `convert-local-file-to-markdown` / `extract-local-file-images` are the commands whose
  // input is the local filesystem, not Graph — route them to executeLocal with the
  // composition-selected FileSystem (the same instance --output-path uses).
  return command.executeLocal !== undefined ? command.executeLocal(deps.fs, params) : command.execute(deps.graph, params, deps.files);
};

export const runRegistryCommand = async <G>(deps: RunRegistryCommandDeps<G>, request: RunRegistryCommandRequest<G>): Promise<Result<unknown, RunRegistryCommandFailure>> => {
  const { command, name } = request;
  const result = await runCommand(deps, request);
  if (!result.ok) return err(toFailure(command, result.error));

  // --output-dir is checked BEFORE --output-path: a media-producing command
  // paired with both should land the media array, not fall into the
  // single-file writer. Order is behaviour, not style.
  if (request.outputDir === undefined) return persistToPath(deps, request, result.value);
  const persistedMedia = await persistMediaIfRequested(deps.fs, request.outputDir, result.value);
  // The error's discriminant doubles as a machine-readable errorCode (`no_media`, `empty_path`,
  // `write_failed`) so an agent routes on the code instead of substring-matching the message.
  if (!persistedMedia.ok) return err({ message: formatOutputDirError(persistedMedia.error, name, deps.commandPackage.registry), code: persistedMedia.error.type });
  // An archive result is the one media result that also has text: its images
  // went to the directory, and --output-path, when given, takes its markdown.
  if (request.outputPath === undefined || !isArchiveResult(persistedMedia.value)) return ok(persistedMedia.value);
  return persistToPath(deps, request, persistedMedia.value);
};

export { buildRenderContext, describeOutputPathError, formatOutputDirError };
