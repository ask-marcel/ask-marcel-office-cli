import type { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { ReadGraph } from '../../infra/read-graph.ts';
import type { WriteGraph } from '../../infra/write-graph.ts';
import type { FileSystem } from '../ports/filesystem.ts';

type CommandSchema = z.ZodType;
/** A command's run on the graph `G` it is given. */
type GraphExecute<G> = (graph: G, params: Record<string, string>) => Promise<Result<unknown, GraphError>>;
/** A command's run on the single package's full client, which holds both graphs. */
type CommandExecute = GraphExecute<GraphClient>;

// `'auth'` and `'contacts'` were declared
// but no command ever used them — removed so the type system enforces "no
// command can ever claim a dead category" and `help-json --category <bad>`
// no longer advertises them as valid.
type CommandCategory = 'drive' | 'excel' | 'sharepoint' | 'tasks' | 'mail' | 'notes' | 'user' | 'calendar' | 'chats' | 'teams' | 'meta' | 'lifecycle';

type CommandHttpMethod = 'GET' | 'POST' | 'PATCH' | 'DELETE';

/*
 * 2026-07-24: one name per flag, one name per command. The alias system
 * (CommandOptionAlias, `aliases?` on options, `commandAliases?` on meta, the
 * composition-level normalizer) was removed: it made the same spelling mean
 * different things across surfaces and blocked uniform unknown-param
 * rejection. `meta.test.ts` now pins the absence.
 */

/**
 * Structured type-hint for a CLI flag value. Surfaces in `help-json` so an
 * LLM can avoid the trial-and-error of "is this an ID or a name?" prose
 * reading. Optional — populate only where the hint is non-obvious from the
 * flag name itself.
 */
type ArgumentHint =
  | { readonly kind: 'idOrName' }
  | { readonly kind: 'magicValue'; readonly values: ReadonlyArray<string> }
  | { readonly kind: 'a1Address' }
  | { readonly kind: 'iso8601' }
  | { readonly kind: 'graphSubpath' };

type CommandOptionMeta = {
  readonly name: string;
  readonly key: string;
  readonly description: string;
  /**
   * `true` for required flags (the historical default; commander rejects the
   * invocation if the flag is missing). `false` for optional flags such as the
   * OData passthrough query parameters (`--top`, `--filter`, …) which
   * commands accept but do not demand.
   */
  readonly required: boolean;
  /**
   * Structured value-type hint for LLM consumers. Optional.
   */
  readonly argumentHint?: ArgumentHint;
};

/**
 * A positional argument (i.e. NOT a `--flag`). Used today only for the
 * `docs` lifecycle command (`ask-marcel-office docs <command>`) but kept as its
 * own field so the manifest never claims a positional is a flag. An LLM
 * consumer reading `help-json` can branch on the presence of
 * `positionalArguments` to know to skip the `--` prefix.
 */
type CommandPositionalArgumentMeta = {
  readonly name: string;
  readonly required: boolean;
  readonly description: string;
};

/**
 * How a paginated command produces subsequent pages. Optional — populate
 * for any command that has `pagination: true`. Lets an LLM tell which
 * cursor field to feed back to `next-page` (or whether `next-page` is even
 * applicable, vs `deltaLink`, vs the header-translation case).
 */
type PaginationStrategy =
  /** Standard: `?$top=N&$skip=K` + `@odata.nextLink` cursor. */
  | 'nextLink'
  /** `?$top=N` + `nextLink` (Graph rejects `$skip` on this endpoint). */
  | 'nextLinkNoSkip'
  /** Delta endpoints — `nextLink` while paging, `deltaLink` on final page. */
  | 'deltaLink'
  /** `--top` translated to `Prefer: odata.maxpagesize` header; `$top` rejected as query. */
  | 'preferMaxPageSize';

/** What a command does to the tenant; see `CommandMeta.effect`. Extended as the writes grow. */
type CommandEffect = 'read' | 'draft' | 'transient-upload';

type CommandMeta = {
  readonly summary: string;
  readonly category: CommandCategory;
  readonly graphMethod: CommandHttpMethod;
  readonly graphPathTemplate: string;
  readonly graphDocsUrl: string;
  readonly options: ReadonlyArray<CommandOptionMeta>;
  readonly positionalArguments?: ReadonlyArray<CommandPositionalArgumentMeta>;
  readonly example: string;
  readonly responseShape?: string;
  readonly bodyTemplate?: string;
  readonly pagination?: true;
  readonly paginationStrategy?: PaginationStrategy;
  /**
   * Graph permission scopes the endpoint requires. The basic Teams web-client
   * token grants ~30 scopes (run `ask-marcel-office status` to see). Commands
   * with unmet scopes return `403 Forbidden: Missing scope` at the wire. Use
   * this for pre-flight checks rather than failing on-the-wire.
   *
   * Source: the Microsoft Graph permissions reference
   * (https://learn.microsoft.com/en-us/graph/permissions-reference) and the
   * command's `graphDocsUrl` page. List the LEAST-PRIVILEGED delegated scope
   * Microsoft documents, unless the token does not carry it: group posts list
   * `Group.Read.All`, the documented higher alternative the basic token holds,
   * because no token carries `Group-Conversation.Read.All`. A command that
   * writes declares the write scope (the PDF converters' `Files.ReadWrite`).
   * Every command declares it here (the central graph-scopes.ts map was
   * dissolved into the commands at package-split step 11), except the ones that
   * call no fixed Graph endpoint: the chat substrate, local files, link parsers,
   * `next-page` and `microsoft-search-query`. meta.test.ts pins that set.
   */
  readonly scopesRequired?: ReadonlyArray<string>;
  /**
   * `true` if the command needs the M365ChatClient elevated token (captured
   * at login from `m365.cloud.microsoft`, ODSP allow-list). The exact set is
   * pinned in meta.test.ts, and the composition root derives the auth
   * fail-fast message's command list from this flag — a new elevated command
   * MUST carry it or the remedy message will omit the command. An LLM should
   * check this field before invoking; if the elevated capture failed at
   * login, these commands fail fast with `secondary_token_unavailable`.
   */
  readonly needsElevatedToken?: true;
  /**
   * Which Teams substrate token the command needs: `'chatsvcagg'`
   * (teams.microsoft.com/api/csa) or `'ic3'` (chat history), both captured
   * at login from `teams.microsoft.com`. Truthy = "needs a substrate token"
   * (the historical boolean semantics); the service value additionally
   * routes the command into the right auth fail-fast list, derived by the
   * composition root. Like `needsElevatedToken`, an LLM should check this
   * before invoking and warm up an interactive `login`; a headless or stale
   * session times out on these (the non-interactive silent-SSO limitation).
   */
  readonly needsSubstrateToken?: 'chatsvcagg' | 'ic3';
  /**
   * `true` if the command returns inlined bytes (`{contentType, size, base64}`
   * or `{contentType, size, text}`) and is therefore a valid target for the
   * global `--output-path` flag. Used by the CLI composition to derive the
   * rejection-message whitelist from the manifest rather than hand-keeping it
   * as a string literal. Audit .
   */
  readonly producesBytes?: true;
  /**
   * `true` when the command's `text` answer is the file's own bytes (a UTF-8 file
   * returned as text), never a conversion, so `--output-path` saves it under any
   * name, the file's own `.xls` or `.doc` included. Converted answers stay barred
   * from binary names (see output-path.ts).
   */
  readonly returnsSourceText?: true;
  /**
   * `true` if the command returns a `media` array (`{ count, media: [{ path,
   * contentType, sizeBytes, base64 }] }`) and is therefore a valid target for
   * the global `--output-dir` flag, which writes each image to a directory.
   * Parallel to `producesBytes` (single-file `--output-path`).
   */
  readonly producesMedia?: true;
  /**
   * What the command does to the tenant. `read` changes nothing (a search POST
   * included). Every other value is a write class, worded in `command-effect.ts`:
   * `draft` creates or updates an UNSENT mail draft; `transient-upload` puts a
   * temporary file in the user's OneDrive and tries to delete it. Required, so a new
   * command chooses its class instead of inheriting `read`. The MCP routing
   * (run-command takes only `read`), the MCP annotations, the top-level `--help`
   * narrative and the per-command docs derive from it, and the manifest /
   * `help-json` carry it on every command. Replaces the `mutates` flag (package
   * split, phase 1 step 11).
   */
  readonly effect: CommandEffect;
  /**
   * Stability tier of the command. Omitted from manifest entries when the
   * command is `'stable'` (the implicit default), surfaced only on
   * `'experimental'` commands so an LLM can prefer stable siblings when they
   * exist. `'experimental'` today means the command rides a Microsoft-internal
   * substrate (chatsvcagg / IC3) that is not in the public Graph API and can
   * break on a Teams web-client update — the docstring "Best-effort, may break
   * on Microsoft client updates" warnings now have a structured pair.
   *
   */
  readonly stability?: 'experimental';
};

/**
 * Present on the rare command whose input is the LOCAL filesystem instead of
 * Graph (`convert-local-file-to-markdown`). The CLI routes execution here, passing its
 * composition-selected FileSystem; `execute` stays as the registry-typed
 * fallback that redirects library consumers to this variant.
 */
type LocalExecute = (fs: FileSystem, params: Record<string, string>) => Promise<Result<unknown, GraphError>>;

/** The meta of a command that changes nothing in the tenant. */
type ReadCommandMeta = CommandMeta & { readonly effect: 'read' };
/** The meta of a command that writes; `command-effect.ts` words each class. */
type WriteCommandMeta = CommandMeta & { readonly effect: Exclude<CommandEffect, 'read'> };

/**
 * A command that changes nothing: it runs on the read graph, which has no
 * member that writes (package split, D9). Its effect is `read`, so the registry
 * can tell it from a write by its meta.
 */
type ReadCommand = {
  readonly schema: CommandSchema;
  readonly execute: GraphExecute<ReadGraph>;
  readonly meta: ReadCommandMeta;
  readonly executeLocal?: LocalExecute;
};

/** A command that writes: it runs on the write graph, basic tier only. */
type WriteCommand = {
  readonly schema: CommandSchema;
  readonly execute: GraphExecute<WriteGraph>;
  readonly meta: WriteCommandMeta;
};

/** What the registry holds: each command typed against the narrower graph it runs on. */
type RegisteredCommand = ReadCommand | WriteCommand;

/**
 * Any command, read or write, run on the full client. The docs, the manifest
 * and the CLI and MCP shells take this type; every registered command is one.
 */
type Command = {
  readonly schema: CommandSchema;
  readonly execute: CommandExecute;
  readonly meta: CommandMeta;
  readonly executeLocal?: LocalExecute;
};

export type {
  ArgumentHint,
  Command,
  CommandCategory,
  CommandEffect,
  CommandExecute,
  CommandHttpMethod,
  CommandMeta,
  CommandOptionMeta,
  CommandPositionalArgumentMeta,
  CommandSchema,
  GraphExecute,
  PaginationStrategy,
  ReadCommand,
  ReadCommandMeta,
  RegisteredCommand,
  WriteCommand,
  WriteCommandMeta,
};
