import type { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';

type CommandSchema = z.ZodType;
type CommandExecute = (graph: GraphClient, params: Record<string, string>) => Promise<Result<unknown, import('../../infra/graph-client.ts').GraphError>>;

// Audit Alex-session §4 follow-up: `'auth'` and `'contacts'` were declared
// but no command ever used them — removed so the type system enforces "no
// command can ever claim a dead category" and `help-json --category <bad>`
// no longer advertises them as valid.
type CommandCategory = 'drive' | 'excel' | 'sharepoint' | 'tasks' | 'mail' | 'notes' | 'user' | 'calendar' | 'chats' | 'teams' | 'meta' | 'lifecycle';

type CommandHttpMethod = 'GET' | 'POST' | 'PATCH' | 'DELETE';

type CommandOptionAlias = {
  readonly name: string; // kebab-case (the user-facing flag, without `--`)
  readonly key: string; // camelCase (commander auto-keyed name)
};

/*
 * Alias policy (audit round-7 I3): the `--folder-id` alias is only exposed
 * on commands whose path is folder-scoped (e.g. `list-folder-files`,
 * `get-drive-delta` — the item ID MUST point at a folder). Commands like
 * `get-drive-item`, `list-drive-item-versions`, and
 * `download-drive-item-content` accept BOTH files and folders, so they
 * keep `--item-id` as the only name to avoid misleading the LLM into
 * thinking the value must be a folder.
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
   * Optional secondary spellings of the same flag. Both the canonical
   * `name` and every alias name are accepted on the command line; values
   * passed under an alias are normalized to the canonical `key` before
   * the schema runs. The canonical name is what `--help` shows first.
   */
  readonly aliases?: ReadonlyArray<CommandOptionAlias>;
  /**
   * Structured value-type hint for LLM consumers. Optional.
   */
  readonly argumentHint?: ArgumentHint;
};

/**
 * A positional argument (i.e. NOT a `--flag`). Used today only for the
 * `docs` lifecycle command (`ask-marcel docs <command>`) but kept as its
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

type CommandMeta = {
  readonly summary: string;
  readonly category: CommandCategory;
  /**
   * Deprecated former command names kept working as commander-level aliases for
   * back-compat after a rename (e.g. `download-onedrive-file-content` →
   * `download-drive-item-content`). The canonical registry key is what `--help`
   * and the manifest list first; each alias here is also accepted on the CLI and
   * surfaced in the manifest so an LLM that learned the old name still resolves.
   */
  readonly commandAliases?: ReadonlyArray<string>;
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
   * token grants ~30 scopes (run `ask-marcel scopes-check` to see). Commands
   * with unmet scopes return `403 Forbidden: Missing scope` at the wire. Use
   * this for pre-flight checks rather than failing on-the-wire. Optional —
   * populated only on commands where the audit confirmed a scope-failure
   * path; absent means "should work with the basic Teams token".
   */
  readonly scopesRequired?: ReadonlyArray<string>;
  /**
   * `true` if the command needs the M365ChatClient elevated token (captured
   * at login from `m365.cloud.microsoft`, ODSP allow-list). Only 3 commands
   * today — the historical-version downloads. An LLM should check this
   * field before invoking; if the elevated capture failed at login, these
   * commands will time out.
   */
  readonly needsElevatedToken?: true;
  /**
   * `true` if the command returns inlined bytes (`{contentType, size, base64}`
   * or `{contentType, size, text}`) and is therefore a valid target for the
   * global `--output-path` flag. Used by the CLI composition to derive the
   * rejection-message whitelist from the manifest rather than hand-keeping it
   * as a string literal. Audit round-8 Wave E2.
   */
  readonly producesBytes?: true;
  /**
   * `true` if the command returns a `media` array (`{ count, media: [{ path,
   * contentType, sizeBytes, base64 }] }`) and is therefore a valid target for
   * the global `--output-dir` flag, which writes each image to a directory.
   * Parallel to `producesBytes` (single-file `--output-path`).
   */
  readonly producesMedia?: true;
  /**
   * `true` if the command performs a write (mutation) against Microsoft Graph
   * rather than a read/search. Today only the two mail-draft commands
   * (`create-mail-draft` = POST, `update-mail-draft` = PATCH) — both produce an
   * UNSENT draft and cannot send. The CLI composition derives its read-only
   * top-level `--help` narrative from this flag (mutating commands vs search
   * POSTs) instead of hardcoding command names, and the manifest/`help-json`
   * surface it so an LLM can tell at a glance which commands change state.
   * Audit 2026-06-15 (F-03).
   */
  readonly mutates?: true;
  /**
   * Stability tier of the command. Omitted from manifest entries when the
   * command is `'stable'` (the implicit default), surfaced only on
   * `'experimental'` commands so an LLM can prefer stable siblings when they
   * exist. `'experimental'` today means the command rides a Microsoft-internal
   * substrate (chatsvcagg / IC3) that is not in the public Graph API and can
   * break on a Teams web-client update — the docstring "Best-effort, may break
   * on Microsoft client updates" warnings now have a structured pair.
   * Audit Alex-session §6.
   */
  readonly stability?: 'experimental';
};

type Command = {
  readonly schema: CommandSchema;
  readonly execute: CommandExecute;
  readonly meta: CommandMeta;
  /**
   * Present on the rare command whose input is the LOCAL filesystem instead of
   * Graph (`convert-local-file`). The CLI routes execution here, passing its
   * composition-selected FileSystem; `execute` stays as the registry-typed
   * fallback that redirects library consumers to this variant.
   */
  readonly executeLocal?: (
    fs: import('../ports/filesystem.ts').FileSystem,
    params: Record<string, string>
  ) => Promise<Result<unknown, import('../../infra/graph-client.ts').GraphError>>;
};

export type {
  ArgumentHint,
  Command,
  CommandCategory,
  CommandExecute,
  CommandHttpMethod,
  CommandMeta,
  CommandOptionAlias,
  CommandOptionMeta,
  CommandPositionalArgumentMeta,
  CommandSchema,
  PaginationStrategy,
};
