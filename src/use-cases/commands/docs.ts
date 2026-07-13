import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { CommandCategory, Command, CommandMeta } from './command-types.ts';
import type { CommandManifest, CommandManifestEntry } from './docs-render.ts';
import { CATEGORY_LABELS, renderCommandMarkdown } from './docs-render.ts';
import { firstSentence } from './first-sentence.ts';
import { lookupScopes } from './graph-scopes.ts';

export type DocsError = { type: 'unknown_command'; readonly name: string; readonly available: ReadonlyArray<string> };

/**
 * Terse manifest entry — only the fields an LLM needs to *discover* a command
 * (i.e. "does this CLI do X?"). Drops `options`, `example`, `graphPathTemplate`,
 * `graphDocsUrl`, `responseShape`, `bodyTemplate`, `paginationStrategy`,
 * `scopesRequired` — everything the LLM only needs once it's already decided
 * to invoke. `stability` is kept (it's a discovery-time concern: LLMs prefer
 * stable siblings when they exist, so they need to see the tag at discovery
 * time, not after a second full-manifest fetch). /§6.
 */
export type TerseManifestEntry = {
  readonly name: string;
  readonly summary: string;
  readonly category: CommandCategory;
  readonly stability?: CommandMeta['stability'];
};

export type TerseManifest = {
  readonly package: string;
  readonly version: string;
  readonly generatedAt: string;
  readonly commands: ReadonlyArray<TerseManifestEntry>;
};

export type ManifestFilterError = { readonly type: 'unknown_category'; readonly category: string; readonly available: ReadonlyArray<string> };

const toEntry = (name: string, cmd: Command): CommandManifestEntry => {
  // Default every `pagination: true` command to `nextLink` strategy when the
  // command file doesn't specify one — that's the standard $top + nextLink
  // shape, true of ~80% of paginated commands. Audit-makes
  // the manifest field always populated on paginated commands so LLM
  // consumers don't need to read prose to learn the cursor mechanism.
  const paginationStrategy = cmd.meta.paginationStrategy ?? (cmd.meta.pagination ? 'nextLink' : undefined);
  // scopesRequired now comes from a central map by
  // default (`graph-scopes.ts`). Per-command inline overrides win so that
  // future commands with non-standard scope needs can declare them locally.
  const scopesRequired = cmd.meta.scopesRequired ?? lookupScopes(name);
  return {
    name,
    summary: cmd.meta.summary,
    category: cmd.meta.category,
    ...(cmd.meta.commandAliases ? { commandAliases: cmd.meta.commandAliases } : {}),
    graphMethod: cmd.meta.graphMethod,
    graphPathTemplate: cmd.meta.graphPathTemplate,
    graphDocsUrl: cmd.meta.graphDocsUrl,
    options: cmd.meta.options,
    example: cmd.meta.example,
    ...(cmd.meta.positionalArguments ? { positionalArguments: cmd.meta.positionalArguments } : {}),
    ...(cmd.meta.responseShape ? { responseShape: cmd.meta.responseShape } : {}),
    ...(cmd.meta.bodyTemplate ? { bodyTemplate: cmd.meta.bodyTemplate } : {}),
    ...(cmd.meta.pagination ? { pagination: cmd.meta.pagination } : {}),
    ...(paginationStrategy ? { paginationStrategy } : {}),
    ...(scopesRequired && scopesRequired.length > 0 ? { scopesRequired } : {}),
    ...(cmd.meta.needsElevatedToken ? { needsElevatedToken: cmd.meta.needsElevatedToken } : {}),
    ...(cmd.meta.needsSubstrateToken ? { needsSubstrateToken: cmd.meta.needsSubstrateToken } : {}),
    ...(cmd.meta.producesBytes ? { producesBytes: cmd.meta.producesBytes } : {}),
    ...(cmd.meta.producesMedia ? { producesMedia: cmd.meta.producesMedia } : {}),
    ...(cmd.meta.mutates ? { mutates: cmd.meta.mutates } : {}),
    ...(cmd.meta.stability ? { stability: cmd.meta.stability } : {}),
  };
};

/**
 * Lifecycle commands aren't backed by a Graph endpoint and aren't in the
 * `commands` registry — they live inline in `cli.ts`. But an LLM that consumes
 * `help-json` (the recommended token-efficient surface) needs to see them too,
 * otherwise it can't authenticate, log out, or fetch per-command docs. These
 * stub entries surface them under category 'lifecycle' with no graphMethod /
 * options (the path template is a `(meta) ...` description so the JSON shape
 * stays uniform with regular commands).
 */
const LIFECYCLE_ENTRIES: ReadonlyArray<CommandManifestEntry> = [
  {
    name: 'login',
    summary:
      'Authenticate against Microsoft Graph using the Teams web client (cached token → refresh → browser fallback). Stores tokens at ~/.ask-marcel/token-cache.json (0600). Reports all four cached tokens (basic / elevated / chatsvcagg / ic3) with their time-left and how to refresh each; pass --force to re-capture every token via the browser in one pass. Run before any Graph command.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) browser-OAuth via Teams web client; not a Graph endpoint',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/auth-v2-user',
    options: [
      {
        name: 'force',
        key: 'force',
        required: false,
        description:
          'Ignore the cache and re-capture every token via the browser. The only way to refresh the elevated (M365) token while the basic token is still valid; the persistent browser profile is reused, so you are usually not re-prompted for credentials.',
      },
    ],
    example: 'ask-marcel-office login --force',
    responseShape:
      '{ status: "authenticated", tokens: { basic, elevated, chatsvcagg, ic3 }, hint } on success. Each token is { available: boolean, expiresInSeconds?: number, refresh: "automatic" | "interactive", reason? }: basic/chatsvcagg/ic3 refresh automatically from the cached refresh token, the elevated (M365) token is re-captured only on an interactive login. `expiresInSeconds` is omitted when the token is not cached; a failed elevated capture this run adds `reason`. Envelope error on cancel/failure.',
  },
  {
    name: 'logout',
    summary:
      'Clear the cached Microsoft Graph token so the next command forces a fresh sign-in. Removes ~/.ask-marcel/token-cache.json; leaves the Playwright browser profile alone.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) deletes ~/.ask-marcel/token-cache.json; not a Graph endpoint',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/auth-v2-user',
    options: [],
    example: 'ask-marcel-office logout',
    responseShape: '{ status: "logged_out" } on success.',
  },
  {
    name: 'update',
    summary:
      'Re-install the latest published ask-marcel-office from npm, in place. Auto-detects whether you originally installed via npm or bun based on the bin path. Token cache is preserved. Do NOT use from a local clone — pull and re-run `bun install` instead.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) shells out to `npm i -g` or `bun add -g`; not a Graph endpoint',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/',
    options: [],
    example: 'ask-marcel-office update',
    responseShape: '{ status: "updated", via: "npm" | "bun" } on success.',
  },
  {
    name: 'docs',
    summary:
      'Print Markdown docs for a single command (the same per-command page that ships in `docs/commands.json`). Pass the command name as a POSITIONAL argument — there is no `--command` flag. For lifecycle commands (login/logout/update/docs) prints the same --help that command would.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) renders Markdown from the in-process command manifest',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/',
    options: [],
    positionalArguments: [
      {
        name: 'command',
        required: true,
        description:
          'Name of the command to show docs for (e.g. `list-mail-messages`). Run `ask-marcel-office --help` or `ask-marcel-office help-json` for the full list. Passed as a bare positional — do NOT prefix with `--command`.',
      },
    ],
    example: 'ask-marcel-office docs list-mail-messages',
    responseShape: 'Markdown text on stdout (NOT JSON-wrapped — this is the one command whose stdout is plain text).',
  },
  {
    name: 'help-json',
    summary:
      'Print the machine-readable command manifest as JSON. For fresh-session discovery use `--terse --category <name>` (~6 KB for one category). The unflagged form is the *full* reference (every option / example / response shape per command) and is roughly 13× the size of `ask-marcel-office --help` — reach for it only after `--terse` has narrowed the search. `--terse` alone projects each entry to `{name, summary, category}` (summary compacted to its first sentence). Categories: lifecycle, drive, excel, sharepoint, tasks, mail, notes, user, calendar, chats, teams, meta.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) renders the in-process command manifest',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/',
    options: [],
    example: 'ask-marcel-office help-json',
    responseShape: '{ package, version, generatedAt, commands: [{ name, summary, category, ... }, ...] }',
  },
];

const buildEntries = (registry: Readonly<Record<string, Command>>): ReadonlyArray<CommandManifestEntry> =>
  [...Object.entries(registry).map(([name, cmd]) => toEntry(name, cmd)), ...LIFECYCLE_ENTRIES].toSorted((a, b) => a.name.localeCompare(b.name));

export const buildManifest = (registry: Readonly<Record<string, Command>>, packageName: string, version: string, now: () => Date = () => new Date()): CommandManifest => ({
  package: packageName,
  version,
  generatedAt: now().toISOString(),
  commands: buildEntries(registry),
});

/**
 * Terse manifest — `{ name, summary, category }` per command. Roughly 95%
 * smaller than the full manifest (no options/example/Graph endpoint per entry).
 * Use `help-json --terse` to surface this to an LLM as the discovery view.
 */
export const buildTerseManifest = (registry: Readonly<Record<string, Command>>, packageName: string, version: string, now: () => Date = () => new Date()): TerseManifest => ({
  package: packageName,
  version,
  generatedAt: now().toISOString(),
  commands: buildEntries(registry).map((e) => ({
    name: e.name,
    // Terse is the per-category discovery view — keep it token-cheap by
    // projecting each multi-sentence summary to its first sentence (the full
    // summary stays in the unflagged manifest). Keeps drive/mail under budget.
    summary: firstSentence(e.summary),
    category: e.category,
    ...(e.stability ? { stability: e.stability } : {}),
  })),
});

/**
 * Filter a `CommandManifest` (or terse variant) down to a single category.
 * Returns `err({ type: 'unknown_category', ... })` if the requested category
 * isn't a known one — the CLI surfaces this through the standard error
 * envelope rather than silently returning an empty list.
 */
export const filterManifestByCategory = <M extends { readonly commands: ReadonlyArray<{ readonly category: CommandCategory }> }>(
  manifest: M,
  category: string
): Result<M, ManifestFilterError> => {
  const knownCategories = Object.keys(CATEGORY_LABELS).toSorted((a, b) => a.localeCompare(b));
  if (!knownCategories.includes(category)) {
    return err({ type: 'unknown_category', category, available: knownCategories });
  }
  const filtered = manifest.commands.filter((c) => c.category === category);
  return ok({ ...manifest, commands: filtered });
};

const findLifecycleEntry = (name: string): CommandManifestEntry | undefined => LIFECYCLE_ENTRIES.find((entry) => entry.name === name);

export const renderSingleCommand = (registry: Readonly<Record<string, Command>>, name: string): Result<string, DocsError> => {
  const cmd = registry[name];
  if (cmd) return ok(renderCommandMarkdown(toEntry(name, cmd)));
  const lifecycle = findLifecycleEntry(name);
  if (lifecycle) return ok(renderCommandMarkdown(lifecycle));
  return err({ type: 'unknown_command', name, available: [...Object.keys(registry), ...LIFECYCLE_ENTRIES.map((e) => e.name)].toSorted((a, b) => a.localeCompare(b)) });
};
