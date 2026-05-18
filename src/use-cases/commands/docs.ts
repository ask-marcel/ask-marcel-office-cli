import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import type { Command } from './command-types.ts';
import type { CommandManifest, CommandManifestEntry } from './docs-render.ts';
import { renderCommandMarkdown } from './docs-render.ts';
import { lookupScopes } from './graph-scopes.ts';

export type DocsError = { type: 'unknown_command'; readonly name: string; readonly available: ReadonlyArray<string> };

const toEntry = (name: string, cmd: Command): CommandManifestEntry => {
  // Default every `pagination: true` command to `nextLink` strategy when the
  // command file doesn't specify one — that's the standard $top + nextLink
  // shape, true of ~80% of paginated commands. Audit-round-7 Wave F: makes
  // the manifest field always populated on paginated commands so LLM
  // consumers don't need to read prose to learn the cursor mechanism.
  const paginationStrategy = cmd.meta.paginationStrategy ?? (cmd.meta.pagination ? 'nextLink' : undefined);
  // Audit round-8 Wave C: scopesRequired now comes from a central map by
  // default (`graph-scopes.ts`). Per-command inline overrides win so that
  // future commands with non-standard scope needs can declare them locally.
  const scopesRequired = cmd.meta.scopesRequired ?? lookupScopes(name);
  return {
    name,
    summary: cmd.meta.summary,
    category: cmd.meta.category,
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
    ...(cmd.meta.producesBytes ? { producesBytes: cmd.meta.producesBytes } : {}),
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
      'Authenticate against Microsoft Graph using the Teams web client (cached token → refresh → browser fallback). Stores tokens at ~/.ask-marcel/token-cache.json (0600). Run before any Graph command.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) browser-OAuth via Teams web client; not a Graph endpoint',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/auth-v2-user',
    options: [],
    example: 'ask-marcel login',
    responseShape: '{ status: "authenticated" } on success; envelope error on cancel/failure.',
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
    example: 'ask-marcel logout',
    responseShape: '{ status: "logged_out" } on success.',
  },
  {
    name: 'update',
    summary:
      'Re-install the latest published ask-marcel from npm, in place. Auto-detects whether you originally installed via npm or bun based on the bin path. Token cache is preserved. Do NOT use from a local clone — pull and re-run `bun install` instead.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) shells out to `npm i -g` or `bun add -g`; not a Graph endpoint',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/',
    options: [],
    example: 'ask-marcel update',
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
          'Name of the command to show docs for (e.g. `list-mail-messages`). Run `ask-marcel --help` or `ask-marcel help-json` for the full list. Passed as a bare positional — do NOT prefix with `--command`.',
      },
    ],
    example: 'ask-marcel docs list-mail-messages',
    responseShape: 'Markdown text on stdout (NOT JSON-wrapped — this is the one command whose stdout is plain text).',
  },
  {
    name: 'help-json',
    summary:
      'Print the full machine-readable command manifest as JSON to stdout (same content as `docs/commands.json`). Token-friendly alternative to `ask-marcel --help`, which is ~40 KB. Includes lifecycle commands (login/logout/update/docs/help-json) under category `lifecycle`.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) renders the in-process command manifest',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/',
    options: [],
    example: 'ask-marcel help-json',
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

const findLifecycleEntry = (name: string): CommandManifestEntry | undefined => LIFECYCLE_ENTRIES.find((entry) => entry.name === name);

export const renderSingleCommand = (registry: Readonly<Record<string, Command>>, name: string): Result<string, DocsError> => {
  const cmd = registry[name];
  if (cmd) return ok(renderCommandMarkdown(toEntry(name, cmd)));
  const lifecycle = findLifecycleEntry(name);
  if (lifecycle) return ok(renderCommandMarkdown(lifecycle));
  return err({ type: 'unknown_command', name, available: [...Object.keys(registry), ...LIFECYCLE_ENTRIES.map((e) => e.name)].toSorted((a, b) => a.localeCompare(b)) });
};
