import type { CommandManifestEntry } from './docs-render.ts';

/**
 * Lifecycle commands aren't backed by a Graph endpoint and aren't in the
 * `commands` registry — `docs` and `help-json` live in `cli-shell.ts`, `login`,
 * `status` and `logout` in `auth-cli.ts`, and `mcp` in `main.ts`. But an LLM that consumes
 * `help-json` (the recommended token-efficient surface) needs to see them too,
 * otherwise it can't authenticate, log out, or fetch per-command docs. These
 * stub entries surface them under category 'lifecycle' with no graphMethod /
 * options (the path template is a `(meta) ...` description so the JSON shape
 * stays uniform with regular commands). The single package hands them to the
 * shells as its `CommandPackage.lifecycle` (`composition/single-package.ts`).
 */
export const LIFECYCLE_ENTRIES: ReadonlyArray<CommandManifestEntry> = [
  {
    name: 'login',
    summary:
      "Authenticate against Microsoft Graph using the Teams web client (cached token → refresh → browser fallback). Stores tokens at ~/.ask-marcel/token-cache.json (0600). On success reports which of the four tokens are currently available and points to `status` for each token's scopes + expiry; pass --force to re-capture every token via the browser in one pass. Run before any Graph command.",
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
          'Ignore the cache and re-capture every token via the browser. Rarely needed, and NOT free: it clears the browser session first so the token grant re-fires, which deletes the 90-day "Stay signed in" cookie and usually means entering a password again. A plain `login` already re-captures the elevated (M365) token when it is missing and leaves that session intact, so use `--force` only when a tier is still stuck after one.',
      },
    ],
    example: 'ask-marcel-office login --force',
    responseShape:
      '{ status: "authenticated", available: string[], unlocked: Record<string,string>, missing: Record<string,string>, hint } on success. `available` lists the token tiers currently cached and fresh — always "basic", plus "elevated" / "chatsvcagg" / "ic3" when present. `unlocked` maps each AVAILABLE tier to a plain-language description of what it lets you read (e.g. `chatsvcagg` -> "Teams chat message content"), so a caller need not know what a tier codename means. `missing` maps each ABSENT tier the same way, with the re-capture remedy appended — an empty object means every tier was captured. `login` redeems the shared refresh token for the two SUBSTRATE tiers (chatsvcagg / ic3) before reporting, so finding them in `missing` means that headless attempt failed, not that nothing was tried; the elevated tier is browser-only and is re-captured by the sign-in itself. Read `missing` to learn what a partial login costs BEFORE a command fails for it. `hint` points to `status` (each token\'s scopes + expiry) and `login --force` (refresh). The detailed per-token status lives in `status`, not here. Envelope error on cancel/failure.',
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
    name: 'status',
    summary:
      'Show the four cached tokens (basic, elevated, chatsvcagg and ic3). For each token, the report shows if the token is available and the seconds before its expiry. It also shows the scopes, the refresh method, and the data that the token lets you read. The command reads only the token cache, and it does not make a Graph call or open a browser. Before a long run that no person watches, compare the `scopesRequired` of a command (in `help-json`) with the `scopes` of the token that the command uses. If you are not signed in, the error code is `not_authenticated`.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) decodes the cached tokens; not a Graph endpoint',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/permissions-reference',
    options: [],
    example: 'ask-marcel-office status',
    responseShape:
      '`{ basic: TokenTier, elevated: TokenTier, chatsvcagg: TokenTier, ic3: TokenTier, hint: string }`, where `TokenTier = { available: boolean, expiresInSeconds?: number, scopes: string[], refresh: "automatic" | "interactive", reads: string, reason?: string }` and `hint` tells how to refresh the tokens. `available` is true only when the token is in the cache and the time before its expiry is more than 5 minutes (300 seconds). `expiresInSeconds` is negative after the expiry, and the key is not there when the cache has no token. `scopes` comes from the `scp` claim of that token, and `reads` gives the data that the token lets you read. `refresh` is `automatic` when the shared refresh token can refresh the token, and `interactive` for the elevated token, which only `login` can capture again. `reason` is there only when `available` is false, and it tells why the token is not available and how to get it again.',
  },
  {
    name: 'docs',
    summary:
      'Print Markdown docs for a single command (the same per-command page that ships in `docs/commands.json`). Pass the command name as a POSITIONAL argument — there is no `--command` flag. For lifecycle commands (login/status/logout/docs) prints the same --help that command would.',
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
      'Print the machine-readable command manifest as JSON. For fresh-session discovery use `--terse --category <name>` (~8 KB for one category). The unflagged form is the *full* reference (every option / example / response shape per command) and is roughly 14× the size of `ask-marcel-office --help` — reach for it only after `--terse` has narrowed the search. `--terse` alone projects each entry to `{name, summary, category}` (summary compacted to its first sentence). Categories: lifecycle, drive, excel, sharepoint, tasks, mail, notes, user, calendar, chats, teams, meta.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) renders the in-process command manifest',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/',
    options: [],
    example: 'ask-marcel-office help-json',
    responseShape: '{ package, version, generatedAt, commands: [{ name, summary, category, ... }, ...] }',
  },
  {
    name: 'mcp',
    summary:
      'Start the Model Context Protocol server over stdio, exposing every command to an MCP client (Claude Code, Claude Desktop, Cursor, …) as six gateway tools (list-commands, get-command-docs, run-command, run-write-command, login, status). Runs in the foreground and serves until the client disconnects — a server entry point, not a one-shot command, so a shell- or MCP-driving agent never calls it directly. Register with `claude mcp add --transport stdio --scope user ask-marcel-office -- ask-marcel-office mcp`. See docs/USAGE.md (MCP server).',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) serves the command registry over MCP stdio; not a Graph endpoint',
    graphDocsUrl: 'https://modelcontextprotocol.io/',
    options: [],
    example: 'ask-marcel-office mcp',
    responseShape: 'None — a long-lived stdio server emitting JSON-RPC frames until the client disconnects (never a JSON envelope).',
  },
];
