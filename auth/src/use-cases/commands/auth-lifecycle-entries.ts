import type { CommandManifestEntry } from '@ask-marcel/office-core/use-cases/commands/docs-render.ts';

/*
 * The lifecycle commands of the auth package `@ask-marcel/office-auth`. It has
 * no registry, so these entries are its whole manifest: the three sign-in
 * commands (`auth-cli.ts`) and its own `docs`, `help-json` and `mcp`. The
 * MCP server of the auth package has two tools, `login` and `status`; no text
 * here offers the token helper (LESSONS 2026-10-15).
 */

/** The sign-in commands: `login`, `logout` and `status`. */
const AUTH_SIGN_IN_ENTRIES: ReadonlyArray<CommandManifestEntry> = [
  {
    name: 'login',
    summary:
      "Authenticate against Microsoft Graph using the Teams web client (cached token → refresh → browser fallback). Stores tokens at ~/.ask-marcel/token-cache.json (0600). On success reports which of the four tokens are currently available and the end of the sign-in session (Microsoft limits it to about 24 hours after the browser sign-in), and points to `status` for each token's scopes + expiry. Pass --renew to start a new sign-in session with the browser cookies kept, or --force to re-capture every token via the browser in one pass. Run before any Graph command.",
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
      {
        name: 'renew',
        key: 'renew',
        required: false,
        description:
          'Start a new sign-in session before the current one ends. Microsoft limits the session to about 24 hours after the browser sign-in, and `status` shows its end as `session.endsAt`. It re-captures every token and a new refresh token in the browser, and keeps the browser cookies, so it usually asks for no password. A plain `login` opens no browser while the session is good. With `--force`, `--force` wins.',
      },
    ],
    example: '{auth-bin} login --force',
    responseShape:
      '{ status: "authenticated", available: string[], unlocked: Record<string,string>, missing: Record<string,string>, session: { endsAt: string, endsInSeconds: number } | { reason: string }, hint } on success. `available` lists the token tiers currently cached and fresh — always "basic", plus "elevated" / "chatsvcagg" / "ic3" when present. `unlocked` maps each AVAILABLE tier to a plain-language description of what it lets you read (e.g. `chatsvcagg` -> "Teams chat message content"), so a caller need not know what a tier codename means. `missing` maps each ABSENT tier the same way, with the re-capture remedy appended — an empty object means every tier was captured. `login` redeems the shared refresh token for the two SUBSTRATE tiers (chatsvcagg / ic3) before reporting, so finding them in `missing` means that headless attempt failed, not that nothing was tried; the elevated tier is browser-only and is re-captured by the sign-in itself. Read `missing` to learn what a partial login costs BEFORE a command fails for it. `session` gives the end of the sign-in session (about 24 hours after the browser sign-in that gave the refresh token), as an ISO time and the seconds left; when the token cache does not hold the time of that sign-in, it gives only a `reason`. `hint` points to `status` (each token\'s scopes + expiry), `login --renew` (a new session) and `login --force` (refresh). The detailed per-token status lives in `status`, not here. Envelope error on cancel/failure. A plain `login` whose cache holds no refresh token, or whose headless refresh of chatsvcagg or ic3 failed, captures every tier and a new refresh token in the browser with the cookies kept. If the cache still holds no refresh token after the sign-in, `login` fails with the error code `session_cannot_renew` (exit 1): the session cannot renew itself, and the message tells you to run `login --force`.',
  },
  {
    name: 'logout',
    summary:
      'Clear the cached tokens and the browser profile so the next command forces a fresh sign-in. Removes ~/.ask-marcel/token-cache.json and the Playwright browser profile (~/.ask-marcel/browser-profile/, or $ASKMARCEL_BROWSER_PROFILE when set). The profile holds the "Stay signed in" cookie, so the next sign-in usually asks for a password again.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) deletes ~/.ask-marcel/token-cache.json and the browser profile; not a Graph endpoint',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/auth-v2-user',
    options: [],
    example: '{auth-bin} logout',
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
    example: '{auth-bin} status',
    responseShape:
      '`{ basic: TokenTier, elevated: TokenTier, chatsvcagg: TokenTier, ic3: TokenTier, refreshToken: "cached" | "missing", session: { endsAt: string, endsInSeconds: number } | { reason: string }, hint: string }`, where `TokenTier = { available: boolean, expiresInSeconds?: number, scopes: string[], refresh: "automatic" | "interactive", reads: string, reason?: string }` and `hint` tells how to refresh the tokens. `available` is true only when the token is in the cache and the time before its expiry is more than 5 minutes (300 seconds). `expiresInSeconds` is negative after the expiry, and the key is not there when the cache has no token. `scopes` comes from the `scp` claim of that token, and `reads` gives the data that the token lets you read. `refresh` is `automatic` when the shared refresh token can refresh the token, and `interactive` for the elevated token, which only `login` can capture again. `refreshToken` tells if the token cache holds the shared refresh token, and never shows the token. When it is `missing`, no token renews itself: the basic, chatsvcagg and ic3 tokens also show `interactive`, and the hint tells you to run `login`. `session` gives the end of the sign-in session: Microsoft limits it to about 24 hours after the browser sign-in, and a renewal does not move that limit. `endsAt` is an ISO time, and `endsInSeconds` is negative after the end. After the end, the basic, chatsvcagg and ic3 tokens show `interactive`, and the hint tells you to run a plain `login`. When the token cache does not hold the time of the browser sign-in, `session` gives only a `reason`: the end is not known until a browser sign-in (`login --renew`) records it. The end is a prediction only: a command still sends the refresh token to Entra. `reason` is there only when `available` is false, and it tells why the token is not available and how to get it again.',
  },
];

export const AUTH_LIFECYCLE_ENTRIES: ReadonlyArray<CommandManifestEntry> = [
  ...AUTH_SIGN_IN_ENTRIES,
  {
    name: 'docs',
    summary:
      'Print Markdown docs for one command of the auth package. Give the command name as a POSITIONAL argument: there is no `--command` flag. For `login`, `logout` and `status`, the page is the same as the --help of that command.',
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
          'Name of the command to show docs for (for example `login`). Run `{bin} --help` or `{bin} help-json` for the full list. Give it as a bare positional argument: do NOT put `--command` before it.',
      },
    ],
    example: '{bin} docs login',
    responseShape: 'Markdown text on stdout (NOT JSON-wrapped: this is the one command whose stdout is plain text).',
  },
  {
    name: 'help-json',
    summary:
      'Print the machine-readable command manifest of the auth package as JSON, with every option, example and response shape. `--terse` gives each entry as `{name, summary, category}` (the summary cut to its first sentence). Every command of this package is in the category `lifecycle`.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) renders the in-process command manifest',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/',
    options: [],
    example: '{bin} help-json',
    responseShape: '{ package, version, generatedAt, commands: [{ name, summary, category, ... }, ...] }',
  },
  {
    name: 'mcp',
    summary:
      'Start the Model Context Protocol server of the auth package over stdio. It gives an MCP client (Claude Code, Claude Desktop, Cursor and others) two tools: `login` and `status`. It runs in the foreground until the client disconnects. It is a server entry point, not a one-shot command, so an agent never calls it directly. Register it with `claude mcp add --transport stdio --scope user {bin} -- {bin} mcp`.',
    category: 'lifecycle',
    graphMethod: 'GET',
    graphPathTemplate: '(lifecycle) serves the sign-in tools over MCP stdio; not a Graph endpoint',
    graphDocsUrl: 'https://modelcontextprotocol.io/',
    options: [],
    example: '{bin} mcp',
    responseShape: 'None: a long-lived stdio server that writes JSON-RPC frames until the client disconnects (never a JSON envelope).',
  },
];
