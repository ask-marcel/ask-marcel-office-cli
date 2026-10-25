# @ask-marcel/office-auth

`@ask-marcel/office-auth` holds the Microsoft 365 sign-in for the
`@ask-marcel/office-read` and `@ask-marcel/office-write` packages: the browser
sign-in through the Teams web client, the token cache, the silent refresh and the
token helper. Its bin is `ask-marcel-office-auth`. Only this package writes the
token cache. The read and write packages get their tokens from its token helper.
You do not need an Azure app registration or the consent of a tenant admin.

## Install

Install it with the read package. Add the write package only on hosts that may
write mail drafts.

```sh
npm i -g @ask-marcel/office-auth @ask-marcel/office-read
ask-marcel-office-auth login
```

The package needs Node 20 or later, or Bun 1.0 or later.

## Commands

| Command | What it does |
|:--|:--|
| `ask-marcel-office-auth login` | Signs you in. A browser opens the first time. When you are signed in, it shows which tokens are available. |
| `ask-marcel-office-auth status` | Shows the four cached tokens (basic, elevated, chatsvcagg and ic3) and the time before each one expires. It makes no Graph call and opens no browser. |
| `ask-marcel-office-auth logout` | Removes the token cache and the browser profile. |
| `ask-marcel-office-auth token --tier <tier>` | The token helper of the read and write packages. You do not run it yourself. |
| `ask-marcel-office-auth mcp` | Starts the MCP server over stdio (see below). |
| `ask-marcel-office-auth docs <command>` | Prints the Markdown page of one command. |
| `ask-marcel-office-auth help-json [--terse]` | Prints the command manifest as JSON. |
| `ask-marcel-office-auth --version` | Prints the version. |

`help-json` lists 6 entries: `login`, `logout`, `status`, `docs`, `help-json`
and `mcp`. `token` is not a `help-json` entry. The full list is in
[COMMANDS.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/auth/COMMANDS.md).

## Sign in

1. Run `ask-marcel-office-auth login`. A browser opens one time. Sign in with
   your usual Microsoft 365 account.
2. When a token lapses, run `ask-marcel-office-auth login` (no flag) again. A
   plain `login` captures a missing elevated token again and keeps the browser
   session, so you usually do not type your password again.
3. Use `ask-marcel-office-auth login --force` only when a tier stays stuck after
   a plain `login`. It clears the browser session, so you usually type your
   password again.

The basic token refreshes itself from the refresh token. The elevated token
has no refresh token, and it lapses after about an hour. Only `login` can
capture it again. `ask-marcel-office-auth status` shows each token and the
time before it expires.

The sign-in state is in `~/.ask-marcel/`: `token-cache.json` (mode 0600) and
`browser-profile/`. `ASKMARCEL_BROWSER_PROFILE` sets another folder for the
browser profile. Each run of the bin also writes
`~/.ask-marcel/token-helper.json`, which tells the read and write packages where
the token helper is. Thus run `ask-marcel-office-auth login` (or `status`) one
time after an install or an update.

The token helper, the token tiers, the variables and the error codes are
in [USAGE.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/auth/USAGE.md).

## MCP server

Register the server with the global bin:

```sh
claude mcp add --transport stdio --scope user ask-marcel-office-auth -- ask-marcel-office-auth mcp
```

Or with npx, with no global install:

```sh
claude mcp add --transport stdio --scope user ask-marcel-office-auth -- npx -y @ask-marcel/office-auth mcp
```

Other MCP clients (Claude Desktop, Cursor, Cline) take an `mcpServers` block:

```json
{
  "mcpServers": {
    "ask-marcel-office-auth": {
      "command": "npx",
      "args": ["-y", "@ask-marcel/office-auth", "mcp"]
    }
  }
}
```

The server has two tools, `login` and `status`. No tool result holds a token.
`logout` is for a terminal only. A browser sign-in can take more than the 60 s
that MCP clients wait by default. Thus set the tool timeout of your client to
about 5 minutes (`MCP_TOOL_TIMEOUT=300000`, or the same setting of your client).
Sign in from a terminal first. Read the npx note in
[USAGE.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/auth/USAGE.md#the-npx-cache)
before you use the npx line.

## Library

```ts
import { homedir } from 'node:os';
import { resolveAuthPaths } from '@ask-marcel/office-auth';

const { tokenCache, browserProfile } = resolveAuthPaths(homedir(), process.env);
```

The library exports `resolveAuthPaths`, the type `AuthPaths` and the type
`Result`. `resolveAuthPaths` gives the same paths that the bin uses. The library
has no token functions. To get a token from your own code, run the token helper
(see [USAGE.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/auth/USAGE.md#the-token-helper)).

## Develop

The [contributor guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/CONTRIBUTING.md)
tells how the workspace, the pre-commit hook and CI work.

This folder is one Bun workspace of the repository. Run `bun install` at the
repository root once, then from this folder:

| Command | What it does |
|:--|:--|
| `bun run src/main.ts <command>` | Runs the bin from the source. A run records `src/token.ts` as the token helper. |
| `bun test` | The tests of this folder |
| `bun run lint` / `bun run lint:strict` | ESLint (fast / type-aware) |
| `bun run typecheck` | `tsc --noEmit` |
| `bun run coverage` | 100% on every tier of `src/` |
| `bun run mutate:changed` | Stryker on the changed domain and use-case files |
| `bun run build` | `dist/`: `cli.js`, `token.js` and `index.js` (core inlined), and the declarations |
| `bun run check:packed` | Packs the package and checks the tarball and a clean consumer install |
| `bun run smoke` | After the build: runs `dist/` under node and bun on temporary homes (`scripts/qa-auth-bundle-smoke.ts`) |
| `bun run docs:gen` | Makes the table of `COMMANDS.md` from the lifecycle entries |
| `bun run check:docs` | Fails when `COMMANDS.md` is out of date. It also fails when a count in `README.md`, `USAGE.md` or `COMMANDS.md` does not match the registry, or when a count in `README.md` or `USAGE.md` has no claim. |

## License

MIT. See [LICENSE](LICENSE).
