# @ask-marcel/office-auth

> Not released yet. This package is part of the package split
> (`docs/plans/2026-10-01-package-split.md` in the repository). Until then, use
> [`ask-marcel-office-cli`](https://www.npmjs.com/package/ask-marcel-office-cli).

`@ask-marcel/office-auth` holds the Microsoft 365 sign-in for the
`@ask-marcel/office-read` and `@ask-marcel/office-write` packages: the browser
sign-in through the Teams web client, the token cache, the silent refresh and the
token helper. Its bin is `ask-marcel-office-auth`. Only this package writes the
token cache. The read and write packages get their tokens from its token helper.

## Commands

| Command | What it does |
|:--|:--|
| `ask-marcel-office-auth login` | Signs in. A browser opens the first time. When you are signed in, it shows the four tokens. `--force` captures every token again through the browser. |
| `ask-marcel-office-auth status` | Shows the four cached tokens (basic, elevated, chatsvcagg and ic3) and the time before each one expires. It makes no Graph call and opens no browser. |
| `ask-marcel-office-auth logout` | Removes the token cache and the browser profile. |
| `ask-marcel-office-auth token --tier <tier>` | The token helper (see below). |
| `ask-marcel-office-auth mcp` | Serves the MCP server over stdio (see below). |
| `ask-marcel-office-auth docs <command>` | Prints the Markdown page of one command. |
| `ask-marcel-office-auth help-json [--terse]` | Prints the command manifest as JSON. |
| `ask-marcel-office-auth --version` | Prints the version. |

The sign-in state is in `~/.ask-marcel/`: `token-cache.json` (mode 0600) and
`browser-profile/`. `ASKMARCEL_BROWSER_PROFILE` sets another folder for the
browser profile.

## The token helper

```sh
ask-marcel-office-auth token --tier <basic|elevated|chatsvcagg|ic3|guest> [--tenant <guid>] [--reject <fingerprint>]
```

The helper prints one JSON line on stdout and nothing more:

- exit 0: `{"accessToken": ..., "expiresOn": ...}`, plus `region` for `chatsvcagg` and `ic3`;
- exit 1: `{"errorCode": ..., "tier": ..., "message": ..., "remedy": ...}`, for
  example `not_authenticated`;
- exit 2: `invalid_arguments`, for a call that the helper does not accept.

`--tenant` goes with the `guest` tier only. `--reject` gives the fingerprint of a
token that Graph refused. The helper then refreshes the token only when the cache
still holds the refused one, and else gives the newer cached token. When no
person is at the terminal, the helper never opens a browser.

The line holds a bearer token. Do not write it to a log or paste it.

`dist/token.js` is the same helper as a small entry that loads no commander,
winston or update notifier, so it starts fast. Each CLI and MCP run of this
package writes the path of `dist/token.js` to `~/.ask-marcel/token-helper.json`.
The read and write packages read that file, so they find the helper without a
PATH search. Thus run `ask-marcel-office-auth login` (or `status`) once after an
install or an update.

## MCP server

```sh
claude mcp add --transport stdio --scope user ask-marcel-office-auth -- ask-marcel-office-auth mcp
```

The server has two tools, `login` and `status`. No tool result holds a token.
`logout` is for a terminal only. A browser sign-in can take more than the 60 s
that MCP clients wait by default, so set the tool timeout of your client to
about 5 minutes (`MCP_TOOL_TIMEOUT=300000`, or the same setting of your client).

## Library

```ts
import { homedir } from 'node:os';
import { resolveAuthPaths } from '@ask-marcel/office-auth';

const { tokenCache, browserProfile } = resolveAuthPaths(homedir(), process.env);
```

The library exports `resolveAuthPaths`, the type `AuthPaths` and the type
`Result`. `resolveAuthPaths` gives the same paths that the bin uses. The token
functions come in a later step of the package split.

## Develop

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

## License

MIT. See [LICENSE](LICENSE).
