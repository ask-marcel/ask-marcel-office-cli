# @ask-marcel/office-read

`@ask-marcel/office-read` holds the 204 read commands: the Microsoft Graph
commands that change nothing in your tenant. They read mail, calendar, OneDrive
and SharePoint, Teams and chats, users, Excel, OneNote and To Do. They convert
Office files, PDF and mail to markdown on the fly. Its bin is
`ask-marcel-office-read`. It cannot write: the mail drafts and the three
attachment-to-PDF conversions are in `@ask-marcel/office-write`.

You do not need an Azure app registration, the consent of a tenant admin or a
client secret. `@ask-marcel/office-auth` signs you in with your own Microsoft 365
account, in a browser, one time.

## Install

Install it with `@ask-marcel/office-auth`, which signs you in. Add
`@ask-marcel/office-write` only on hosts that may write drafts.

```sh
npm i -g @ask-marcel/office-auth @ask-marcel/office-read
ask-marcel-office-auth login
```

The package needs Node 20 or later, or Bun 1.0 or later, on macOS, Windows or
Linux.

## Quick start

```sh
ask-marcel-office-read my-quick-context
ask-marcel-office-read list-mail-messages --top 5
ask-marcel-office-read list-calendar-view --start-date-time today --end-date-time +7d
ask-marcel-office-read convert-local-file-to-markdown --path ./report.docx
```

`my-quick-context` gives your user, the ids of your drive, inbox, calendar,
first plan and notebook, and your time zone, in one call. The last line reads a
file on your disk. It runs offline, with no sign-in.

## The difference in practice

| You need | Raw Microsoft Graph | `@ask-marcel/office-read` |
|---|---|---|
| Access | An app registration, the consent of an admin, the rotation of a secret | `ask-marcel-office-auth login`: one browser sign-in with your own account |
| A mail and its attachments | 6 or more round trips. You convert the HTML to text. | `convert-mail-to-markdown`: one call, markdown out |
| A pptx, a pdf or a legacy .doc as text | Raw bytes. You bring a converter. | `download-drive-item-as-markdown`: one call |
| A useful error | `BadRequest: Invalid filter clause` | `hint: "string literals MUST use single quotes; embed one by doubling it"` |
| A PDF of 5 MB | Base64 that fills the context window | `--output-path` writes it to disk. The model reads 3 lines. |
| Throttling | HTTP 429. You guess the backoff. | `retryAfterSeconds` in the error envelope |

Compared with the tools that you know:

| | Microsoft Graph SDKs | CLI for Microsoft 365 (PnP) | `@ask-marcel/office-read` |
|---|---|---|---|
| Sign-in | An Entra app registration, always. Each auth flow needs a client id. | An Entra app registration. Its `setup` command makes one, and its permissions can need the consent of an admin. | Your own account, in a browser, one time. No app registration. |
| Made for | Backend services | Admins and scripts | LLMs that call tools: markdown out, small fields, a repair hint on each error |
| Can it break things | It can do what the permissions of your app allow. | Yes: it creates, changes and deletes items in the tenant. | No. The read package cannot write. The mail drafts and the three attachment-to-PDF conversions are in `@ask-marcel/office-write`, which you install only where you need it. |

## The three walls it removes

### Sign in as a person, not as an app

Microsoft Graph usually needs an Azure app registration, the consent of a
tenant admin and client secrets before the first call. Here,
`ask-marcel-office-auth login` opens a real browser window on the usual
Microsoft sign-in (Playwright does this). It captures the same token that the
Teams web client uses. A personal or an enterprise Microsoft 365 account works.
The tokens are in `~/.ask-marcel/token-cache.json` (mode 0600), and they refresh
without a browser. `ask-marcel-office-auth status` shows the scopes and the
expiry of each token, with no Graph call and no browser. The read package gets
its tokens from the auth package and keeps no token cache.

### Safe for an autonomous agent

None of the 204 read commands changes your tenant. The read package has no
`send-mail`, no `create-event`, no `upload-file` and no delete command. Its graph
sends a POST only to the search and free/busy endpoints, and refuses each other
POST before it goes out. Thus an agent can explore a mailbox, and you do not
have to review each call. The package sends no analytics. It talks only to
Microsoft 365 and to npm. In Microsoft 365, it talks to Graph, and the chat
commands also talk to the Teams chat and media services. It asks npm only for
the latest version.

### Answers sized for a context window

Graph answers are made for backend services. These answers are made for
models. A listing gives a set of selected fields, not each property of each
item. Use `--select` to get other fields. A mail thread comes without its quoted
reply chain, which makes real threads 78-90% smaller. An inline image becomes a
`[inline image: logo.png]` placeholder, not kilobytes of base64. A binary answer
goes to disk through `--output-path`. Each failure gives
`{error, hint, source, retryAfterSeconds}`, so the model can repair its call.

## See it work

To read a mail and its attachments with raw Graph, you get the message, the
list of attachments and the bytes of each attachment. Then you resolve each
SharePoint link of the body and convert the HTML to text. That is six round
trips or more. Here, `convert-mail-to-markdown --message-id <id>` does it in one
call. The same converter reads an Outlook `.msg` file on your disk, with no
sign-in. This is the output for a sample file:

```console
$ ask-marcel-office-read convert-local-file-to-markdown --path ./mail.msg
# Quarterly Report — Q3 Summary

**From:** Jordan Avery <jordan.avery@example.com>
**Recipients:** Sam Rivera <sam.rivera@example.com>

Hi Team,

Please find the quarterly figures attached.

Best regards,
Jordan

## Attachments

### summary.txt

Attachment body text — line one.
Line two.
```

Each attachment is converted in the same answer. For a mail in your mailbox,
the attachments come with their ids, ready for the next call:

```console
$ ask-marcel-office-read convert-mail-attachment-to-markdown --message-id "AAMkAD..." --attachment-id "AAMkAD..."
```

An xlsx comes back as one markdown table for each sheet. When the answer is in
a deck that a person shared three weeks ago:

```console
$ ask-marcel-office-read microsoft-search-query --query "Q3 budget filetype:pptx"
$ ask-marcel-office-read download-drive-item-as-markdown --drive-id "b!abc..." --item-id "01BYE..."
```

## What your agent can reach

| Surface | In practice |
|---|---|
| Outlook mail | Search and read mail as markdown. Convert each attachment, also in a nested `.zip`, `.msg`, `.eml` or forwarded mail. Resolve the SharePoint links of a body. Get your signature. Find the drafts of a thread. |
| OneDrive and SharePoint | Find each drive and site that your token can reach. Search files. Read a document as markdown or PDF. Read the version history, and the line diff between two versions or two files. Read the comments of a document, the files changed since a date and who can open a site. Resolve share links, also into partner tenants where you are a guest. |
| Calendar | "What is on this week" with relative dates (`today`, `start-of-week`, `+7d`), the details of an event and free/busy lookups |
| People and directory | People search, user profiles, the directory around you, and your own identity and ids in one call |
| Teams | Your teams, channels, channel posts as JSON or as a markdown transcript, what changed since a date, members and tabs, chats, and the chat history with files, reactions and pasted screenshots |
| Planner and To Do | Plans (also of a group) and their label names, buckets, tasks, checklists, due dates and the web link of each To Do task |
| Excel | Live workbook reads: worksheets, used ranges and tables |
| OneNote | Notebooks, sections and the content of pages |
| Search and utilities | Microsoft Search across the tenant, pagination with `next-page`, and the conversion of a local file offline |

[COMMANDS.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/COMMANDS.md)
gives each command, with its parameters and its Graph endpoint.

## Any file becomes markdown

One conversion pipeline has four entry points: a OneDrive or SharePoint item,
an Outlook attachment, a `.zip` archive, and a file on your disk
(`convert-local-file-to-markdown` runs offline, with no sign-in).

| Source | What comes back |
|---|---|
| docx, docm, dotx | Markdown. `--include-metadata true` adds the comments, the tracked changes, the hidden text and a flag for VBA macros. |
| xlsx, xlsm, csv | One markdown table for each sheet. A sheet above 50 000 cells gives a plan to read it in bands. |
| pptx, pptm | The text of each slide, with the speaker notes (`## Slide N`). When the layout is important, use the PDF of the deck and a vision model. |
| pdf | The text layer. A scanned PDF gives a pointer to the vision route. |
| odt, ods, odp | Headings, lists, tables, the text of each slide, and the comments |
| Legacy .doc, .xls | `.xls` reads as `.xlsx`. `.doc` gives plain text. Convert a `.ppt` to PDF first. |
| Outlook .msg, .eml | The full mail: the headers, the body without the quoted chain, and each attachment, converted. An Outlook item in a mail (a forwarded mail) reads the same way. |
| html, htm | The page as markdown, without its head, and the embedded images as placeholders. A page above 1 MB becomes plain text. |
| .zip | Each entry, converted in one call. GBK and CP437 entry names are decoded correctly. |
| Loop, Whiteboard | Graph converts them on the server. |

To get the pictures, not the words: `extract-drive-item-images`,
`extract-mail-attachment-images` and `extract-local-file-images` get the
embedded images of a docx, xlsx, pptx or pdf (also the originals at full
resolution and the images on hidden slides), ready for a vision model. On the
drive and mail extractors, `--pages 1-3` limits a scanned PDF. The mail
extractor also opens the files of a forwarded mail. `extract-teams-chat-message-images` gets the pasted screenshots
of a chat message. `--output-dir` writes the images to disk.

## Made for the agent loop

- **It teaches itself.** `help-json --terse` gives a small JSON manifest for the
  first contact of a model. Add `--category mail` to get one category.
  `docs <command>` prints the full docs of one command: the response shape, the
  examples and the Graph endpoint.
- **Errors that repair the call.** Each failure is
  `{ok: false, error, errorCode?, hint?, source, retryAfterSeconds?}`. Hints
  cover more than 20 frequent Graph mistakes. For example, a Teams URL that you
  give to `resolve-mail-link` gets the name of the correct command. A wrong
  command, flag or parameter ends with "Did you mean …?", on the CLI, over MCP
  and in the library.
- **Pagination with no state.** Each listing gives a cursor.
  `next-page --url "<cursor>"` continues each of them.
- **Relative dates.** `--start-date-time start-of-week --end-date-time +7d`. Named
  days start at midnight in your time zone (`--tz`, else `ASKMARCEL_TZ`, else the
  zone of the machine).
- **For pipes.** `--output raw-json` prints the payload alone, with no envelope,
  for `jq`.
- **Binary control.** A binary answer of many MB never goes to stdout by
  accident. Without `--output-path`, you get a summary of one line. With it, the
  bytes go to disk, and the envelope has `savedTo`.
- **Reach across the tenant.** `list-accessible-drives` joins each way in which a
  delegated token can find a drive. These are the Teams libraries, the group
  sites, the private channels, the files shared with you, your activity and the
  other site libraries.
  `search-all-accessible-sites` reads all the pages of the search index. This is
  the most that you can reach with no tenant-admin rights.

## Commands

| Command | What it does |
|:--|:--|
| `ask-marcel-office-read <command>` | Runs one of the 204 read commands, for example `list-mail-messages` or `convert-local-file-to-markdown`. |
| `ask-marcel-office-read --help` | Lists the commands with a one-sentence summary each. |
| `ask-marcel-office-read docs <command>` | Prints the Markdown page of one command. |
| `ask-marcel-office-read help-json [--terse] [--category <name>]` | Prints the command manifest as JSON. |
| `ask-marcel-office-read mcp` | Starts the MCP server over stdio (see below). |
| `ask-marcel-office-read --version` | Prints the version. |

`help-json` lists 207 entries: the commands, and `docs`, `help-json` and `mcp`.
The 204 read commands are 199 GET reads and 5 read-only POST (four Microsoft
Search queries and one free/busy lookup). The tables of each category are in
[COMMANDS.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/COMMANDS.md).

A write command (for example `create-mail-draft`) or a sign-in command (for
example `login`) is not in this package. The bin tells you which package has it
and gives the command line.

## Sign-in

The read package never opens a browser and keeps no token cache. For each
token, it reads the `ASKMARCEL_TOKEN_<TIER>` variable first (for example
`ASKMARCEL_TOKEN_BASIC`). Else it runs the token helper of
`@ask-marcel/office-auth`: `ASKMARCEL_TOKEN_COMMAND`, else the helper that
`~/.ask-marcel/token-helper.json` names, else `ask-marcel-office-auth token` on
PATH. Only the auth bin writes that file, so run `ask-marcel-office-auth login`
(or `status`) one time after an install or an update.

When a token lapses, run `ask-marcel-office-auth login` (no flag). A plain
`login` captures a missing elevated token again and keeps the browser session.
Use `ask-marcel-office-auth login --force` only when a tier stays stuck after a
plain `login`: it clears the browser session, so you usually type your password
again.

## MCP server

The read bin is also an MCP server over stdio. Register it in Claude Code with
the global bin:

```sh
claude mcp add --transport stdio --scope user ask-marcel-office-read -- ask-marcel-office-read mcp
```

Or with npx, with no global install:

```sh
claude mcp add --transport stdio --scope user ask-marcel-office-read -- npx -y @ask-marcel/office-read mcp
```

Claude Desktop, Cursor, Cline and other MCP clients take the same `mcpServers`
block. Only the location of the file is different:

| Client | Config file |
|:--|:--|
| Claude Desktop (macOS) | `~/Library/Application Support/Claude/claude_desktop_config.json` |
| Claude Desktop (Windows) | `%AppData%\Claude\claude_desktop_config.json` |
| Cursor, Cline, Windsurf and others | See the MCP docs of your client. The block does not change. |

```json
{
  "mcpServers": {
    "ask-marcel-office-read": {
      "command": "npx",
      "args": ["-y", "@ask-marcel/office-read", "mcp"]
    }
  }
}
```

The first start downloads the package into the npx cache, so it is slow one
time. An app that you start from the Dock or the Start menu (Claude Desktop) does
not get the PATH of your shell. If the server does not start, put the absolute
path that `which npx` prints in `"command"`. The npx cache has one more risk
for the sign-in: read the npx note in the
[auth usage guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/auth/USAGE.md#the-npx-cache).

The server has three tools, not one tool for each command:

| Tool | What it does |
|:--|:--|
| `list-commands` | Gives the short manifest. Start here. `category` makes the list smaller. |
| `get-command-docs` | Gives the full docs of one command: options, endpoint and example. |
| `run-command` | Runs one of the read commands. It has `readOnlyHint: true`, so a client can approve it automatically. |

A write command gets a pointer to the write server
(`ask-marcel-office-write mcp`), and `login` gets a pointer to the auth server.
Sign in from a terminal first, with `ask-marcel-office-auth login`.

## Library

```ts
import { commands, createReadClient } from '@ask-marcel/office-read';

const graph = createReadClient();
const drives = await commands['list-drives']?.execute(graph, {});
if (drives?.ok) console.log(drives.value);
```

The library exports four values and four types:

- `commands`: the 204 read commands, each on the read graph alone;
- `createReadGraph(tokens)`: the read graph over a token source (every token
  tier, and a POST that reaches only the search and free/busy endpoints);
- `createEnvThenHelperTokenSource()`: the tier variables of the process, else the
  token helper of `@ask-marcel/office-auth`;
- `createReadClient()`: the two above together, the same read graph that the
  bin uses;
- the types `Command`, `ReadGraph`, `TokenSource` and `Result`.

Each `execute` returns a `Result`: `{ ok: true, value }` or `{ ok: false, error }`.
`@ask-marcel/office-read/commands.json` is the command manifest.

`convert-local-file-to-markdown` and `extract-local-file-images` read a file on
your disk. The library exports no file system, so run these two commands from the
bin. To give the library your own token, set a tier variable or
`ASKMARCEL_TOKEN_COMMAND` (see the
[usage guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/USAGE.md)).

## More docs

- [USAGE.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/USAGE.md):
  the output formats, OData options, relative dates, `--output-path`,
  pagination, errors, the MCP server, the library and the variables.
- [COMMANDS.md](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/COMMANDS.md):
  the commands of each category, with their parameters and Graph endpoints.
- [@ask-marcel/office-auth](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/auth/README.md):
  the sign-in.
- [@ask-marcel/office-write](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/write/README.md):
  the mail drafts and the three attachment-to-PDF conversions.

## Develop

The [contributor guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/CONTRIBUTING.md)
tells how the workspace, the pre-commit hook and CI work.

This folder is one Bun workspace of the repository. Run `bun install` at the
repository root once. In a new checkout, run the auth bin once
(`bun run src/main.ts status` in `auth/`), so that the locator names a token
helper. Then, from this folder:

| Command | What it does |
|:--|:--|
| `bun run src/main.ts <command>` | Runs the bin from the source. |
| `bun test` | The tests of this folder |
| `bun run lint` / `bun run lint:strict` | ESLint (fast / type-aware) |
| `bun run typecheck` | `tsc --noEmit` |
| `bun run coverage` | 100% on every tier of `src/` |
| `bun run mutate:changed` | Stryker on the changed domain and use-case files |
| `bun run build` | `dist/`: `cli.js` and `index.js` (core inlined), the declarations and `commands.json` |
| `bun run check:packed` | Packs the package and checks the tarball and a clean consumer install |
| `bun run smoke` | After the build: runs `dist/` under node and bun on every converter, the MCP server and the read bin (`scripts/qa-bundle-smoke.ts`) |
| `bun run smoke:packed` | After the build: installs the packed tarball in an empty folder and runs the smoke on it (`scripts/qa-packed-install-smoke.ts`) |
| `bun run docs:gen` | Makes the tables of `COMMANDS.md` from the registry |
| `bun run check:docs` | Fails when `COMMANDS.md` or `dist/commands.json` is out of date. It also fails when a count in `README.md`, `USAGE.md` or `COMMANDS.md` does not match the registry, or when a count in `README.md` or `USAGE.md` has no claim. |

## License

MIT. See [LICENSE](LICENSE).
