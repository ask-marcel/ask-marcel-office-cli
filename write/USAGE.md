# Usage guide of @ask-marcel/office-write

This guide tells how to use the `ask-marcel-office-write` bin: the draft flow,
the rules of each draft command, the PDF conversions, the MCP server, the
library and the variables. To install the package, read [README.md](README.md).
The list of commands is in [COMMANDS.md](COMMANDS.md). The output formats, the
errors and `--output-path` are the same as in the
[read usage guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/USAGE.md).

## The draft flow

The write bin makes and changes drafts. It never sends a mail. You read the
draft, and you send it from Outlook. The ids come from the read bin:

```sh
ask-marcel-office-read search-mail-messages --query "subject:Q3 budget"
ask-marcel-office-write create-reply-draft --reply-to-message-id "AAMkAD..." --comment "Confirmed. I attach the numbers on Friday."
ask-marcel-office-read get-mail-message --message-id "AAMkAD..."
ask-marcel-office-write update-mail-draft --message-id "AAMkAD..." --comment "Confirmed. The numbers come on Friday."
```

1. Find the message with a read command (`search-mail-messages`,
   `list-mail-messages`), and copy its id.
2. Make the draft. Each command gives a short confirmation (the id, the
   subject, the recipients, the importance and `bodyPreview`), not the full
   body.
3. Read the full draft with `ask-marcel-office-read get-mail-message`.
4. Change it with `update-mail-draft` until it is correct.
5. Send it from Outlook.

To find a draft that you made before, use `ask-marcel-office-read find-mail-drafts`.

## The draft commands

### create-mail-draft

```sh
ask-marcel-office-write create-mail-draft --subject "Q3 Report" --body-content "Please review the report." --to-recipients "alice@example.com,bob@example.com"
```

The draft goes to the Drafts folder, or to the folder of `--mail-folder-id`.
Recipients are email addresses, separated by commas.

### create-reply-draft

The draft is on the thread of the message. It has the `RE:` subject, the
recipients and the quoted history. Your text goes above the quote. The draft
answers all the recipients by default. To answer only the sender, add
`--reply-all false`.

### create-forward-draft

```sh
ask-marcel-office-write create-forward-draft --forward-message-id "AAMkAD..." --to-recipients "bob@example.com" --comment "Bob owns this now."
```

The draft has the `FW:` subject and the quoted message. Your comment goes above
the quote.

### update-mail-draft

`update-mail-draft` changes only the fields that you give. Give at least one
field.

- On a draft of `create-reply-draft` or `create-forward-draft`, use `--comment`.
  It changes only the text above the quoted history, and it keeps the quote as
  it is.
- `--body-content` replaces the full body. On a draft that has a quote, the bin
  refuses it, because it removes the thread. Add
  `--replace-quoted-history true` only when you want to remove the quote.
- An empty value of a recipient flag clears that list. Use it to remove
  recipients that a reply or a forward got. When you do not give a recipient
  flag, the list does not change.

```sh
ask-marcel-office-write update-mail-draft --message-id "AAMkAD..." --cc-recipients ""
```

## The PDF conversions

`convert-mail-attachment-to-pdf`, `convert-calendar-event-attachment-to-pdf`
and `convert-group-post-attachment-to-pdf` give an attachment as a PDF. Use the
PDF when the layout of a deck is important to a vision model.

```sh
ask-marcel-office-write convert-mail-attachment-to-pdf --message-id "AAMkAD..." --attachment-id "AAMkAD..." --output-path ./deck.pdf
```

- A file attachment goes to a temporary file in the `.ask-marcel-temp` folder of
  your OneDrive. Graph converts it, and then the bin tries to delete that file.
  A failed cleanup can leave the file in that folder.
- A reference attachment (a link to a OneDrive or SharePoint file) is converted
  where it is, with no upload.
- A plain-text source comes back as its own bytes, with `passthrough: true`.
- A PDF source comes back as its own bytes, with no `passthrough` field. Graph
  does not convert a PDF to a PDF.
- An Outlook item (an attached mail, event or contact) is refused. Use
  `convert-mail-attachment-to-markdown` of the read bin for it.
- An image is refused with `api_error` 415, in all three conversions.
- A large mail attachment can take a long time: up to about 22 minutes in the
  worst case.

Use `--output-path <file>` to write the PDF to a file. Without it, the text
output gives one line with the type and the size, and no bytes.

## MCP server

`ask-marcel-office-write mcp` serves the write commands to an MCP client over
stdio. The [README](README.md#mcp-server) gives the lines that register it.

```text
list-commands { category? }                                      → the short manifest
get-command-docs { command }                                     → the full docs of one command
run-write-command { command, params?, outputPath?, outputDir? }  → runs a write command
```

`run-write-command` runs the 7 write commands. It has `readOnlyHint: false`.
`params` are the flags without the `--` prefix, in camelCase, and each value is
a string. A read command gets a pointer to the read server
(`ask-marcel-office-read mcp`). Register this server only on hosts that may
write drafts.

## Library

```ts
import { commands, createWriteClient } from '@ask-marcel/office-write';

const graph = createWriteClient();
const draft = await commands['create-mail-draft']?.execute(graph, {
  subject: 'Q3 Report',
  bodyContent: 'Please review the report.',
  toRecipients: 'alice@example.com',
});
if (draft?.ok) console.log(draft.value);
```

| Export | What it is |
|:--|:--|
| `commands` | Each of the 7 write commands, by name. Each one has `meta` and `execute(graph, params)`. |
| `createWriteGraph(tokens)` | The write graph over a `TokenSource`: the basic token, with `get`, `post`, `patch`, `put` and `delete`. |
| `createEnvThenHelperTokenSource()` | The token source of the bin: the tier variables of the process, else the token helper of the auth package. |
| `createWriteClient()` | `createWriteGraph(createEnvThenHelperTokenSource())`, the graph that the bin uses. |
| `Command`, `WriteGraph`, `TokenSource`, `Result` | The types of the values above. |

`meta.effect` is `draft` for the mail-draft operations and `transient-upload`
for the PDF conversions.

## Your own token

The write bin and the library get each token as the read package does. They
read the `ASKMARCEL_TOKEN_BASIC` variable first. Else they run the token helper:
`ASKMARCEL_TOKEN_COMMAND`, else the locator file, else `ask-marcel-office-auth`
on PATH. The
[read usage guide](https://github.com/ask-marcel/ask-marcel-office-cli/blob/main/read/USAGE.md#your-own-token)
gives the details.

```sh
ASKMARCEL_TOKEN_COMMAND=/opt/vault/token-helper ask-marcel-office-write create-mail-draft --subject "Test" --body-content "Test" --to-recipients "alice@example.com"
```

## Variables

| Variable | What it does | Default |
|:--|:--|:--|
| `ASKMARCEL_LOG_LEVEL` | The log level. All logs go to stderr. | `error` |
| `ASKMARCEL_BINARY_TIMEOUT_MS` | The time for one binary download (a PDF), in milliseconds | `300000` (5 minutes) |
| `ASKMARCEL_TOKEN_COMMAND` | The absolute path of a token helper | the locator file, then PATH |
| `ASKMARCEL_TOKEN_BASIC` | The basic token | not set |
| `HOME` / `USERPROFILE` | The folder that holds `~/.ask-marcel/` | required |
