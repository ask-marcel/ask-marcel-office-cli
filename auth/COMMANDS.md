# Command reference of @ask-marcel/office-auth

This page lists the commands of the `ask-marcel-office-auth` bin of `@ask-marcel/office-auth`.

The auth package has no Microsoft Graph command. It signs you in for `@ask-marcel/office-read` and `@ask-marcel/office-write`, and keeps their tokens.

`bun run docs:gen` in this folder makes the table below from the lifecycle entries of the package. CI fails when the table is out of date. Do not edit the table by hand.

<!-- AUTO-GENERATED-LIFECYCLE:BEGIN -->

| Command | What it does |
|:--|:--|
| `ask-marcel-office-auth docs` | Print Markdown docs for one command of the auth package. |
| `ask-marcel-office-auth help-json` | Print the machine-readable command manifest of the auth package as JSON, with every option, example and response shape. |
| `ask-marcel-office-auth login` | Authenticate against Microsoft Graph using the Teams web client (cached token → refresh → browser fallback). |
| `ask-marcel-office-auth logout` | Clear the cached tokens and the browser profile so the next command forces a fresh sign-in. |
| `ask-marcel-office-auth mcp` | Start the Model Context Protocol server of the auth package over stdio. |
| `ask-marcel-office-auth status` | Show the four cached tokens (basic, elevated, chatsvcagg and ic3). |

<!-- AUTO-GENERATED-LIFECYCLE:END -->

The bin has one more command, which `help-json` does not list:

| Command | What it does |
|:--|:--|
| `ask-marcel-office-auth token --tier <tier>` | The token helper of the read and write packages. It prints a token for them. You do not run it yourself. |
