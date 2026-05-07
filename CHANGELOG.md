# Changelog

All notable changes to `ask-marcel-office-cli` are documented here.

## 1.0.0

The first stable release. Two breaking changes consolidate the public output
contract; the rest is additive.

### Breaking — output contract

- **Errors emit on stdout**, not stderr. `process.exitCode = 1` still
  distinguishes failure, so shell scripts that branch on the exit code keep
  working — but anything that read errors from stderr (`cmd 2>err.json`) needs
  to merge streams or read stdout instead. An LLM piping `ask-marcel <cmd>
  | jq` no longer needs `2>&1`.
- **Every command output is wrapped in the v1 envelope**:
  - Success: `{ ok: true, data: <payload>, nextLink?: string, count?: number }`
  - Error: `{ ok: false, error: "<message>" }`

  `@odata.nextLink` and `@odata.count` from the underlying Graph payload are
  lifted to the top of the envelope and removed from `data`. Consumers who
  parsed `value[0]` as the first item now read `data.value[0]`.

### Added — OData query passthrough on every list/search command

Every `list-*` / `search-*` / `get-*-delta` command now accepts the six
standard OData query parameters as optional flags, so an LLM can shrink large
responses on the fly:

```
--top <n>       maximum items per page
--skip <n>      offset
--select <csv>  comma-separated field list (huge payload-size win)
--filter <kql>  server-side predicate
--orderby <kql> sort expression
--expand <nav>  inline navigation properties
```

Four commands keep `buildCommand` because their hard-coded `$filter` would
collide with a user-supplied `--filter`: `list-conversation-messages`,
`list-incomplete-todo-tasks`, `list-incomplete-planner-tasks`,
`search-onenote-pages`.

### Added — `my-quick-context`

New meta command that issues five Graph calls in parallel (`/me`, `/me/drive`,
`/me/mailFolders/inbox`, `/me/todo/lists`, `/me/calendar`) and returns
`{ user, primaryDriveId, inboxId, todoLists, primaryCalendarId }` in one
round trip. Replaces the audit's 5-call discovery chain.

### Fixed

- `microsoft-search-query` no longer 400s. Splits `entityTypes` into two
  `requests[]` entries so Graph stops rejecting `person` mixed with
  file/mail/event types.
- `list-conversation-messages` no longer trips Graph's `InefficientFilter`.
  Drops the `$orderby=receivedDateTime` from the OData query.
- `list-sharepoint-site-items` is removed. Microsoft Graph has no list-less
  site/items collection endpoint; `get-sharepoint-site-item`'s docstring now
  points at the two-step discovery chain
  (`list-sharepoint-site-lists` → `list-sharepoint-site-list-items`) that
  Graph actually supports.
- `list-groups` summary no longer advertises a `--top` flag it didn't
  register. Project-wide invariant added so every `--flag` mentioned in any
  command summary must be a real option or alias on that command.
- `next-page` routes nextLinks under `/me/chats` and `/chats/...` via the
  elevated M365ChatClient token. Chat pagination no longer 403s.
- `search-onenote-pages` accepts `--query` as an alias for
  `--title-substring`, matching the convention used by every other search
  command.

### Added — flag aliases

- `--todo-list-id` is now accepted by every command that takes
  `--todo-task-list-id` (`--task-list-id` alias preserved).
- `get-sharepoint-site-item` accepts `--list-item-id` (alias for `--item-id`).
  `get-sharepoint-site-list-item` accepts `--item-id` (alias for
  `--list-item-id`). LLMs that write either spelling from memory now hit the
  right flag.

### Quality

- Bun `JSON.stringify` already escapes every U+0000–U+001F control character
  and U+2028 / U+2029 separator. The audit's "raw control chars" claim in
  the four insight commands does not reproduce against the actual code path;
  regression-guard tests pin the contract.

## Older

Earlier history is in the git log. See `git log --oneline` for individual
commits up to and including v0.11.0.
