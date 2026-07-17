# Mail-draft feature set: HTML reply/forward, --comment revision, reply-all toggle, clear-list semantics, get-mail-signature

Status: **implemented 2026-07-17**. Kept as the design record — the reasoning below is why the
shipped shape is what it is; the commit messages carry only the conclusions.

All ten planned commits landed, in order, under the messages planned for them:

| Planned | Shipped |
|:--|:--|
| C1 `refactor(mail): export quote-boundary finders from mail-quote-stripper` | `ad12400` |
| C2 `feat(mail): draft-comment-splicer pure module` | `734404e` |
| C3 `feat(mail): --reply-all toggle on create-reply-draft` | `a40af45` |
| C4 `feat(mail): --body-content-type HTML splice on create-reply-draft` | `391bcfa` |
| C5 `feat(mail): --body-content-type HTML splice on create-forward-draft` | `3348130` |
| C6 `fix(mail): defined clear-list semantics on update-mail-draft` | `790d376` |
| C7 `feat(mail): update-mail-draft --comment revises the reply text above the quote` | `85e030d` |
| C8 `feat(mail): signature-extractor helper` | `c777409` |
| C9 `feat(mail): get-mail-signature command` | `e820f51` |
| C10 `docs(mail): regenerate manifest for get-mail-signature; bump command counts` | `b954b57` |

What the plan did not anticipate, recorded in `.claude/LESSONS.md` (2026-07-17): Graph does not always
emit `appendonsend`, so `QUOTE_BOUNDARIES`' marker families are load-bearing rather than redundant;
and the size gate and the mutation gate pull in opposite directions, which is why C9 became four
commits instead of one.

(The original `Base: HEAD 3c4efe7` line is dropped — that commit lived on a pre-purge branch that has
since been deleted and garbage-collected, so the reference no longer resolves.)

## Context

A 2026-07-16 field report from the ask-marcel skill (which drives drafts through the CLI binary only) lists four gaps instruction-driven use cannot work around:

1. Reply/forward comments are plain-text only. The reporter believed `bodyContentType` exists at the library level; that was published v2.2.0 behavior, where reply/forward PATCHed `body` and silently destroyed the quoted history (regression fixed 2026-07-13 by removing the field). HEAD has zero HTML support on these two commands, and Graph's `comment` param HTML-escapes its input.
2. Graph-created drafts carry no signature; callers need the user's `id="Signature"` block from a recent sent message, with cid: logos inlined as base64.
3. Reply is reply-all only, and `update-mail-draft --cc-recipients ""` is silently treated as "flag absent" (truthiness gates), so an inherited CC list can never be cleared.
4. `update-mail-draft --body-content` replaces the whole body, so revising the comment on a threaded draft wipes the quoted history.

Decisions confirmed: splice+PATCH for HTML (relaxing "never PATCH body" to "never PATCH body without the quote in it"); empty string `""` clears recipient lists; signature ships as a read-only `get-mail-signature` command (not a `--with-signature` flag); reply-to-sender ships as `--reply-all true|false` on `create-reply-draft` (default true, stance in summary rewritten).

Registry grows 182 → 183.

## Core mechanism (requests 1 + 4 share it)

Graph reply/forward draft bodies are full HTML docs (`<html><head><style>…</style></head><body dir="ltr">…`). The `<head>` styles are load-bearing for the quoted tail, so:

- **F1 (create with HTML comment)** = pure insert, no head surgery: `html.slice(0, cut) + commentHtml + html.slice(cut)` where `cut = findQuoteBoundary(html)`. Everything Graph minted (doctype, head, `<body>` open tag, the empty-comment head) is kept verbatim. Boundary not found → insert right after the `<body…>` open tag (fallback index 0 for fragments); nothing is ever dropped.
- **F4 (revise comment on existing draft)** = replace `[findBodyInsertStart(html), cut)`. Replace, not insert, or repeated edits accumulate stale comments. Boundary not found → actionable error ("no quoted history to preserve; use --body-content"), because replacing an unbounded head is exactly the 2026-07-13 regression class.
- **Boundary-marker injection guard**: if the user's raw HTML comment itself matches a quote boundary (pasted `gmail_quote` div, `border-top:solid #E1E1E1`, bold From:+Sent: pair), the NEXT `--comment` edit would cut at the injected marker and leave ghost content. Both raw-HTML paths must reject `findQuoteBoundary(commentHtml) !== -1` with a validation_error, before any Graph call.
- Idempotency holds because all boundary markers (`appendonsend`, `divRplyFwdMsg`, localized header block) live in the tail we keep. Exchange SafeHTML on PATCH preserves these structural ids (OWA round-trips drafts through the same pipeline); live verification step 1/2 proves it end-to-end.

## Commits (each ≤10 files, ≤300 lines incl. tests + regenerated docs; TDD red-first; conventional messages; run `bun run docs:gen` inside every commit touching a `meta` block)

### C1 `refactor(mail): export quote-boundary finders from mail-quote-stripper`
Files (2): `src/use-cases/commands/mail-quote-stripper.ts`, new `mail-quote-stripper.test.ts` (direct tests currently live only in commands.test.ts ~4578-4715; leave those untouched as regression cover).
- Extract the private cut-finding into exported `findQuoteBoundary(html): number` (earliest of `QUOTE_BOUNDARIES` merged with the **widened** confirmed-header index, or -1) and `findPlainTextQuoteBoundary(text): number`. `stripQuotedReplies`/`stripQuotedPlainText` become delegators, behavior byte-identical.
- Tests: exact index for `appendonsend` preceding `divRplyFwdMsg` (earliest-wins); -1 on marker-free HTML (bare `<blockquote>`); bold From/Sent fixture returns the widened `<p>` index (kills widenToBlockStart-bypass mutants); plaintext underscore-rule / "On … wrote:" / -1 / earliest-of-several; delegation equivalence `stripQuotedReplies(h).html === h.slice(0, findQuoteBoundary(h)) + STRIP_MARKER`.

### C2 `feat(mail): draft-comment-splicer pure module`
Files (2): new `src/use-cases/commands/draft-comment-splicer.ts` + test.
- `escapeTextAsHtml` (escape `& < > " '`, `\r\n|\n` → `<br>`, no wrapper div), `findBodyInsertStart` (`/<body[^>]*>/i` match end, else 0), `insertCommentAboveQuote(html, commentHtml)` (F1 op + body-open fallback), `replaceCommentAboveQuote(html, commentHtml)` (F4 op; `boundaryFound:false` → input unchanged), `replacePlainTextCommentAboveQuote(text, comment)` (`comment + '\n\n' + tail`). Each returns `{ html|text, boundaryFound }`. Imports the C1 finders.
- Tests: exact-string assertions on a realistic full-doc fixture (head+style, `<body dir="ltr">`, empty head div, `appendonsend`, `divRplyFwdMsg`, tail). Escaping in one assertion covering all five chars + both newline forms; uppercase `<BODY>`; fragment → 0; insert preserves head/styles/empty div; replace removes the minted empty div, tail intact from `appendonsend`; plaintext exact separator.

### C3 `feat(mail): --reply-all toggle on create-reply-draft` (request 3a)
Files (4): `src/use-cases/commands/create-reply-draft.ts`, its test, docs/commands.json, docs/COMMANDS.md.
- Schema `replyAll: z.enum(['true','false']).optional()` (string-valued flag, precedent `inlineImages`). `const action = replyAll === 'false' ? 'createReply' : 'createReplyAll'`.
- Meta: new option (magicValue hint true/false); rewrite the "Reply-all by design" summary to "reply-all by default; pass --reply-all false to reply to the sender only"; graphPathTemplate must contain the literal `{reply-all}` placeholder (meta.test invariant).
- Tests: `'false'` → POST `/createReply`; `'true'` and omitted → `/createReplyAll` (kills `!== 'false'` ↔ `=== 'true'` mutants); `'maybe'` → validation_error; comment payload unchanged both paths.

### C4 `feat(mail): --body-content-type HTML splice on create-reply-draft` (request 1a)
Files (4): same as C3.
- Schema `bodyContentType: z.enum(['Text','HTML']).optional()`. Text/omitted path byte-identical to today (single POST + optional body-free subject PATCH).
- HTML branch: (1) reject `findQuoteBoundary(bodyContent) !== -1` before any Graph call; (2) POST `{comment: ''}`; (3) use `created.value.body`, fallback GET `/me/messages/{id}?$select=body` if absent; (4) draft body contentType `text` → validation_error **naming the created draft id** + remedy (draft exists by then); (5) `insertCommentAboveQuote`; (6) single PATCH `{ body: {contentType:'HTML', content: merged}, subject? }`.
- Meta: option added; bodyTemplate gains `{body-content-type}`; summary sentence for the HTML mode.
- Tests: HTML happy path with exact merged string + `patches.length === 1`; no GET when POST response carries body; POST-without-body → GET path pinned incl. `$select=body`; GET failure → passthrough, no PATCH; text-draft + HTML → error containing draft id, no PATCH; no boundary → comment right after `<body dir="ltr">`, everything retained; comment containing `gmail_quote` → error with `posts.length === 0`; explicit `'Text'` → wire-identical to omitted.

### C5 `feat(mail): --body-content-type HTML splice on create-forward-draft` (request 1b)
Files (4): `src/use-cases/commands/create-forward-draft.ts`, its test, docs pair.
- Mirror of C4 via `createForward` (+ `toRecipients` in POST); the single PATCH merges `body` + `ccRecipients?` + `subject?`. Tests mirror C4 plus cc+subject+HTML in one PATCH, and Text path with cc still does today's body-free PATCH.

### C6 `fix(mail): defined clear-list semantics on update-mail-draft` (request 3b)
Files (4): `src/use-cases/commands/update-mail-draft.ts`, its test, docs pair.
- The three recipient gates and `importance` → `!== undefined`; the at-least-one guard → `.every((v) => v === undefined)` over the six fields; `parseRecipients('')` already returns `[]`, which is Graph's clear payload.
- Option descriptions + guard message: "pass an empty string to clear the list".
- Tests: `ccRecipients:''` → PATCH `{ccRecipients: []}` and satisfies the guard alone; same for to/bcc; guard still fires with only messageId, new message asserted verbatim; existing all-fields test untouched.

### C7 `feat(mail): update-mail-draft --comment revises the reply text above the quote` (request 4)
Files (4): same as C6. If docs push past 300 lines, move the docs pair into C10.
- Schema `comment: z.string().min(1).optional()`; mutually exclusive with `bodyContent` (validation_error); joins the at-least-one guard.
- Comment path: GET `/me/messages/{id}?$select=body,isDraft` → require `isDraft === true` → matrix: html draft + Text comment → `replaceCommentAboveQuote(html, escapeTextAsHtml(comment))`; html + HTML → boundary-injection guard then raw replace; text + HTML → validation_error; text + Text → `replacePlainTextCommentAboveQuote`. `boundaryFound === false` → error "draft has no quoted history to preserve — use --body-content to replace the whole body". Single PATCH carrying `{ body: { contentType: <draft's own, passed through verbatim>, content } }` merged with any other provided fields.
- Meta: `--comment` option; `--body-content-type` description rewritten (format of --body-content, or of --comment in comment mode; Text default is escaped into HTML drafts, HTML raw, HTML rejected on Text drafts); bodyTemplate gains `{comment}`.
- Tests: all four matrix cells with exact PATCH strings (escaping proven via a `<b>&\n` comment); GET path pinned; `isDraft:false` → error, no PATCH; comment+bodyContent → error, no GET; no-boundary → exact message; comment+subject → one PATCH with both; contentType passthrough `'html'` in → `'html'` out (kills `'HTML'` literal mutant); GET failure passthrough; comment alone satisfies guard.

### C8 `feat(mail): signature-extractor helper` (request 2, part 1)
Files (2): new `src/use-cases/commands/signature-extractor.ts` + test.
- `extractSignatureBlock(html): string | undefined`: `limit = findQuoteBoundary(html)` (or `html.length`); find `/<div[^>]*\bid="Signature"/i` with index < limit (never lifts a stale signature out of quoted history); balanced-div scan (`/<div\b|<\/div>/gi`, depth counter) to the matching close. `undefined` on no marker / marker only in tail / unbalanced markup. String-index based, consistent with the repo (no DOM parser).
- Tests: nested-divs signature extracted exactly; signature only after `divRplyFwdMsg` → undefined; head+tail signatures → head one; unbalanced → undefined; `<div style=… id="Signature">` attribute order; case-insensitive; 3-deep depth fixture (kills ±1 mutants).

### C9 `feat(mail): get-mail-signature command` (request 2, part 2)
Files (5): `src/use-cases/commands/convert-mail-to-markdown.ts` (export existing private names, ~2 lines: `ATTACHMENT_METADATA_SELECT`, `attachmentsListSchema`, `isInlineImage`, `fetchInlineImageBytes`, `formatBytes` + types; cross-command import precedent: read-mail-attachment.ts), new `get-mail-signature.ts`, new test, `src/use-cases/commands/index.ts` (+2), `src/use-cases/commands/graph-scopes.ts` (+1: `['Mail.Read']`, required by graph-scopes.test).
- Flow: `--message-id` given → candidates `[id]`, no scan; else GET `/me/mailFolders/sentitems/messages?$top=10&$orderby=sentDateTime%20desc&$select=id,sentDateTime` (percent-encode the space, hardcoded path). Then **sequential** per-id GET `?$select=body,sentDateTime,hasAttachments`, stop at first `extractSignatureBlock` hit (1 body fetch in the common case). No hit → actionable error: "no OWA signature block (`<div id=\"Signature\">`) found in the last N sent messages — Outlook-desktop-composed mail doesn't carry it; pass --message-id to pin a message sent from OWA".
- On hit: if `hasAttachments`, list attachments with `ATTACHMENT_METADATA_SELECT`, keep `isInlineImage` candidates whose `cid:${contentId}` occurs in the block, `fetchInlineImageBytes` each (2MB cap), `embedInlineImages`. Do **not** run `replaceUnresolvedCidImages` (a text placeholder inside a signature destined for drafts destroys information): oversize/failed fetches keep the cid: ref and are named in `note`.
- Envelope `{ contentType:'text/html', size: utf8 byte length, text: block, sourceMessageId, sentDateTime?, inlinedImages: n, note? }`; meta `producesBytes: true` (so `--output-path` works via the persistIfRequested text branch, extra fields tolerated per the convert-command `note` precedent); category mail, GET, optional `message-id` with `--id` alias; **no `mutates`** (meta.test pinned set stays the 4 draft commands, untouched).
- Tests: scan URL pinned verbatim; stops after first hit (call-order array proves message 2's body never fetched); tail-only-signature message skipped, next scanned; `--message-id` pin → zero scan calls; pinned message without signature → error naming the limitation; empty sentitems → error; referenced small cid → data-URI embedded + `inlinedImages:1`; oversize cid → skipped + note with formatBytes size; unreferenced attachment → no bytes fetch; attachments-list failure → signature still returned with note; `hasAttachments:false` → no attachments call; multi-byte-char size fixture; malformed attachments shape → note.

### C10 `docs(mail): regenerate manifest for get-mail-signature; bump command counts`
Files (3): docs/commands.json (docs:gen), docs/COMMANDS.md (headline auto-bumps to 183; hand-fix the stale "(184)" / "179-vs-184" sentence on line 3 against live `ask-marcel-office help-json | jq '.commands | length'`), README.md (three hand-written spots: line 3 `182 commands` → 183; line 28 `174 GET endpoints` → 175 and `= 182` → 183; line 194 `All 182 commands` → 183).

CHANGELOG + version bump stay out of scope: the repo writes the release section at publish time, and 37 commits are already pending release; this work folds into that.

## Verification

Per commit: `bun test`, `bun run lint:strict`, `bun run typecheck`, `bun run coverage` (100% all tiers), `bun run mutate:staged` (≥90% on staged use-case files); the 8-gate pre-commit hook enforces the same.

Live smoke (own mailbox, drafts only, nothing sends; delete drafts from OWA after):
1. Threaded message → `create-reply-draft --body-content '<p><b>bold</b> &amp; entity</p>' --body-content-type HTML` → `get-mail-message --id <draft> --select body`: comment sits above `appendonsend`/`divRplyFwdMsg`, quoted From/Sent block + `<head><style>` intact; eyeball in OWA and Outlook desktop. This also confirms Exchange SafeHTML preserves the boundary ids (the plan's main external risk).
2. `update-mail-draft --id <draft> --comment 'revised once'` then `--comment 'revised twice'` → GET after each: exactly one comment, quote intact both times (idempotency).
3. `create-reply-draft --reply-all false` → draft toRecipients = sender only; default → full set.
4. Draft with cc set: `update-mail-draft --cc-recipients ""` → GET → `ccRecipients: []`.
5. `get-mail-signature` bare and with `--message-id`; `--output-path` a sig.html and open it (inlined data-URIs render); a desktop-Outlook-only sent history should produce the actionable error.
6. Forward variant of step 1 via `create-forward-draft --body-content-type HTML`.
7. `bun run docs:gen` twice → second run byte-identical; `help-json` count = 183.

## Risks

- Exchange SafeHTML rewriting PATCHed HTML: mitigated by boundary redundancy (three marker families) + live step 1/2 round-trip.
- `id="Signature"` is OWA/new-Outlook only: documented in F2's summary and its no-hit error.
- Draft exists when the F1 HTML path errors mid-flight (text-mode thread, PATCH failure): every such error names the created draft id so the caller can fix it with update-mail-draft or delete it in Outlook.
- meta.test invariants (placeholder ↔ option lock-step, kebab↔camel, mutates pinned set) will fail loudly if any meta edit drifts; that's the safety net, not a risk to work around.
