# Lessons | Memory Across Sessions

Full detail for the Lessons section of `SKILL.md`. Covers when to trigger, how to decide what is a lesson, entry templates, worked examples.

## The two files

```
.claude/
  LESSONS.md           # committed, team-shared, reviewed in PRs
  lessons.local.md     # gitignored, personal to the current developer
```

Both files follow the same entry format. The only difference is scope and visibility. A compaction pass (below) adds an archive beside each, `lessons.archive.md` and `lessons.local.archive.md`: history, never read at session start.

## When to trigger

### Start-of-session signals

Read the lesson files BEFORE any tool call or code action when the user's first message does any of the following:
- References a file, a feature, a bug, or asks for code.
- Says "continue", "picking this up", "working on X today", "back on Y".
- Opens with a specific request: "let's add X", "fix the Y", "refactor Z".
- Precedes any scaffolding, editing, running tests, or installing dependencies.

Skip the start-of-session read only when the first message is clearly not a coding session (pure Q&A about a library, generic architecture chat, a list that needs no repo context). When unsure, read the lessons. The cost of reading two short files is tiny; the cost of repeating a past mistake is not.

### End-of-session signals

Trigger the extraction workflow when the user signals the session is wrapping:
- Explicit: "thanks", "that's all", "ship it", "I'm done", "closing this out".
- Implicit: the user confirms the final change works and goes silent on new requests.
- Time-based: a natural stopping point after a feature is complete and tests pass.
- User-initiated: "anything to add to the lessons?", "save this for next time".

When you are genuinely uncertain whether the session is ending, ask once: "Session wrapping? I will scan for lessons to capture." Do not propose entries without that signal.

### What a "substantive" session looks like

Worth capturing from:
- The user corrected something you wrote.
- You debugged a non-obvious failure.
- An architectural choice was reached with tradeoffs discussed.
- A tool, library, or API behaved unexpectedly.
- A workaround was invented for a real constraint.

Not worth capturing:
- The user asked a factual question and got an answer.
- You wrote boilerplate and it worked first try.
- The session was pure exploration with no decisions.

## The three kinds of entries

Every lesson must fit one of these. If it does not fit, do not capture it.

### 1. `[mistake]`

Something Claude did wrong that the user had to correct. Goal: future Claude will not do this again.

Strong signals:
- User says "no", "not like that", "that's wrong", "we don't do that here".
- User rewrites your output and the rewrite matches a pattern (not just a typo).
- User points to a file or convention you should have checked first.
- You apologised and produced a second version.
- The same kind of error appeared more than once in the session.

Weak signals (do NOT capture):
- The user simply preferred a different variable name.
- You made a typo.
- The user changed their mind about the requirement.
- Your first draft was fine but they wanted a variant.

The test: could a future Claude avoid this mistake by reading a 3-sentence entry? If yes, capture.

### 2. `[decision]`

A choice the user made (or we made together) that shapes future work on this codebase.

Strong signals:
- We weighed two or more options explicitly.
- The user chose one with a stated reason.
- The choice will constrain future work (touches architecture, deployment, data model, naming scheme).
- The user said "let's go with", "let's use", "let's put this in", "always do X".

Weak signals (do NOT capture):
- Picking between two equivalent names.
- Deciding which file to open first.
- Routine sequencing of tasks within a single feature.

The test: would this decision still be relevant in 3 months? If yes, capture.

### 3. `[gotcha]`

A non-obvious fact about the codebase, toolchain, runtime, or environment that cost time to figure out.

Strong signals:
- A tool, library, API, or runtime behaved differently from docs or expectation.
- We spent real time debugging a non-obvious failure and found the cause.
- A config, flag, or environment variable mattered and was not discoverable by reading code.
- A version mismatch, path quirk, permission issue, or proxy setting affected behaviour.
- We found an undocumented or poorly documented behaviour.

Weak signals (do NOT capture):
- A simple syntax error we fixed.
- A test that failed because of a typo.
- User error that will not recur.

The test: if a new developer hit this tomorrow, would this entry save them an hour? If yes, capture.

## Entry format (strict)

```
## [kind] YYYY-MM-DD | short title in lowercase

[2 to 5 sentences: what happened, what the correct answer is, why it matters.]

[Optional one-liner: Rule for next time / Applies to / Affects.]
```

- `kind` is one of `mistake`, `decision`, `gotcha`. Literal, no other values.
- Date is today in ISO format (`YYYY-MM-DD`).
- Title is short, lowercase, descriptive. No marketing language.
- Body is prose, 2-5 sentences. No bullet lists inside an entry. If you need more than 5 sentences, the entry is too big.
- Optional tail: one of `Rule for next time`, `Applies to`, or `Affects`. Pick at most one.

## Routing to the right file

**`LESSONS.md` (committed) when:**
- Every team member would benefit from knowing it.
- It concerns the codebase, architecture, or shared conventions.
- It concerns external services, APIs, or deployment infrastructure the team depends on.
- It is a decision that binds future work.

**`lessons.local.md` (personal, gitignored) when:**
- It is about YOUR workflow, YOUR preferences, YOUR reminders.
- It concerns your local dev setup (paths, aliases, shell, VPN, proxy).
- It is a note you would not want to defend in a PR review.
- It concerns a one-off investigation that is unlikely to matter to anyone else.

When in doubt, personal. The team file has a higher bar.

## Start-of-session workflow

1. Check for `.claude/LESSONS.md`. If it exists, read it in full.
2. Check for `.claude/lessons.local.md`. If it exists, read it in full.
3. If both missing, note internally that this repo has no lessons file yet and proceed normally. Do not create the files pre-emptively.
4. Keep the lessons in mind throughout the session. When a rule applies, follow it silently. Do not narrate "according to LESSONS.md".
5. If a lesson contradicts something the user just asked for, stop and surface the conflict in one sentence. Example: "LESSONS.md says Firebase jobs live in separate repos, but you are asking me to add one to the monorepo. Overriding the lesson, or should we put it in a standalone repo?"
6. If a journal is over its cap (100 entries or ~15 KB), offer the compaction pass once, in one sentence: "`LESSONS.md` is 40 KB against its ~15 KB cap; compact it with atelier-distill?" Never run it unasked.

## End-of-session workflow

1. Scan the conversation from top to bottom.
2. For each moment, ask: mistake, decision, gotcha, or nothing?
3. Draft a candidate list. Maximum 5 entries per session; if you have more, you are over-capturing.
4. For each candidate, decide the target file using the routing rules above.
5. Show the candidate list to the user before writing. One line each: `[kind] target-file | title`.
6. On user approval, append. Never edit or delete past entries; retiring them is the compaction pass's job.
7. If the user rejects an entry, drop it silently. Do not argue.
8. If the user reframes the entry in their own words, use their words verbatim.

## Append-only rules for the files

- Never delete an entry. If superseded, add a new `[decision]` entry that references and overrides the old one. Keep the history.
- Never edit past entries except typo fixes. A lesson was true at the time it was written; that history matters.
- Sort newest-first within each kind section, or maintain a flat reverse-chronological list. Either scheme is fine; pick one and be consistent within a file.
- No filler preamble. The file starts with a one-line description and the first entry. No table of contents, no "how to read this file" section.
- The one exception to the first two rules is the compaction pass below: approved, in its own commits, and it removes nothing it has not archived or that git does not keep.

## Compaction pass

A journal is read in full at every session start, so its size is paid on every session, and a superseded or already-enforced entry is noise the next reader has to discount. Append-only keeps the history honest between passes; the compaction pass keeps the file worth reading. The atelier-distill companion skill runs it.

**Trigger.** `LESSONS.md` or `lessons.local.md` over 100 entries or ~15 KB (offered once at session start, never run unasked), or the user asks to clean up, prune, or reorganize the journals. Age alone is no reason: a two-month-old journal can be several times over the cap, and a year-old gotcha can still bind.

**Verdicts.** Every entry gets exactly one, with its evidence:

| Verdict | When | Evidence | Result |
|---|---|---|---|
| keep | still true, binding, in format | none | untouched |
| tighten | still true, but past 5 sentences or bulleted | the rewrite beside the original | rewritten to 2-5 sentences keeping every command, flag, version, path, and number; the original archived |
| merge | two or more entries carry one lesson | their dates and titles | one entry whose tail reads `Merges:` with the originals' dates; the originals archived |
| graduate | a gate, a reference, or `CLAUDE.md` now states or enforces its rule, not only the one instance it fixed | the enforcing `file:line` | archived with that pointer; the rule lives where it is enforced |
| archive | superseded by a newer entry, or about something that no longer exists | the superseding entry, or the absence | moved verbatim |
| move | filed in the wrong tier (a personal note in the team file, team knowledge in a personal store) | the routing rule it breaks | moved to the right file |
| delete | an exact duplicate of an entry that stays, or noise that fits none of the three kinds | the surviving twin, or the test it fails | removed from a git-tracked file; archived instead in an untracked one |
| promote | a recurring `[gotcha]` or `[mistake]` that nothing enforces | the recurrences | kept, and proposed as a `CLAUDE.md` line or a gate; the pass does not build it |

**The archive.** `.claude/lessons.archive.md` beside the team file (committed) and `.claude/lessons.local.archive.md` beside the personal one (gitignored). Same entry format, newest first, each entry followed by one line: `Archived YYYY-MM-DD: <verdict>, <evidence>.` Nothing reads the archive at session start; grep it when a question needs the why.

**Guarantees.**
1. Report before any write: entry counts and sizes before and after, every verdict with its evidence, the full text of every rewrite and merge.
2. Apply only what the user approves, by group or by number.
3. Archive before removing. A file git does not track (the personal journal, an agent's own memory folder) is backed up before the first write, since git cannot restore it.
4. Afterwards every entry of the old file is accounted for: kept, rewritten, merged into a named entry, archived, moved, or deleted as the twin of a named entry.
5. The pass lands in its own commits, sliced to the commit-size gate, each on the user's yes (rule 25).

## File starters

Create these files only when the first real lesson is captured. Do not create them pre-emptively. The one exception is a repo born from atelier-greenfield, which seeds the team journal with its header so the journal exists from the first commit.

### `.claude/LESSONS.md` starter

```markdown
# Lessons (committed)

Append-only institutional memory for this codebase. See the atelier skill's `references/lessons.md` for the format and rules.

Each entry is one of `[mistake]`, `[decision]`, or `[gotcha]`. Newest first.

---
```

### `.claude/lessons.local.md` starter

```markdown
# Lessons (personal, gitignored)

Append-only personal notebook. See the atelier skill's `references/lessons.md` for the format and rules.

Each entry is one of `[mistake]`, `[decision]`, or `[gotcha]`. Newest first.

---
```

### `.gitignore` addition

When creating `lessons.local.md` for the first time, add these lines to `.gitignore` if not already present (the second covers its archive):

```
.claude/lessons.local.md
.claude/lessons.local.archive.md
```

Do not gitignore the whole `.claude/` folder. The team needs `LESSONS.md` and any skills.

## Sample entries

### Mistake (team-relevant)

```markdown
## [mistake] 2026-04-23 | used npm install instead of bun add

On the pricing feature, I suggested `npm install zod` to add the validation library.
This repo is Bun-only; the correct command is `bun add zod`. User corrected immediately.
The atelier skill lists Bun as the only package manager. I should have checked
the conventions before suggesting any shell command.
Rule for next time: read CLAUDE.md and any skills before suggesting shell commands.
```

### Mistake (personal-relevant)

```markdown
## [mistake] 2026-04-23 | assumed user wanted a class when they said "Service"

When the user asked for a `UserService`, I defaulted to `class UserService {...}`.
User prefers a module of arrow functions even for things named with `-Service` suffix.
The naming does not imply the implementation style in this codebase.
Rule for next time: "Service" means "module", not "class".
```

### Decision (team-relevant)

```markdown
## [decision] 2026-04-23 | Firebase admin jobs live in separate Bun-script repos

Debated folding Firebase admin jobs into the main Next.js monorepo vs keeping them
in standalone Bun-script repos. Decided standalone: different deploy targets, different
secrets surface, different failure modes. The monorepo hosts only the user-facing app
and its supporting packages.
Applies to: any future "should this job go in the monorepo" question.
```

### Decision (personal-relevant)

```markdown
## [decision] 2026-04-23 | I install all dev CLIs through mise

Need a single tool that pins per-project versions of Bun and Node without
sudo. asdf works but config is verbose; brew installs system-wide and drifts.
Mise keeps every project's `.tool-versions` honest and stays out of PATH globals.
Applies to: any new language runtime or CLI I need locally.
```

### Gotcha (team-relevant)

```markdown
## [gotcha] 2026-04-23 | Stryker runs through npm even in a Bun-only repo

Mutation testing failed with 'npm not found' in a container that only had Bun.
Stryker's plugin loader and commandRunner resolve through npm/npx regardless of
the test runner, so stryker.conf.json keeps `packageManager: "npm"` and CI images
need node+npm installed even though tests execute via `bun test`.
Affects: every repo using the mutation gate.
```

### Gotcha (personal-relevant)

```markdown
## [gotcha] 2026-04-23 | bun install fails behind a transparent HTTPS proxy without NODE_EXTRA_CA_CERTS

Fresh machine on a network with TLS interception. `bun install` died with
"unable to verify the first certificate". Bun reads `NODE_EXTRA_CA_CERTS`
the same way Node does: exporting it to the proxy's root CA bundle fixes
both `bun install` and `bunx`.
Affects: every shell session on this network; add to `~/.zshenv`.
```

## Superseding an old entry

When a new lesson contradicts an older one, do not edit the older one. Add a new `[decision]` entry that references and overrides.

The older entry:

```markdown
## [gotcha] 2026-04-23 | bun install rewrites bun.lock on drift, CI needs --frozen-lockfile

A CI run silently updated bun.lock because plain `bun install` reconciles the lockfile
when package.json drifted. CI must run `bun install --frozen-lockfile` so drift fails
the build instead of mutating the lockfile.
Affects: every CI pipeline in Bun repos.
```

The superseding entry, appended later:

```markdown
## [decision] 2026-05-15 | SUPERSEDES 2026-04-23 gotcha on bun install lockfile drift

The `--frozen-lockfile` flag moved from per-pipeline steps into the shared CI install
script, so individual pipelines no longer set it by hand.
Supersedes: "bun install rewrites bun.lock on drift, CI needs --frozen-lockfile" (2026-04-23).
Affects: CI pipelines created after 2026-05-15.
```

Both entries stay in the file. A reader can trace the history.

## What a poorly-written entry looks like

Do not do this:

```markdown
## [mistake] 2026-04-23 | made some errors in the pricing code

We were working on a pricing feature and I made a few mistakes along the way.
User had to correct me a few times on naming and on the structure. We eventually
got it right after some back and forth, and the final version worked well.
The user was patient throughout. Overall a good session with lessons learned
about being more careful with conventions and paying attention to details.
```

Problems: vague, narrative, nothing transferable, reads like a session summary.

Rewrite as specific and concrete:

```markdown
## [mistake] 2026-04-23 | put validation logic inside the entity instead of the use-case

I added input validation inside the `Order` record's transform functions. User moved
it into the `placeOrder` use-case, where orchestration lives. Entities hold invariants,
not validation of external input. Validation belongs at the application boundary.
Rule for next time: validate at the use-case boundary, not inside the domain record.
```

Specific, transferable, 3 sentences.

## Example candidate list

After a 2-hour session on a pricing feature, Claude proposes:

```
Candidates for lessons files:

1. [mistake]  lessons.local.md | wrapped Money in a class, user corrected to branded type
2. [gotcha]   LESSONS.md       | Stryker resolves through npm/npx; keep packageManager "npm" even with bun test
3. [decision] LESSONS.md       | discount tiers live in a dispatch record, not a switch

OK to append all three? (reply: all / none / numbers)
```

User replies `1 3` (skip 2, it is already in CLAUDE.md).

Claude appends entries 1 and 3 to their target files and closes the session.

## Deduplication before writing

Before appending, scan the target file for near-duplicates. If the new entry is:

- **Already there:** skip it silently. The lesson is captured.
- **A stronger variant of an older entry:** do NOT edit the old entry. Add the new one; in its body, reference the older one by date and title.
- **A contradiction of an older entry:** surface it to the user. Sometimes a past lesson no longer applies (toolchain changed, convention shifted). Capture the new entry with explicit reference to what it supersedes.

## Common pitfalls

- **Over-capturing.** If every session produces 5+ entries, the bar is too low. Aim for 0-3 per substantive session; zero is fine.
- **Generic advice.** "Always test first" is not a lesson. It is a platitude. Capture the specific, local, concrete.
- **Session narrative.** "First we did X, then Y" is not a lesson. It is a log. Strip the narrative, keep the transferable insight.
- **Preserving the wrong thing.** Capturing what Claude did right is not useful; that is just successful work. Capture what would have gone wrong without correction.

## Overriding the triggers

If the user explicitly says "no lessons" or "skip the journal", comply silently. Do not argue.

If the user explicitly says "always capture X kind of thing", treat that as a standing instruction for the current project. Write it into the top of `.claude/lessons.local.md` as a meta-rule.

## Relationship to other atelier rules

The lessons files capture what is specific to THIS codebase, THIS team, THIS deployment. They do not re-state universal engineering principles from atelier's other references.

When a lesson directly contradicts a general atelier default (a tunable threshold, a preferred tool, a reference's advice), the lesson wins for this repo only; it never overrides hard rules 1-37 or the two confirmation gates, and a lesson that would is surfaced to the user instead of applied. State the contradiction once in the entry, so future Claude understands why the local rule overrides the global one.

## Harvesting lessons as an audit source

The lessons files are not only a memory for the next session. Across many repos they are the best available audit signal for the atelier standard itself, because they record the one thing a static review of the skill cannot produce: real friction, captured at the moment it cost someone time. A design review tells you whether the standard is internally coherent. Dogfooding tells you whether it executes. The accumulated lessons tell you whether it survives contact with real work.

Each kind of entry maps to a distinct audit finding:

- A recurring `[mistake]` across repos means a rule is not landing. Future Claude keeps doing the thing the standard forbids, which points to a rule that is buried, under-explained, or not loud enough in the red flags. The fix is usually in the skill, not the developer.
- A `[gotcha]` is undocumented friction and a direct candidate for a doc fix or a shipped guard. If the same toolchain surprise shows up in three repos' journals, it belongs in a reference file or an asset, not rediscovered each time.
- A `[decision]` that keeps getting re-litigated in different repos means the standard is missing a default it should state. If every team debates the same fork, atelier should pick a side, or name it explicitly as a real fork.

### The harvest workflow

1. Collect `.claude/LESSONS.md` and `.claude/lessons.archive.md` from the repos that use atelier (the committed files, not the personal ones). The archive's graduated entries show which lessons already had to become a gate or a doc line.
2. Cluster entries by theme, ignoring the repo-specific specifics. Five journals each carrying a `[gotcha]` about the same flag is one audit item, not five.
3. Rank clusters by recurrence and cost. A gotcha in four repos outranks a one-off, however painful the one-off was.
4. Each surviving cluster is an audit item: a doc gap to close, a rule to sharpen, a guard to ship, or a default to declare. Route it the way any conformance finding is routed (atelier-review-me owns that lens).

### The caveat

This is a sample of pain that someone bothered to write down, in repos that keep the journal. It is biased toward friction that was noticed and toward disciplined teams. It tells you what hurt, not what is silently wrong-but-tolerated, and it says nothing about the repos that never adopted the journal at all. So treat it as one leg of a tripod: static audit catches internal incoherence, dogfooding catches "does it execute", the lessons harvest catches "does it hold up in real use". No single leg replaces the others.

The practical consequence: when dogfooding atelier on a real repo, keep its `.claude/LESSONS.md` rather than deleting the repo the moment the gates go green. A handful of retained journals becomes a standing audit backlog you can sweep periodically, a far stronger input than any single review pass.
