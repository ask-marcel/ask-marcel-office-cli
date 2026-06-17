---
name: atelier-grill-me
description: Relentlessly interview the user about a plan, design, or architectural decision — one question at a time, each with your recommended answer, exploring the codebase before asking — until every branch of the decision tree is resolved and you reach shared understanding. Use when the user wants to stress-test or pressure-test a plan, de-risk a big decision before building, walk through a design's tradeoffs, or says "grill me".
---

# Grill me

Pressure-test a plan or design by interviewing the user until there is genuinely shared understanding: every consequential branch of the decision tree resolved, and the dependencies between those decisions settled in order. The point is not interrogation for its own sake — it is to surface the assumptions, edge cases, and forks now, at conversation cost, instead of mid-build or in production.

This is the on-demand, relentless counterpart to "think before coding." That habit asks for clarification when warranted; this skill commits to walking the *entire* tree before a line is written.

## When to use

- The user asks to be grilled, or to stress-test / pressure-test a plan or design.
- A decision carries real stakes and a bad call is expensive to reverse: architecture, data model, public API shape, a migration, a dependency choice, a security boundary.
- A plan is vague, broad, or hides many unstated branches.

Match intensity to stakes. A five-question interrogation of "rename this variable" is noise — skip the grilling for trivial or already-well-specified tasks.

## How to run the interview

1. **Frame neutrally.** Restate the goal in a sentence or two, without tilting it toward a preferred answer, and confirm you have it right before drilling in.
2. **Explore before asking.** Anything the codebase can answer, answer yourself — read the files, the config, the existing patterns. Never ask the user what a `grep` would tell you. Bring findings, not homework.
3. **One question at a time.** Ask a single question, wait for the answer, then ask the next. A numbered list of ten questions overwhelms, and the answers come back shallow.
4. **Lead with your recommended answer.** For each question, say what you'd choose and why in one line, so the user can reply "yes" or correct you. A question with no recommendation just hands the work back to them.
5. **Walk the tree depth-first, resolving dependencies in order.** When one decision constrains others, settle it first and follow its branch; don't open a new branch until the current one is closed. Name the dependency when it matters ("if we keep static export, the per-request-caching question disappears").
6. **Track state.** Keep a running tally of decisions resolved and questions still open, so neither of you loses the thread on a long tree.
7. **Stop at shared understanding.** You are done when every consequential branch is resolved and you could implement without guessing. Say so explicitly — don't trail off mid-interview.

## Output

When the interview converges, write a short **decision record**: the goal, each decision with its one-line rationale, and the first concrete next step. Keep it tight — it is the spec the implementation will follow.

Grill toward the *simplest* design that survives the questions, not the most elaborate one — every answer you recommend should still respect YAGNI and "the cheapest code is the code you never wrote."

In a repo that keeps an `.claude/LESSONS.md` journal, propose the durable choices as `[decision]` entries (append on approval) so the next session inherits them rather than re-litigating the tree.
