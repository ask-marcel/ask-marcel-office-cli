---
name: qa-audit
description: Run the full-surface QA health check of the ask-marcel CLI by executing docs/QA-PLAYBOOK.md phase by phase — offline gates, bundle smoke tests, manifest-driven parameter matrix, conversion contract matrix, live read-only Graph drift probes, agent-ergonomics journeys, and gap/roadmap analysis. Use when the user asks to audit the CLI, run QA, check CLI health, find bugs/regressions across the command surface, or re-verify after a Microsoft Graph change. Produces a severity-ranked report and a proposed fix plan; never fixes code during the run.
---

# QA audit runner

You are auditing the `ask-marcel` CLI. The authoritative procedure is **`docs/QA-PLAYBOOK.md`** — read it in full before doing anything, then execute phases A→H in order. This skill only adds the operating rules:

1. **Track phases as tasks.** One task per playbook phase (A–H); mark in_progress/completed as you go. Full surface every run — no sampling, no tiers.
2. **Report, then plan.** You find and document; you do NOT fix during the run (not even trivial P1s — note them, finish the sweep). After delivering the report, propose a severity-ordered fix plan and WAIT for approval.
3. **Read-only tenant access.** Only GET/search commands. Respect every safety rule in playbook §0 (no external URL fetches, single-quote `b!` ids, `--output-path` for multi-MB bytes).
4. **Raw findings are private.** Write the full report to `.claude/qa-reports/YYYY-MM-DD.md` (gitignored — tenant file names/subjects/IDs may appear there). The ONLY committed artifacts of a run are: the sanitized Run-log row appended to playbook §I, plus any updates to the §E2 drift register and §G roadmap register.
5. **Compare against the previous run.** Read the most recent `.claude/qa-reports/*.md` first; anything that passed before and fails now is a REGRESSION and outranks same-severity new findings.
6. **Manifest-driven, not hand-typed.** Regenerate worklists from `docs/commands.json` (snippets are in the playbook) so the audit scales as commands are added.
7. **Keep the playbook honest.** If a check is wrong/stale (a contract changed intentionally, a fragile probe is obsolete), update `docs/QA-PLAYBOOK.md` in the same run — the playbook is a living document, but contract-matrix changes need the maintainer's explicit OK.
8. **Budget**: a full run is several hours of focused work. If context runs low mid-run, write partial findings to the report file FIRST, then continue — never lose evidence.
