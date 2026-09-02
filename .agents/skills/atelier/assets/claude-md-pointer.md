This repo follows the atelier coding standard. Consult the `atelier` skill for every
code task here; its hard rules 1-34 bind (TDD with hand-written fakes, `Result` at IO
boundaries, branded types at trust boundaries, and the production disciplines: privacy,
isolation, reliability, observability). Run the `atelier-review-me` skill before landing
changes. Journals: `.claude/LESSONS.md` (append-only memory), `.claude/PLAN.md` (current plan).
If this repo vendors or pins the skill, that pin is a dependency: re-check it on the
dependency-scan cadence and re-sync doctrine and gates together (`references/governance.md`).
