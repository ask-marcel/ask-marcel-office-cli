# CLAUDE.md

Guidance for Claude Code when working in this repository.

## Always use the atelier skill

**Every code task in this repo MUST follow the [atelier](https://github.com/vdelacou/atelier)
skill** — writing, editing, scaffolding, testing, refactoring, reviewing, or debugging. The skill is
installed at `.claude/skills/atelier` (also in `.agents/skills/atelier`) and is the **source of
truth** for conventions: Clean Architecture (`src/{domain,use-cases,infra,presenter,composition}`),
strict TDD (red-green-refactor), `Result<T, E>` at IO boundaries, branded types at trust boundaries,
Bun-only toolchain, and the eight-gate pre-commit hook. Rules are non-negotiable — when in doubt,
consult the skill before acting. If the skill and any doc disagree, the skill wins.

Always use Atelier Skill

Before Commit always use Atelier Review Me

When user Ask Questions or want to change anything, always use Atelier Grill Me


## Commits

Use **conventional commit** messages (e.g. `feat:`, `fix:`, `chore:`, `docs:`).


