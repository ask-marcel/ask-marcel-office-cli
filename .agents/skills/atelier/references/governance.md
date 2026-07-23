# Governance (no black boxes, clear ownership)

Anyone with a stake in the project should be able to see its real state at any moment, and every part of the system should have a name next to it. Opacity and orphaned ownership are how projects quietly rot. The README discipline (Behavioural Guideline #5), the durable plan, and the lessons journal already cover the essentials of legibility; this reference adds decision records, contract-generated API docs, measurable commitments, the honest backlog, and the ownership machinery.

Two ground rules first: the project has **one working language** (docs, comments, commit messages, identifiers), chosen once and kept everywhere, because a mixed-language repo taxes every reader; and **documentation drift is a defect**: if the README no longer matches how the project installs, runs, or deploys, the change is not finished even with green tests (Behavioural Guideline #5; `references/workflow.md`, README consistency).

## README stays runnable (docs-check in CI, 12.1)

Drift-as-a-defect is a review duty until a gate makes it mechanical. Keep a README that actually installs and runs the project (exact, runnable commands, not prose), and **fail CI when it goes stale**: the shipped `assets/check-docs.sh` runs the fenced ```bash block under the README's `## Verify` heading, so a documented command that no longer works fails the pull request. Keep the Verify block self-contained and fast (a health curl, a smoke command, a path assertion), not the full install.

```yaml
# .github/workflows/docs-check.yml
name: docs-check
on: [pull_request]
jobs:
  readme-runs:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: oven-sh/setup-bun@v2
      - run: bash scripts/check-docs.sh # runs the README's ## Verify commands
```

The atelier repo's own `scripts/smoke-test.sh` is the reference implementation: it follows this README's install steps verbatim into a scratch repo and fails if any of them break, which is exactly a docs-check for a project whose product is its instructions.

## Decision records (why is it like this)

Two tiers, one rule: the record changes in the same commit as the code it explains, so it cannot drift.

- **Every significant decision** gets a one-line `[decision]` entry in `.claude/LESSONS.md` (append-only; superseded by a newer entry when it changes). This is the index and stays the default (`references/lessons.md`).
- **Decisions with rejected alternatives and a reversal path worth keeping** (a vendor, a storage engine, a deliberate lock-in, a security tradeoff) additionally get a full decision record: `docs/adr/NNNN-title.md`, committed with the change. The atelier-grill-me interview output is the natural draft.

```markdown
# 0007: Encrypt state client-side
- Status: accepted   - Date: 2026-07-06
## Context
<the forces, in two or three lines>
## Decision
<what was chosen>
## Options considered
- <option>. Rejected: <why, one line>
## Consequences
- <cost accepted>
- Reversal: <the concrete steps that undo this>
```

The test of a good record: a maintainer who was not in the room can answer "why is it like this" and "how would we undo it" without asking anyone.

## API documentation is generated from the contract

If the project exposes an API, its reference documentation derives from the same schema that validates requests, with a real example per endpoint, published where consumers can reach it. A hand-maintained API wiki silently disagrees with the running code within a month; an API without documentation someone could onboard against is a private API you happen to have left exposed.

```ts
// one Zod schema validates the request AND emits the OpenAPI, examples included
const Invoice = z.object({
  amountCents: z.number().int().positive().openapi({ example: 4200 }),
  currency: z.literal('EUR').openapi({ example: 'EUR' }),
}).openapi('Invoice');
export const createInvoiceRoute = createRoute({
  method: 'post', path: '/v1/invoices',
  request: { body: { content: { 'application/json': { schema: Invoice } } } },
  responses: { 201: { description: 'Created' } },
});
```

Java: MicroProfile OpenAPI annotations on the resource render the spec from the code itself (`references/java-quarkus.md`). Either way the wire contract is also where DTO shapes stop: the internal model is mapped at the boundary (`references/architecture.md`, The internal model is yours).

## Numbers, not adjectives

"Fast", "secure", and "well tested" mean nothing until they are numbers someone agreed to and anyone can check. Commit a thresholds file (SLOs, latency budgets, error rates: the same one `references/observability.md` alerts against) and give stakeholders live access from day one: the repository, the pipeline, the board, the dashboards. Not a monthly summary; the actual thing.

## One honest backlog

- One shared, visible tracker is the source of truth: if it is not on the board, it is not work. No shadow spreadsheet, no "can you also..." in DMs that never lands.
- Bugs are first-class issues, not a hand-maintained "known issues" page.
- Deliberately deferred work is visible with its why, not silently absent.
- Status is honest: "blocked, waiting on X" beats an "in progress" that has not moved in a week.

## Ownership is explicit (a name next to everything)

Shared ownership with no name attached is how things rot: everyone assumes someone else has it.

```bash
# .github/CODEOWNERS: every path maps to an accountable owner; last match wins
*                 @org/platform        # default owner, nothing is orphaned
/packages/core/   @org/domain-team
/docs/adr/        @org/architecture
```

Back it with a short RACI note (`docs/OWNERSHIP.md`): for each area, exactly **one** Accountable, any number of Responsible, who is Consulted and Informed. If two teams claim Accountable, split the area. For anything that can break, you should be able to name its owner in seconds.

## Separation of duties

Whoever requests a sensitive change is never its sole approver:

- Required independent review on `main`; the author cannot approve their own change; new commits dismiss stale approvals; `enforce_admins` so nobody is exempt; no direct pushes.
- Production access follows least privilege and passes through review; infrastructure write access belongs to the pipeline alone (`references/delivery.md`, Humans are read-only).
- Branch protection is code, not a console setting.

This coexists with trunk-based development (`references/workflow.md`): small same-day branches through a required review are still trunk-based; weeks-long divergence is not.

## Audit trail

Sensitive mutations record who, what, and why, durably, in the same transaction as the change, so accountability is real rather than nominal. Approvals, exceptions, and emergency access leave a record.

```ts
export const upgradePlan = (deps: Deps) =>
  async (ctx: { actorId: ActorId; reason: string }, orgId: OrgId): Promise<Result<void, PlanError>> =>
    deps.db.transaction(async (tx) => {
      await tx.update(orgs).set({ plan: 'enterprise' }).where(eq(orgs.id, orgId));
      await tx.insert(auditLog).values({ actorId: ctx.actorId, action: 'plan.upgrade', target: orgId, reason: ctx.reason, at: new Date() });
    });
```

## Finding problems is safe; "done" is verifiable by its owner

- If whoever spots an issue is saddled with owning it, people stop looking. Reward detection; route the fix deliberately (a spotted issue becomes a first-class backlog item, not the finder's homework).
- Whoever is accountable must be able to verify completion themselves: "done" is a re-runnable check (a test, a script, a command with an exit code) the owner can execute, not a screenshot or a status-meeting claim. Evidence anyone can reproduce is the standard (`references/workflow.md`, Verification discipline).

## The platform is a product

Whatever paved road the team ships (templates, gate assets, scaffolds) is run like a product: an owner in CODEOWNERS, a changelog, a support channel, a deprecation policy, and a feedback loop. A golden path nobody maintains gets quietly forked around.

## Review checklist

1. Significant choice in this change: `[decision]` entry, and an ADR if alternatives and a reversal path are worth keeping?
2. API surface changed: does the published spec regenerate from the schema, examples included?
3. Any commitment stated as an adjective that should be a number in the thresholds file?
4. New area of code or infra: does CODEOWNERS map it to exactly one Accountable owner?
5. Sensitive mutation: audit row in the same transaction, actor and reason included?
6. Does anything here bypass required review or widen production access? That is a separation-of-duties change; treat it as sensitive.
