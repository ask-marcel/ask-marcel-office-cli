# Delivery and operations (boring is the compliment)

Shipping is a routine, not an event. This reference covers the path to production and what runs there: the automated pipeline, infrastructure as code, environments, running as little as possible yourself, release artifacts, recovery drills, and postmortems. Trunk-based development, the small-commit gate, and CI CVE scanning already live in `references/workflow.md`; whether delivery is improving is measured in `references/metrics.md`; this file is everything from the merge onward.

## The pipeline is the only deployer

- Build, test, and deploy through an automated pipeline; eliminate manual deployment entirely. A flaky pipeline is itself a bug: expect well above nine builds in ten green.
- Deploy small and often; each change is easy to understand and to roll back in one step (seconds, not a maintenance window).
- Roll out progressively: a canary slice first, promote on healthy metrics, so a bad change reaches a few users rather than everyone.

```yaml
# deploy.yml (sketch): canary, watch the SLO, promote; rollback is one re-run of the last good image
- run: ./scripts/deploy --canary 10 --image "$IMAGE"
- run: ./ci/watch-error-rate.sh --max 0.01 --window 5m   # aborts the job if breached
- run: ./scripts/deploy --promote 100 --image "$IMAGE"
```

Emit a machine-readable deployment event from the same job, so delivery statistics are computed, never remembered: that event stream is what the DORA metrics are derived from (`references/metrics.md`).

## Infrastructure is code; humans are read-only

- Every resource lives in version-controlled files (OpenTofu/Terraform or equivalent) and the whole environment rebuilds from scratch with one command in minutes. A resource clicked into existence in a console has no history, no review, and no way back.
- **Only the pipeline holds write access to production infrastructure.** Humans get read-only; a change reaches the cloud by merging reviewed code. That makes every infra change versioned, reviewed, and reversible by construction, and it is the separation-of-duties control (`references/governance.md`) expressed mechanically.
- Branch protection itself is code (required reviews, required checks, `enforce_admins`), not a setting someone toggled once.

## Environments: separate, honest, disposable

- Development, staging, and production are isolated; production credentials never live in code (`references/security.md`).
- Any component runs on its own against fakes (that is what the port discipline buys), so nobody needs the entire system standing to work on one piece.
- Spin up throwaway environments per branch or load test and destroy them on close (an IaC workspace keyed to the PR number), so validation never queues behind one shared, half-broken staging.
- Test data in every non-production environment is synthetic (hard rule 34; `references/privacy.md`).

## Run as little as possible yourself

Every server you patch and certificate you renew is attention leaking away from the product.

- **Managed over self-run.** Managed databases, queues, and object stores; never hand-run something a platform operates more reliably, never a VM you then own.
- **No servers you log into.** Deploy immutable container images and replace them; if the answer to an incident is "SSH in and poke around", the design is wrong. The shipped Dockerfile (`references/bun-typescript.md`, Containerization) is already immutable and non-root; there is nothing to log into by construction.
- **Automatic TLS only.** Build on platforms that issue and renew certificates themselves; an expired-certificate outage is a self-inflicted wound.

## Rent open standards, so any cloud can run it

A managed service is safe to depend on when it is the managed form of an open interface:

| Need | Open interface |
|:---|:---|
| Compute | OCI container |
| Relational data | Postgres wire protocol |
| Cache | Redis protocol |
| Objects | S3-compatible API |
| Identity | OIDC (`references/security.md`, rule 33) |
| Email | SMTP behind a port |
| Telemetry | OpenTelemetry (`references/observability.md`) |

The application reads everything from injected configuration and never imports a cloud SDK outside an adapter (rule 12's config module + the port discipline); the proprietary remainder lives in the IaC layer, which is per-cloud by nature. When a proprietary service is genuinely worth the lock-in, take it deliberately: behind a port, with the exit written down in a decision record (`references/governance.md`).

**The proof is mechanical**: a compose file boots the full stack on generic pinned backends (Postgres, MinIO), and CI runs the smoke suite against it on every merge. If that boots, the app depends on interfaces, not a cloud, and it is the same mechanism that puts a new laptop on the full system before lunch.

```yaml
# compose.yaml: the portability gate
services:
  app:     { build: ., environment: [DATABASE_URL, S3_ENDPOINT, OIDC_ISSUER] }
  db:      { image: "postgres:17.5" }                              # pinned (rule 19's spirit)
  objects: { image: "minio/minio:RELEASE.2025-04-22T22-12-26Z" }   # a bare tag is the mutable-latest anti-pattern
```

## Release artifacts and the supply chain

Gate 3 keeps secrets out of the diff and `bun audit` watches CVEs (`references/workflow.md`); the release side has three more duties:

- **Immutable, versioned artifacts**: build once, address by digest, never deploy a mutable `latest` tag.
- **A bill of materials** recorded per artifact (`syft <image> -o spdx-json > sbom.json`).
- **Signatures** bound to the digest (`cosign sign <image>@sha256:...`), so only trusted, unmodified builds reach production.

Base images are pinned like any dependency and updated on a deliberate cadence, so a pinned version never quietly rots into a known-vulnerable one.

## Backups you have actually restored

An untested backup is a rumour. Schedule a drill (at least quarterly; weekly is cheap in CI) that restores the latest backup into a scratch database, asserts the data is really there, times the restore so you know your recovery time, and drops the scratch. A real incident should be a rehearsal, not a first attempt.

```yaml
on: { schedule: [{ cron: '0 3 * * 1' }] }
steps:
  - run: pg_restore --clean --dbname "$SCRATCH_DB_URL" latest.dump   # timed
  - run: psql "$SCRATCH_DB_URL" -c "SELECT count(*) FROM receipts" | grep -qv ' 0$'
  - run: psql "$ADMIN_URL" -c "DROP DATABASE scratch_restore"
```

## Learn from every failure (blameless postmortems)

After an incident, run a blameless review that ends in owned backlog items, or the same outage is merely scheduled to recur. Commit it; the durable lessons also land as `[gotcha]` entries in `.claude/LESSONS.md` (`references/lessons.md`).

```markdown
# Postmortem: YYYY-MM-DD <title>
- Impact: <who, how long, how bad>
- Timeline: <detect, mitigate, resolve timestamps>
- Root cause: <the gap, not the person>
- Action items: each a tracked ticket with an owner and a due date
```

Fix what actually broke, not just the alert that missed it (`references/observability.md`, Alert hygiene).

## Self-service, and the platform as a product

Provisioning an environment, a pipeline, or an access grant should be a declarative request a team owns (a committed file a controller reconciles), not a ticket queue. Whatever paved road exists (templates, scaffolds, gate assets: for this standard, the atelier skill itself and atelier-greenfield) is a product: owned, versioned, documented, with a feedback loop, so teams follow it because it is easier than not.

## Review checklist (changes touching deploy or infra)

1. Does the change deploy through the pipeline, in a small increment, with a one-step rollback?
2. Any infra change: expressed in code, reviewed, applied by the pipeline (not a console)?
3. Does anything new require SSH, a hand-renewed certificate, or a self-run stateful service? Rework it.
4. New external dependency: managed form of an open interface, or a deliberate, ADR-recorded lock-in behind a port?
5. Release artifact: immutable, digest-addressed, SBOM + signature attached; base image pinned?
6. Is there a restore drill that would notice if backups silently broke?
7. Did the pipeline emit its deployment event, and does the service have a cost line someone watches? (`references/metrics.md`)
