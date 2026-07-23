# Reliability (design for failure)

Everything breaks: disks die, dependencies hang, processes crash mid-request, two people edit the same row. Reliability is not the absence of failure; it is failure made explicit, bounded, and recoverable. This reference is the doctrine behind hard rules 29 (every outbound call has a deadline), 30 (data changes are additive and reversible), and 31 (no lost updates). Errors-as-values is the foundation and lives in `references/result-type.md`; reliability targets and alerting live in `references/observability.md`.

## Every network call has a deadline (hard rule 29)

A call with no timeout is a thread you may never get back; one hung dependency becomes a full outage as callers pile up behind it.

- **Deadline on every outbound call**, set in the infra adapter: `AbortSignal.timeout(ms)` on `fetch`, the client's timeout option on an SDK. No default-patience clients.
- **Retries are bounded and jittered**, and they branch on the error kind: retry `rate-limited` and `io`, never `unauthorized`. Use `retryOnErr` (`references/result-type.md`, Retry) with jittered backoff; a blind retry loop hands a struggling dependency triple traffic at its worst moment.
- **A retried operation that is not naturally idempotent carries an idempotency key**, so the retry can never double-charge or double-send. The provider (or your outbox consumer, below) dedupes on it.
- **Circuit breakers are bought, not bundled.** The timeout is always cheap; add a breaker only for a dependency that has actually earned one. Same instinct as deferring the build while keeping the seam (`references/complexity.md`).

```ts
// src/infra/payments-http.ts: deadline + idempotency key at the adapter
const call = (): Promise<Response> =>
  fetch(url, {
    method: 'POST',
    body,
    headers: { 'idempotency-key': key }, // provider dedupes; the retry is safe
    signal: AbortSignal.timeout(2_000), // fail fast, free the caller
  });
const res = await retryOnErr(() => toResult(call()), (e) => e.kind === 'io' || e.kind === 'rate-limited', { maxAttempts: 3, baseDelayMs: 200, jitter: true });
```

Java: `HttpClient.newBuilder().connectTimeout(...)` plus a per-request `.timeout(...)`, `Idempotency-Key` header, `@Retry(maxRetries = 3, jitter = 200)` on the adapter (`references/java-quarkus.md`).

## Reads are explicit; writes go through the mapper

An ORM or query builder is welcome on the write path, where its mapping and safety are what you want. On the **hot read path**, keep the query explicit and tunable, hand-written SQL or a typed query builder that emits visible SQL, so you can see it, EXPLAIN it, and tune it. A lazy relation walk fires a cascade of hidden queries (the N+1) that is invisible until it is slow.

```ts
// BAD: relation walk; one hidden SELECT per row, nothing to EXPLAIN
const orgs = await db.query.orgs.findMany({ with: { receipts: { with: { lines: true } } } });

// GOOD: one explicit read you can tune; the write stays with the builder
const rows = await db.execute(sql`
  SELECT r.id, r.total_cents, count(l.id) AS line_count
  FROM receipts r LEFT JOIN receipt_lines l ON l.receipt_id = r.id
  WHERE r.org_id = ${orgId}
  GROUP BY r.id ORDER BY r.created_at DESC LIMIT 50`);
await db.insert(receipts).values(newReceipt);
```

The repository port hides which style serves which side; callers see only the contract.

## Keep reads fast as the table grows

- **Page with a keyset cursor, never OFFSET.** OFFSET re-reads from row zero on every page: page 1000 does 1000x the work. A keyset page on a unique compound cursor `(createdAt, id)` costs the same at any depth, and ties neither skip nor duplicate. Back it with a composite index on the cursor columns; without the index, keyset degrades to a scan.
- **Stream or bulk-load large result sets** instead of holding them in memory: cursor through the driver for exports, one row at a time, constant memory.

```ts
export const listAfter = (cursor?: Cursor): Promise<Receipt[]> =>
  db.select().from(receipts)
    .where(cursor ? afterCursor(cursor) : sql`true`) // (createdAt, id) > (cursor.createdAt, cursor.id)
    .orderBy(asc(receipts.createdAt), asc(receipts.id))
    .limit(50); // the last row's (createdAt, id) is the next cursor
```

UUIDv7 primary keys (`references/isolation.md`, Identifiers) keep the same index locality on id-ordered scans.

## Do not fire and forget (the outbox)

When a committed change must trigger a side effect (an email, an event, a downstream update), a best-effort call inside the request vanishes if the process crashes between commit and send, and moving the call inside the transaction is no better: a rollback still mails.

- Record the intent **in the same transaction** as the state change.
- A separate worker delivers it with retries and marks it done.
- Delivery is **idempotent on a dedupe key**, because retries guarantee duplicates: a resent email is noise, a resent payment is an incident.

```ts
await db.transaction(async (tx) => {
  await tx.insert(receipts).values(r);
  await tx.insert(outbox).values({ topic: 'receipt_ready', payload: r.id, dedupeKey: `receipt_ready:${r.id}` });
});
// a poller reads unsent rows, sends through the port, marks sent, retries on Err
```

The worker is a use-case; the sender is a port with a fake; the poller loop is composition. Java: persist an `OutboxEntry` in the same `@Transactional` method, deliver from a `@Scheduled` worker.

## No lost updates (hard rule 31)

When two actors can edit the same record, last-write-wins silently destroys one of them. Carry a version on every read, require it back on every write, reject stale writes with a conflict the client resolves.

```ts
const updated = await db.update(orders)
  .set({ notes, version: sql`${orders.version} + 1` })
  .where(and(eq(orders.id, id), eq(orders.version, expectedVersion)))
  .returning({ id: orders.id });
if (updated.length === 0) return err({ kind: 'stale_write' } as const); // HTTP: 409 + current state
```

A write that cannot prove what it read is a write you refuse. Over HTTP the same idea is `ETag` + `If-Match`. Java: JPA `@Version`; an `OptimisticLockException` maps to 409 with the current state so the client reloads and merges.

## Data lifecycle (hard rule 30)

**Soft-delete by default.** Stamp `deletedAt` and keep the row; reads exclude it; recovery is a flag flip; a scheduled retention sweep decides real removal later.

```ts
const updated = await db.update(orders)
  .set({ deletedAt: new Date() })
  .where(and(eq(orders.id, id), isNull(orders.deletedAt)))
  .returning({ id: orders.id });
// reads stay honest: .where(isNull(orders.deletedAt))
```

The deliberate exception is subject erasure under privacy law (`references/privacy.md`): personal fields are hard-deleted or anonymized. The retention sweep is where both rules run on schedule.

**Every schema change is a versioned migration.** Never a hand-run ALTER; the migration is code, reviewed and committed (Drizzle numbered migrations; Flyway `V8__*.sql` in Java).

**Change shipped contracts additively (expand-contract).** Never rename or drop a column or response field a live client still reads:

1. **Expand**: add the new nullable column or field; backfill.
2. **Migrate**: ship code that dual-writes and reads the new shape; verify.
3. **Contract**: drop the old one in a later release, when nothing reads it.

The same discipline applies to API responses: version or add fields; a client you cannot see must never wake up to a field that vanished. A destructive one-shot rename is a red flag in review.

## Stateless by default; cache deliberately

- No per-user state pinned in process memory (a module-level `Map` of sessions dies on restart and breaks the second replica). State lives in the store or a shared cache; any replica can serve any request; scaling is adding copies.
- Cache with an explicit TTL and a stated invalidation rule, where measurement says it counts. A cache with no invalidation story is a stale-data bug on a timer.
- Components should cost near nothing when idle (`references/metrics.md`, Cost is a first-class metric).

## Performance is a budget, not a hope

Commit to response-time numbers (p95/p99 per route) and prove them under production-like load before shipping: a load-test gate in the pipeline (k6 or similar) with a threshold that fails the build, e.g. `p(99) < 300ms` at expected peak. Finding the p99 from angry users is the Don't. Targets live with the other thresholds (`references/governance.md`, Numbers not adjectives); alerting on them is `references/observability.md`.

## Money, time, and the types that carry proof

Parse, don't validate (rule 12): validate once at the boundary, then carry the fact in a type. Two primitives are dangerous enough to call out (full pattern in `references/clean-code.md`, Wrap all primitives):

- **Money is integer minor units** (`cents`) behind a branded type or value record, never a float: `0.1 + 0.2 !== 0.3`, and the rounding error lands on an invoice.
- **Instants are UTC** behind a type; a timezone is a display concern applied at the presentation edge, never stored in the domain value.

## Separate the analytical store from the operational one (10.14)

The transactional database serves the running application only. Anything that reads at volume (reporting, dashboards, bulk export, data science) reads a **separate analytical copy** (a warehouse or lake) fed by ETL or change-data-capture, never the production store directly. A heavy scan on the primary competes with real users and couples every report to the app's private schema.

```text
[ app + users ] --user-scoped (7.1)--> [ PRIMARY OLTP ]
                                              | CDC / nightly ETL (one direction: out)
                                              v
                                       [ WAREHOUSE / LAKE ]  <-- BI, exports, ML read here
```

Writes stay on the transactional path, done by a user; reads at scale move to the copy, so the app database is tuned for the transaction and the warehouse for the scan, and neither fights the other. The pipeline is the **one sanctioned bulk reader**: its own narrowly granted job at the database layer (`references/isolation.md`, Shrink the blast radius), never the user-facing API (`references/isolation.md`, No service-token backdoor). The analytical copy inherits pillar 6, so subject erasure and the retention sweep propagate to it, and the platform is chosen under the open-interface rule (`references/delivery.md`) or its lock-in taken deliberately behind a port.

## Executable tripwires

The mechanical slices of rules 29 and 30 ship as staged-diff gates (`references/workflow.md`, Discipline tripwires): `assets/check-io-deadlines.sh` blocks an infra file that calls `fetch` (or opens a Java `HttpClient`) with no deadline marker, and `assets/check-data-lifecycle.sh` blocks a hard delete in application code and destructive DDL in a non-contract migration. Exceptions ride on path conventions, never inline suppressions: erasure/retention/prune/sweep paths for the sanctioned hard deletes, a `*contract*` filename for the deliberate contract-step migration. Tripwires, not proofs; the checklist below stays the review duty.

## Review checklist (changes touching IO, persistence, or state)

1. Every new outbound call: deadline set in the adapter? Retry bounded, jittered, and kind-filtered? Idempotency key if retried and not naturally idempotent? (rule 29)
2. New read path: explicit query on the hot path, keyset pagination, streaming for large sets?
3. Committed change with a side effect: outbox row in the same transaction, idempotent delivery?
4. Mutable shared record: version carried and checked, stale write rejected as a conflict? (rule 31)
5. Deletion: soft-delete with filtered reads (or a justified erasure path)? Schema change a versioned migration? Shipped contract changed additively? (rule 30)
6. New state: does it survive a restart and a second replica? Cache TTL + invalidation stated?
7. Hot endpoint: is there a latency budget, and does a load test hold it?
