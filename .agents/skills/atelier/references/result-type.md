# Error handling with `Result<T, E>`

Expected failures are data. Exceptions are for bugs. Every port that crosses an IO boundary and every use-case returns a `Result<T, E>`. Thrown exceptions are reserved for programmer errors — unreachable code, invariant violations, genuine crashes. The top-level `main.ts` catches them and reports "crashed (unexpected)".

This reference explains the pattern, the discriminated-union error design, the `try/catch` quarantine, the fan-out batch convention, the retry rewrite, and the testing patterns that come with it.

## The `Result` type and helpers

Lives in `src/domain/result.ts`. Zero dependencies. Pure.

```ts
export type Result<T, E> =
  | { readonly ok: true; readonly value: T }
  | { readonly ok: false; readonly error: E };

export const ok = <T>(value: T): Result<T, never> => ({ ok: true, value });
export const err = <E>(error: E): Result<never, E> => ({ ok: false, error });

export const mapResult = <T, U, E>(r: Result<T, E>, f: (t: T) => U): Result<U, E> =>
  r.ok ? ok(f(r.value)) : r;

export const mapError = <T, E, F>(r: Result<T, E>, f: (e: E) => F): Result<T, F> =>
  r.ok ? r : err(f(r.error));

export const andThen = <T, U, E>(r: Result<T, E>, f: (t: T) => Result<U, E>): Result<U, E> =>
  r.ok ? f(r.value) : r;

export const unwrap = <T, E>(r: Result<T, E>): T => {
  if (!r.ok) throw new Error(`unwrap on err: ${JSON.stringify(r.error)}`);
  return r.value;
};
```

`unwrap` is reserved for bootstrapping code and tests where the error branch is genuinely impossible. Production code pattern-matches on `.ok`.

## Per-port discriminated-union errors

Every secondary port declares its own error type as a discriminated union. The `kind` field is exhaustively matchable.

```ts
// src/use-cases/ports/token-refresher.ts
export type TokenRefreshError =
  | { readonly kind: 'expired'; readonly message: string }
  | { readonly kind: 'rate-limited'; readonly retryAfterMs: number; readonly message: string }
  | { readonly kind: 'unauthorized'; readonly message: string }
  | { readonly kind: 'unknown'; readonly message: string };

export type TokenRefresher = {
  readonly refresh: (current: string) => Promise<Result<string, TokenRefreshError>>;
};
```

```ts
// src/use-cases/ports/sheets.ts
export type SheetsError =
  | { readonly kind: 'read-failed'; readonly message: string }
  | { readonly kind: 'write-failed'; readonly message: string }
  | { readonly kind: 'not-found'; readonly message: string };

export type Sheets = {
  readonly readRows: (tab: string) => Promise<Result<ReadonlyArray<SheetRow>, SheetsError>>;
  readonly appendOrUpdate: (tab: string, row: SheetRow) => Promise<Result<void, SheetsError>>;
  readonly deleteRow: (tab: string, id: string) => Promise<Result<void, SheetsError>>;
};
```

**Why discriminated union, not a plain `Error` or a string:** callers can branch on `kind` with exhaustive type checking. A new error variant triggers a TypeScript error in every switch that forgot it. Messages are for humans; kinds are for code.

## Use-case errors aggregate to `StepError`

Use-cases wrap port errors into a shared `StepError` so the top-level pipeline can aggregate uniformly without pattern-matching on every port-specific error.

```ts
// src/use-cases/ports/step-error.ts
export type StepError = {
  readonly step: string;
  readonly cause: string;
  readonly message: string;
};

export type Summary = { readonly published: number; readonly errored: number };

export type PostToChannel = (input: PostInput) => Promise<Result<Summary, StepError>>;
```

Typical use-case shape:

```ts
export const createPostTelegram = (deps: {
  telegram: Telegram;
  sheets: Sheets;
  logger: Logger;
}): PostToChannel => async (input) => {
  const rows = await deps.sheets.readRows('POST');
  if (!rows.ok) return err({ step: 'postTelegram', cause: rows.error.kind, message: rows.error.message });

  let published = 0;
  let errored = 0;
  for (const row of rows.value) {
    const sent = await deps.telegram.send(row.channel, row.message);
    if (!sent.ok) {
      deps.logger.warn('telegram.send.failed', { row: row.id, kind: sent.error.kind });
      errored += 1;
      continue;
    }
    const written = await deps.sheets.appendOrUpdate('POST', { ...row, sent: 'yes' });
    if (!written.ok) {
      deps.logger.warn('sheets.write.failed', { row: row.id, kind: written.error.kind });
      errored += 1;
      continue;
    }
    published += 1;
  }
  return ok({ published, errored });
};
```

## `try/catch` quarantine

After adopting `Result`, the repo has exactly these places that may contain `try/catch`:

| Location | What for |
|:---|:---|
| `src/infra/**` | Every adapter wraps its third-party call (`fetch`, `googleapis`, `Bun.file`, `twitter-api-v2`). The catch translates thrown exceptions into the port's discriminated-union error variants and returns `err(...)`. |
| `src/domain/**` | Only for native synchronous APIs whose normal contract is to throw: `JSON.parse` (with a safe fallback), `URL` constructor (validation gate), `decodeURIComponent`. |
| `src/main.ts` | Exactly one top-level catch. Sends the bug report ("crashed (unexpected)"), calls `process.exit(1)`. |

`*.test.ts` files and `src/test-helpers/**` sit outside the quarantine — test code may catch (mirrors hard rule 20's test carve-out).

`src/use-cases/**` has **zero** `try/catch`. Every port call is pattern-matched on `.ok`. If you catch a thrown exception inside a use-case, the port has lied about its contract — fix the port, don't silence the symptom.

```ts
// BAD - try/catch inside a use-case hides a lying port contract
export const placeOrder = async (input, deps): Promise<Result<Summary, StepError>> => {
  try {
    await deps.orders.save(input);
    return ok({ saved: 1, errored: 0 });
  } catch (e) {
    return err({ step: 'placeOrder', cause: 'unknown', message: formatError(e) });
  }
};

// GOOD - pattern-match on the Result the port returns
export const placeOrder = async (input, deps): Promise<Result<Summary, StepError>> => {
  const saved = await deps.orders.save(input);
  if (!saved.ok) {
    return err({ step: 'placeOrder', cause: saved.error.kind, message: saved.error.message });
  }
  return ok({ saved: 1, errored: 0 });
};
```

**Mapping errors to an HTTP status.** That flatten is also why an inbound HTTP adapter cannot exhaustively switch a use-case failure to `401`/`404`/`429`: `cause` is now a plain `string`, not the literal `kind` union. So decide precise client errors (`400`) upstream at the branded request checkpoint where the type is still narrow, and default a use-case `StepError` to `500`. To honor a typed status from a port failure, compute it at the `.ok` guard — where `error.kind` is still a literal union TypeScript checks for totality — and carry it as plain data through the flatten (the way `retryOnErr` branches on the live `e.kind` before any flatten). See `references/architecture.md` § Inbound HTTP (server archetype).

## Adapter pattern: the `try/catch` boundary

An infra adapter is the only place where a `throw` from a third-party library gets turned into a `Result`. This is the quarantine — everything beyond it is exception-free.

```ts
// src/infra/sheets-google.ts
export const createSheetsGoogle = (client: SheetsClient): Sheets => ({
  readRows: async (tab) => {
    try {
      const response = await client.spreadsheets.values.get({ range: tab });
      return ok(rowsFromResponse(response));
    } catch (e) {
      return err(classifySheetsError(e));
    }
  },
  appendOrUpdate: async (tab, row) => {
    try {
      await client.spreadsheets.values.append({ range: tab, resource: { values: [toRow(row)] } });
      return ok(undefined);
    } catch (e) {
      return err(classifySheetsError(e));
    }
  },
  // ...
});

const classifySheetsError = (e: unknown): SheetsError => {
  const message = formatError(e);
  if (message.includes('not found')) return { kind: 'not-found', message };
  if (message.includes('quota')) return { kind: 'write-failed', message };
  return { kind: 'read-failed', message };
};
```

Every `catch (e)` uses the shared `formatError(err: unknown): string` helper from `src/domain/utilities/format-error.ts` — never `String(e)`, which returns `"[object Object]"` for non-Error throws (SonarJS S6551).

## Fan-out batch semantics: `ok(summary)` with an `errored` count

Use-cases that iterate over many rows (a Telegram batch send, an enrichment pass, a social-network post loop) catch per-row port errors internally, increment an `errored` counter, and return `ok({ published, errored })`. They do **not** return `err(...)` on a single-row failure.

```ts
// BAD - one bad row aborts the whole batch
for (const row of rows.value) {
  const sent = await deps.telegram.send(row.channel, row.message);
  if (!sent.ok) return err({ step: 'postTelegram', cause: sent.error.kind, message: sent.error.message });
  published += 1;
}

// GOOD - bad rows are counted; the batch completes
for (const row of rows.value) {
  const sent = await deps.telegram.send(row.channel, row.message);
  if (!sent.ok) {
    deps.logger.warn('telegram.send.failed', { row: row.id, kind: sent.error.kind });
    errored += 1;
    continue;
  }
  published += 1;
}
return ok({ published, errored });
```

`err(...)` is reserved for use-case-level prerequisites that prevent the batch from starting at all — the initial `sheets.readRows` fails, the API credentials are missing, the pipeline config is malformed. Rows are best-effort. This keeps a run-pipeline that orchestrates multiple channels from short-circuiting when one of them has a single bad record.

## Retry: `retryOnErr`, not `withRetry`

The pre-Result `withRetry(() => port.call())` retried on thrown exceptions. After the migration, ports do not throw — so the retry path is dead code. Replace it with a Result-aware retry that branches on the error `kind`.

```ts
// src/domain/utilities/retry-on-err.ts
export type RetryOpts = { readonly maxAttempts: number; readonly baseDelayMs: number; readonly jitter?: boolean };

export const retryOnErr = async <T, E>(
  fn: () => Promise<Result<T, E>>,
  shouldRetry: (error: E) => boolean,
  opts: RetryOpts
): Promise<Result<T, E>> => {
  let last: Result<T, E> = await fn();
  for (let attempt = 1; attempt < opts.maxAttempts; attempt += 1) {
    if (last.ok || !shouldRetry(last.error)) return last;
    const backoff = opts.baseDelayMs * 2 ** (attempt - 1);
    await sleep(opts.jitter === false ? backoff : backoff * (0.5 + Math.random())); // jittered by default: synchronized retries stampede a struggling dependency
    last = await fn();
  }
  return last;
};

// usage: retry only on rate-limited, never on unauthorized
const refreshed = await retryOnErr(
  () => deps.tokens.refresh(current),
  (e) => e.kind === 'rate-limited',
  { maxAttempts: 3, baseDelayMs: 500 }
);
```

Never use `withRetry` from the pre-Result era on a port call. It will silently never retry.

Retry is one third of hard rule 29; the other two live in the adapter, not here:

- **The deadline.** Every outbound call the adapter makes carries an explicit timeout (`AbortSignal.timeout(ms)` on `fetch`, the SDK's timeout option), translated to an `Err` kind like any other failure. `retryOnErr` bounds the attempts; the deadline bounds each attempt. Without it, one hung dependency parks the process.
- **The idempotency key.** When the operation is not naturally safe to repeat (a payment, a send), the adapter attaches an idempotency key so a retried attempt can never double-execute. See `references/reliability.md` for the full pattern, including the transactional outbox for side effects that must survive a crash.

## Testing Result-returning code

### Fakes with an `errors` knob

In-memory fakes expose a `errors` map that lets tests inject `err(...)` returns for specific operations. Every port fake supports this pattern.

```ts
// src/test-helpers/sheets-fake.ts
export const createSheetsFake = (config?: {
  tabs?: Partial<Record<string, ReadonlyArray<SheetRow>>>;
  errors?: {
    readRows?: SheetsError;
    appendOrUpdate?: SheetsError;
    deleteRow?: SheetsError;
  };
}): Sheets => {
  const store = new Map<string, SheetRow[]>();
  for (const [tab, rows] of Object.entries(config?.tabs ?? {})) store.set(tab, [...(rows ?? [])]);

  return {
    readRows: async (tab) => {
      if (config?.errors?.readRows) return err(config.errors.readRows);
      return ok(store.get(tab) ?? []);
    },
    appendOrUpdate: async (tab, row) => {
      if (config?.errors?.appendOrUpdate) return err(config.errors.appendOrUpdate);
      const rows = store.get(tab) ?? [];
      store.set(tab, [...rows.filter((r) => r.id !== row.id), row]);
      return ok(undefined);
    },
    // ...
  };
};
```

In a test, reach for the `errors` knob — never for a mocking library.

```ts
it('when the post-write fails, the row is counted as errored and the batch completes', async () => {
  const sheets = createSheetsFake({
    tabs: { POST: [row1, row2] },
    errors: { appendOrUpdate: { kind: 'write-failed', message: 'quota' } },
  });
  const telegram = createTelegramFake();
  const logger = createLoggerFake();

  const result = await createPostTelegram({ sheets, telegram, logger })(input);

  expect(result.ok).toBe(true);
  expect(result.ok && result.value).toEqual({ published: 0, errored: 2 });
  expect(logger.calls.filter((c) => c.event === 'sheets.write.failed')).toHaveLength(2);
});
```

The error-map value must be the port's discriminated-union error shape — cast literals to `{ kind: 'write-failed' as const, message: 'quota' }` so TypeScript narrows to the union member.

### Rejection assertions: `captureRejection`

SonarJS S4123 fires on `await expect(p).rejects.toThrow(...)` because the matcher chain is not recognised as a `Thenable`. The fix is a tiny helper that reads more clearly anyway. The canonical implementation ships in `assets/capture-rejection.ts` — copy it verbatim (don't hand-retype `String(e)` into the non-Error branch; that would trip S6551, the very rule the helper set exists to respect). Shape:

```ts
// src/test-helpers/capture-rejection.ts — see assets/capture-rejection.ts for the full body
export const captureRejection = async (promise: Promise<unknown>): Promise<Error> => {
  try {
    await promise;
  } catch (e) {
    if (e instanceof Error) return e;
    // formatNonError avoids String(e) (SonarJS S6551) — full impl in the asset.
    // { cause: e } preserves the original value (ESLint preserve-caught-error).
    throw new Error(`captureRejection: rejected with non-Error value: ${formatNonError(e)}`, { cause: e });
  }
  throw new Error('captureRejection: expected promise to reject, but it resolved');
};

// test
const err = await captureRejection(doSomethingThatThrows());
expect(err.message).toBe('expected message');
```

Use this for every promise-rejection assertion in `*.test.ts`. Never the matcher chain.

### Testing the `unwrap` invariant

In test code where the error branch is genuinely unreachable (the fake does not have `errors` set, so the call cannot err), `unwrap` is acceptable. It makes intent explicit and keeps the test body short.

```ts
const result = unwrap(await createPostTelegram(deps)(input));
expect(result).toEqual({ published: 2, errored: 0 });
```

If `unwrap` ever actually throws in a test, the test name is wrong or the fake setup is incomplete.

## Migration checklist when a port starts returning `Result`

1. Declare `PortError` as a discriminated union next to the port type. Give each variant a `kind` string literal and a `message`.
2. Change the port's return type from `Promise<T>` to `Promise<Result<T, PortError>>`.
3. Update every adapter in `src/infra/` to wrap its third-party call in `try/catch` and return `err(classify(e))`.
4. Update every caller in `src/use-cases/` to pattern-match on `.ok`. Delete any existing `try/catch` around port calls.
5. Update the fake in `src/test-helpers/` to add an `errors` knob.
6. Replace any `withRetry(() => port.call())` with `retryOnErr(...)` and an explicit "retry on kind X" predicate.
7. Run the 4-check loop (`bun test`, `bun run lint`, `bun run typecheck`, `bun run coverage`). See `references/workflow.md`.
