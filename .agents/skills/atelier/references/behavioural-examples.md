# Behavioural examples (before → after)

Concrete before/after pairs for the four Behavioural Guidelines in `SKILL.md`, written in this repo's idiom (arrow functions, no `class`, `Result<T, E>`, branded types, fakes not mocks). Read alongside the guidelines; this file is the worked-example layer.

The key framing: the "before" versions are rarely *wrong* in isolation — they follow real patterns (dispatch maps, DI factories, validation). The defect is **timing and reach**: complexity added before it is needed, or edits that travel past the request. Good code solves today's problem simply, not tomorrow's prematurely.

---

## 1. Think before coding

### Hidden assumptions

**Request:** "Add a feature to export user data."

✗ **Before** — silently decides scope, destination, fields, format, and ships a use-case that exports every user to a hardcoded path. Four assumptions, none surfaced.

✓ **After** — surface the assumptions before writing code:

```
Before implementing I need to pin down four things:
1. Scope — all users, or a filtered/authorised subset? (privacy + volume)
2. Delivery — download, an API endpoint returning JSON, or a background job + notification?
3. Fields — which attributes? Some are PII and must cross a SafeExport checkpoint.
4. Volume — rough row count? It decides streaming vs in-memory.

Simplest thing that could work: an authenticated use-case returning a paginated
Result<UserPage, ExportError>. File export is more work — confirm before I build it.
```

### Multiple interpretations — present options with effort

**Request:** "Make the search faster."

✗ **Before** — picks one meaning silently and writes 200 lines of caching + indexing + async.

✓ **After** — name the interpretations and attach rough effort/tradeoff so the choice is informed:

```
"Faster" could mean three different things — current search is ~500ms typical:
1. Lower latency (cache hot queries, add an index) — ~½ day, risk: cache invalidation.
2. More throughput (connection pooling, batch the IO port) — ~1 day, risk: pool tuning.
3. Faster *perceived* speed (stream partial results) — ~1 day, mostly presenter-layer.
Which matters for your use case? They pull the design in different directions.
```

### Explore before asking; one question at a time

**Request:** "Add caching to the guides loader."

✗ **Before** — fires a batch of questions the codebase already answers:

```
A few questions before I start:
1. Which cache library should I use?
2. Where are guides loaded?
3. What should the cache key be?
4. How long should entries live?
5. Is this server- or client-side?
```

✓ **After** — answer from the code what the code can answer, then ask the *one* real question, leading with a recommendation:

```
Checked the code: guides load in src/lib/guides/guides.ts via Bun.file at build
time (static export — next.config.ts has output: 'export'), so the loader runs
once per build, not per request. Most "caching" is therefore moot.

The one real decision: memoize the parsed-MDX result within a single build run?
  Recommended: yes — a module-level Map keyed by slug, ~5 lines, no dependency
  (lazy ladder rung 3). I'd skip a cache library entirely.

Proceed with that, or did you mean per-request caching? (That implies dropping
static export — a much bigger change, so I want to confirm before assuming it.)
```

Four questions evaporated because the codebase answered them; the fifth became a one-line yes/no with a recommendation and a flagged dependency (static export). That is the difference between handing the user homework and handing them a decision.

---

## 2. Simplicity first

### Over-abstraction — even in a class-free idiom

**Request:** "Add a function to calculate a discount."

✗ **Before** — a strategy dispatch map + a config record + a DI factory, for one discount type. Class-free, fully SOLID-shaped — and 25 lines of ceremony before the first number is computed:

```ts
type DiscountStrategy = (amount: Money) => Money;

const discountStrategies: Record<DiscountKind, (arg: number) => DiscountStrategy> = {
  percentage: (pct) => (amount) => scaleMoney(amount, pct / 100),
  fixed: (cap) => (amount) => minMoney(money(cap, amount.currency), amount),
};

type DiscountConfig = { strategy: DiscountStrategy; minPurchase: Money; maxDiscount: Money };

export const createDiscountCalculator =
  (config: DiscountConfig) =>
  (amount: Money): Money => {
    if (isLessThan(amount, config.minPurchase)) return money(0, amount.currency);
    return minMoney(config.strategy(amount), config.maxDiscount);
  };
```

✓ **After** — one arrow function over the branded `Money`/`Percentage` types:

```ts
export const percentageDiscount = (amount: Money, percent: Percentage): Money => scaleMoney(amount, percent / 100);
```

Open/Closed via a dispatch map is *correct* once you genuinely have several discount kinds (SOLID "O"). With one, it is speculation. The Rule of Three governs abstractions too — extract the strategy map at the third kind, not the first.

### Speculative features

**Request:** "Save user preferences to the database."

✗ **Before** — a `createPreferenceManager` with optional cache, validator, and `merge`/`validate`/`notify` flags. Every branch is a feature nobody asked for, and each is now a thing to test and maintain.

✓ **After** — one use-case that saves, returning a `Result` at the IO boundary:

```ts
export const savePreferences =
  (deps: { prefs: PreferencesRepo }) =>
  (userId: UserId, prefs: Preferences): Promise<Result<void, PreferencesError>> =>
    deps.prefs.save(userId, prefs);
```

Add caching when a profile shows a hot path; validation when bad data actually appears — and then at a branded-type checkpoint, not behind a boolean flag.

---

## 3. Surgical changes

### Drive-by refactoring

**Request:** "Fix the bug where an empty email crashes the signup parser."

✗ **Before** — fixes the crash *and* adds username length/charset rules nobody asked for:

```ts
const parseSignup = (raw: RawSignup): Result<Signup, ValidationError> => {
  if (raw.email.trim().length === 0) return err({ type: 'validation', field: 'email' });
  const address = email(raw.email);
  if (raw.username.length < 3) return err({ type: 'validation', field: 'username' }); // not requested
  if (!/^[a-z0-9]+$/iu.test(raw.username)) return err({ type: 'validation', field: 'username' }); // not requested
  return ok({ address, username: username(raw.username) });
};
```

✓ **After** — only the empty-email guard; the username path is left exactly as it was:

```ts
const parseSignup = (raw: RawSignup): Result<Signup, ValidationError> => {
  if (raw.email.trim().length === 0) return err({ type: 'validation', field: 'email' });
  const address = email(raw.email);
  return ok({ address, username: username(raw.username) });
};
```

Every changed line traces to "empty email crashes." Username validation, if wanted, is a separate request and a separate commit.

### Style drift

**Request:** "Add logging to the upload use-case."

✗ **Before** — reaches for `console.log` (a hard-rule violation), retypes the signature, swaps quote style, and reflows whitespace around the change.

✓ **After** — add the injected `Logger` port and one log line, matching the file's existing style:

```ts
export const uploadReport = async (deps: { storage: ReportStorage; logger: Logger }, path: SafePath): Promise<Result<void, UploadError>> => {
  deps.logger.info('uploading report', { path });
  return deps.storage.put(path);
};
```

Only two lines moved: the `logger` dependency and the log call. No reformatting, no retyping, no `console.*`.

---

## 4. Goal-driven execution

### Vague vs verifiable

**Request:** "Fix the authentication system."

✗ **Before** — "I'll review the auth code, find issues, improve them, and test." No success criterion, so there is nothing to loop against.

✓ **After** — turn it into a plan with a verifiable check per step:

```
"Fix auth" needs a concrete failure. If it is "sessions survive a password change":
1. RED   — test: changing the password revokes existing sessions   → fails (reproduces it)
2. GREEN — revoke sessions inside the changePassword use-case       → that test passes
3. Edge  — a second concurrent session is revoked too              → added test passes
4. Guard — full suite stays green                                   → no regression
What is the specific symptom you are seeing?
```

### Test-first, at the primary port

**Request:** "Sessions aren't revoked when the password changes."

✓ **After** — reproduce with a failing test before touching production code. The SUT is the use-case (primary port); the domain runs real; only the secondary ports (`SessionRepo`, `Clock`) are faked:

```ts
test('when a member changes their password, prior sessions are revoked', async () => {
  const sessions = createInMemorySessionRepo({ 'u-1': [session('s-old')] });
  const result = await changePassword(
    { user: userId('u-1'), newPassword: password('correct horse') },
    { sessions, clock: fixedClock('2026-01-01T00:00:00Z') }
  );
  expect(result.ok).toBe(true);
  expect(await sessions.activeFor(userId('u-1'))).toHaveLength(0);
});
```

Red first, then the smallest change that turns it green.

---

## Anti-pattern quick reference

| Guideline | Anti-pattern (this idiom) | Fix |
|:---|:---|:---|
| Think before coding | Silently picks scope / format / fields | List assumptions; present options *with* effort + tradeoff |
| Simplicity first | Dispatch map + DI factory + config record for one case | One arrow function; abstract on the Rule of Three |
| Surgical changes | A bug fix that also adds validation, retypes, reflows | Change only the lines that fix the reported issue |
| Goal-driven execution | "I'll review and improve the code" | RED test reproduces → GREEN → suite stays green |

The through-line is the framing from the top of this file: the "before" versions are defects of **timing and reach**, not of engineering — the simple version is the smallest thing correct today, cheap to extend when tomorrow's requirement actually shows up.
