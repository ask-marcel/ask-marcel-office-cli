# Testing Strategy

## The school: Outside-in classicist

The SUT of every unit test is a **primary port**: a use case, command handler, or application service at the hexagonal boundary. Inside the port, the full domain runs real: entities, value objects, domain services, aggregate roots. The only test doubles are **fakes** for secondary ports (repository, email sender, clock, token decoder, payment gateway, any adapter to the outside world).

> **Note on examples.** Some example port signatures in this file are elided to `Promise<T>` for brevity where error handling is not the lesson. Real IO ports return `Promise<Result<T, PortError>>` and use-cases return `Promise<Result<Summary, StepError>>`: hard rule 16, see `references/result-type.md`.

Benefits:

- Refactoring the domain never breaks tests.
- Tests describe business scenarios, so they read as living documentation.
- The design pressure lands on the right boundary: when a test is hard to write, the port's contract is wrong, not the entity.

The school is Ian Cooper's (*TDD, Where Did It All Go Wrong?*) and the classic Detroit/Chicago school before it. Each of its three rules answers a pattern of test pain the industry learned the hard way: the SUT is the primary port (§ What goes where), domain collaborators run real and only secondary ports get fakes (the table under § What goes where), and there are no mocks, ever (§ No `mock` from `bun:test`). What it buys: internal restructurings never break tests, tests read as specifications a new team member learns the product from, and when a test is hard to write the pressure lands on the port's contract, not on an entity that "needs a helper".

## The loop: Red, Green, Refactor

```
RED -> GREEN -> REFACTOR -> RED -> ...
```

**RED.** Propose a failing test that describes the behaviour you want, and get the user's confirmation before writing it (SKILL.md hard rule 24: tests are confirmation-gated; never create, change, or delete one silently). The test uses domain language, describes WHAT rather than HOW, and is a concrete example, not an abstract statement:

```ts
// BAD - abstract
it('can add numbers', () => { /* ... */ });

// GOOD - concrete example
it('when adding 2 + 3, returns 5', () => { /* ... */ });
```

**GREEN.** Write the simplest possible code to make the test pass. Two strategies: *Fake It* (return a hardcoded value, `export const add = (a: number, b: number): number => 5;`) and *Obvious Implementation* (`a + b`) when you know the solution. Prefer Fake It when learning or unsure; let more tests drive the real implementation.

**REFACTOR.** This is where design happens. Look for duplication (wait for the Rule of Three), functions longer than 10 lines to extract, poor names to improve, complex conditions to simplify, raw primitives in domain positions to promote to branded types.

### The Three Laws

1. No production code unless it makes a failing test pass.
2. No more test code than sufficient to fail (compilation failures count).
3. No more production code than sufficient to pass the one failing test.

### Bug fixes: the regression test comes first

A bug is a missing test: the suite said green while the behaviour was wrong, so the suite has a hole exactly bug-shaped. Fixing the code without first filling the hole leaves you with no proof the fix addresses the actual defect, and nothing to stop the same regression returning. The loop is RED-GREEN-REFACTOR with a sharper RED:

1. **Reproduce as a failing test.** Write the smallest test that fails for the bug's reason, named as the business scenario that went wrong (`'when a refund is issued twice, the second attempt is rejected'`, not `'fix double refund'`). Propose it and get confirmation first (hard rule 24), like any test.
2. **Watch it fail, and read WHY.** The red run must fail with the bug's symptom. A test that fails for a setup error, or passes immediately, does not capture the bug; fix the test, not the code, until the failure is the bug.
3. **Fix to green.** The smallest production change that makes the regression test pass without breaking the rest of the suite.
4. **Refactor**, then look sideways: the same hole often exists in sibling paths (the other branch, the other adapter, the plural endpoint). Each one found gets its own failing test first.

This applies mid-implementation too. When you are building feature A and trip over broken behaviour B, do not silently patch B on the way past: stop, reproduce B as its own failing test (confirmation-gated), fix it, and keep the fix in its own commit, then return to A. The temptation to fold a drive-by fix into an unrelated diff is how untested fixes ship. The only sanctioned inversion is a live incident where mitigation cannot wait for a test: mitigate, say so in the commit body, and make the regression test the first act of the follow-up; the incident is not closed while the hole is open.

### The Rule of Three

Only extract duplication when you see it THREE times: a wrong abstraction is more expensive to undo than duplication is to tolerate. Duplication #1, leave it. #2, note it, leave it. #3, now extract it.

### Triangulation

Each new test sculpts the solution toward a general implementation; think of degrees of freedom, each test carving out one until the implementation handles all cases. Implementing `isPalindrome`: test `'mom'` and fake it (`return true`); test `'hello'` and the fake fails, so generalise (compare halves); test `''` and the edge case forces explicit handling; test `'racecar'` and the general case is confirmed.

### Transformation Priority Premise

When going from RED to GREEN, prefer the simpler transformation; higher in this list is simpler, and jumping to a complex one too early is how speculative code arrives.

| Priority | Transformation |
|:---|:---|
| 1 | `{}` to `null` |
| 2 | `null` to constant |
| 3 | constant to variable |
| 4 | unconditional to conditional |
| 5 | scalar to collection |
| 6 | statement to recursion |
| 7 | value to mutated value |

## Test the code you own; trust your dependencies

The first question before writing any test is *whose behaviour am I pinning?* Test only the code this repo owns. Never write a test whose real assertion is that a third-party library, the runtime, or the framework behaves as documented: that test pins someone else's contract, breaks when they release, and proves nothing about your code. Trust your dependencies; if one is genuinely suspect, the answer is to pin its version (hard rule 19) or replace it, not to grow a test suite around it.

This single principle is why several other rules look the way they do:

- **Adapters test the translation, not the SDK.** An infra adapter's job is to turn a library's contract into `Result<T, PortError>`. The test feeds a slice of the SDK's real surface (the two-constructor pattern, hard rule 13) and asserts that *your* mapping of success and error is correct, not that the SDK itself works. You are testing the seam, not the library behind it. See the infra-adapter section below and `references/testing-infra.md`.
- **SDK-bridge lines are coverage-exempt.** A line whose only job is to construct or call into a third-party SDK has no behaviour of yours to cover, so it is exempt from the line-coverage gate rather than wrapped in a contortion test. See `references/workflow.md` (SDK-bridge lines).
- **Domain pieces are used, not tested.** Entities, value objects, and domain services run real inside a primary-port test (the classicist rule above). You own them, but you pin their behaviour *through the port*, not in isolation, so they stay free to refactor. The one exception, a few direct tests for a value object or domain service with genuinely non-trivial logic, is spelled out under Value-object tests below (hard rule 14 names it too); a regex-shaped id is not it.
- **Prop-pure components are not unit-tested.** A design-system component is a deterministic prop→JSX map with no logic of its own (hard rule 21); there is nothing to own a test. It is covered by the design-system lint block and review, never by React Testing Library ceremony that re-proves React renders props.

The same instinct underlies the mock ban (hard rule 13): you write fakes for the secondary-port contracts *you define*, and for code you do not own you inject a thin slice of its real surface, you never reach into a dependency to puppet it.

## The testing pyramid

```
         /\
        /  \        E2E / Acceptance tests (FEW)
       /----\       full system, slow, brittle
      /      \
     /--------\
    /          \    Integration tests (SOME)
   /            \   real secondary-port adapters
  /--------------\
 /                \  Unit tests (MANY)
/                  \ primary-port SUT, real domain, faked secondary ports
--------------------
```

## Test types

### Unit tests

A unit is a **behaviour**, not a function. The SUT is a primary port; the domain runs real; secondary ports are fakes. Most tests in the codebase are unit tests.

```ts
import { describe, expect, it } from 'bun:test';
import { placeOrder } from './place-order';
import { createInMemoryOrderRepo } from './fakes/in-memory-order-repo';
import { createInMemoryCustomerRepo } from './fakes/in-memory-customer-repo';
import { money } from '../money/money';
import { customerId } from '../customers/customer-id';

describe('placeOrder', () => {
  it('when a premium customer buys 100 EUR, the order total is 80 EUR', async () => {
    const orders = createInMemoryOrderRepo();
    const customer = customerId('c-1');
    const customers = createInMemoryCustomerRepo({ [customer]: { tier: 'premium' } });

    await placeOrder(
      { customer, items: [{ sku: 'SKU-1', price: money(10_000, 'EUR') }] },
      { orders, customers }
    );

    const [saved] = await orders.findByCustomer(customer);
    expect(saved.total).toEqual(money(8000, 'EUR'));
  });
});
```

Notice what is **real**: the `placeOrder` use case, every domain function it calls, the `Money` value object, the `Order` entity, the pricing rules. What is **faked**: `orders` and `customers`, the two secondary ports.

### Integration tests

Test secondary-port adapters against the real outside world: a real database, a real HTTP API (in a sandbox), a real queue. These prove that the `postgres*Repo` fulfils the same contract as `createInMemory*Repo`. Fewer than unit tests; run in a separate CI stage.

```ts
describe('postgresOrderRepo', () => {
  let repo: OrderRepo;

  beforeAll(async () => {
    repo = createPostgresOrderRepo(testDb);
  });

  it('saves an order and retrieves it by customer', async () => {
    const customer = customerId('c-1');
    const order = buildOrder({ customer, total: money(8000, 'EUR') });
    await repo.save(order);
    const [found] = await repo.findByCustomer(customer);
    expect(found).toEqual(order);
  });
});
```

Contract tests (below) let you run the same assertions against the in-memory fake and the Postgres adapter, so divergence between them is caught automatically.

### E2E / acceptance tests

Drive the real user interface against a deployed stack. Slowest, most brittle. Critical paths only.

```ts
describe('checkout flow', () => {
  it('premium customer buys one item at 100 EUR and sees Order Confirmed', async () => {
    await page.goto('/products');
    await page.click('[data-testid="add-to-cart"]');
    await page.click('[data-testid="checkout"]');
    await page.fill('[name="card"]', '4242424242424242');
    await page.click('[data-testid="pay"]');

    expect(await page.textContent('h1')).toBe('Order Confirmed');
  });
});
```

### Performance / load tests

The layer the other three ignore: unit, integration, and E2E all prove correctness at n=1. When a route has a latency budget (`references/reliability.md`, Performance is a budget), a load test proves it under production-like traffic and fails the pipeline when breached, e.g. k6 with `thresholds: { http_req_duration: ['p(99)<300'] }` at the expected peak. Few of these: hot endpoints and known-risky queries, not every route. The pagination and N+1 disciplines in `references/reliability.md` are what make passing them possible as tables grow.

### Regression tests (every fixed bug becomes one)

A bug fix without a test is a bug scheduled to return. The loop for any defect: write the test that reproduces it first (red), name it after the defect in domain language with a `regression:` prefix, then fix production code until green (SKILL.md, Behavioural Guideline #4). The test is permanent; it never gets deleted when the code it pins is refactored.

```ts
test('regression: empty cart totals to zero, not NaN', () => {
  expect(total([])).toBe(0);
});
```

The same discipline extends to LLM holes: a production miss becomes a labeled eval case (`references/ai.md`).

---

## Arrange-Act-Assert

Structure EVERY test this way. The ACT should call the primary port; the ASSERT should read state from a fake (or the returned result).

```ts
it('when a premium customer buys a 100 EUR item, the order total is 80 EUR', async () => {
  // ARRANGE - real domain, faked secondary ports
  const orders = createInMemoryOrderRepo();
  const customer = customerId('c-1');
  const customers = createInMemoryCustomerRepo({ [customer]: { tier: 'premium' } });

  // ACT - call the primary port
  await placeOrder(
    { customer, items: [{ sku: 'SKU-1', price: money(10_000, 'EUR') }] },
    { orders, customers }
  );

  // ASSERT - read state from the fake
  const [saved] = await orders.findByCustomer(customer);
  expect(saved.total).toEqual(money(8000, 'EUR'));
});
```

### Writing AAA backwards

When stuck, write the test in reverse: the ASSERT first (what do you want to verify?), then the ACT (what action produces that result?), then the ARRANGE (what setup does that action need?).

---

## Test naming

Every test name is a **complete business scenario** in domain language. Not the name of a function, not "should work correctly", not "happy path". A reader who has never seen the code should understand the scenario from the title alone.

### Bad | technical, function-oriented

```ts
it('should work correctly', () => { /* ... */ });
it('handles the edge case', () => { /* ... */ });
it('getDiscount returns 20 when tier is premium', () => { /* ... */ });
it('calculateTotal applies tax', () => { /* ... */ });
```

### Good | business scenarios

```ts
it('when a premium customer buys a 100 EUR item, the order total is 80 EUR', () => { /* ... */ });
it('when the cart is empty, checkout is rejected', () => { /* ... */ });
it('when a VAT-registered EU customer orders, the invoice shows no VAT', () => { /* ... */ });
```

### Format options

```ts
// option 1 | when <scenario>, then <outcome>
it('when adding a 100 EUR item to an empty cart, the total is 100 EUR', () => { /* ... */ });

// option 2 | given <context>, when <scenario>, then <outcome>
describe('given a premium customer', () => {
  describe('when they check out a 100 EUR cart', () => {
    it('the order total is 80 EUR and a confirmation email is sent', () => { /* ... */ });
  });
});
```

Avoid titles that name functions (`getDiscount`, `calculateTotal`, `isValid`). If a title contains a function name, the test is almost certainly targeting the wrong SUT.

---

## Test doubles

Three shapes are permitted: **dummy**, **stub**, **fake**. Hand-written spies (a fake that also records its inputs) are allowed when outcome assertions are not enough. Mocks from a mock library are banned: see the "No mocks" rule below.

### Dummy

A record passed but never used. Satisfy the port with real no-ops, never `{} as Logger`, which is the non-narrowing `as` cast the skill bans.

```ts
const dummyLogger: Logger = { info: () => {}, warn: () => {}, error: () => {} };
const service = createUserService(realRepo, dummyLogger);
```

### Stub

Returns predefined values.

```ts
const stubRepo: UserRepo = {
  findById: async () => buildUser({ name: name('Test'), email: email('t@t.com') }),
  save: async () => {},
};
```

### Fake (preferred)

A working in-memory implementation of the contract.

```ts
export const createInMemoryUserRepo = (): UserRepo => {
  const store = new Map<UserId, User>();
  return {
    save: async (user) => {
      store.set(user.id, user);
    },
    findById: async (id) => store.get(id) ?? null,
  };
};
```

Fakes let tests assert on final state (the thing the domain actually cares about) rather than on call sequences, so they survive refactoring.

### Fakes with an `errors` knob

When the code under test returns `Result<T, E>`, the fake needs an optional `errors` config so tests can hit the error branch without a mocking library. Every port fake exposes this knob.

```ts
export const createSheetsFake = (config?: {
  tabs?: Partial<Record<string, ReadonlyArray<SheetRow>>>;
  errors?: { readRows?: SheetsError; appendOrUpdate?: SheetsError; deleteRow?: SheetsError };
}): Sheets => {
  // each operation checks its errors entry first: if set, return err(config.errors.<op>);
  // otherwise act on an in-memory store seeded from config.tabs.
};
```

Full implementation: `references/result-type.md` (fakes with error injection).

### Batch use-cases: `ok(summary)` with an `errored` count

Use-cases that iterate over many rows catch per-row port errors internally, increment an `errored` counter, and return `ok({ published, errored })`. They do **not** return `err(...)` per row. Tests assert on the summary and (optionally) on the logger-fake calls:

```ts
it('when one of three rows fails to post, the batch completes with errored=1', async () => {
  const sheets = createSheetsFake({ tabs: { POST: [row1, row2, row3] } });
  const telegram = createTelegramFake({ errors: { [row2.channel]: { kind: 'rate-limited', message: '429' } } });
  const logger = createLoggerFake();

  const result = unwrap(await createPostTelegram({ sheets, telegram, logger })(input));

  expect(result).toEqual({ published: 2, errored: 1 });
  expect(logger.calls.filter((c) => c.level === 'warn')).toHaveLength(1);
});
```

`err(...)` from a batch use-case is reserved for prerequisites: the initial `sheets.readRows` fails, or credentials are missing. See `references/result-type.md` for the full rationale.

### Hand-written spy

When a test must assert that an outbound call happened (e.g. a notification was sent), write a fake that records its inputs in a field. No mocking library.

```ts
type EmailSpy = { sentEmails: Email[]; send: (to: Email, message: string) => Promise<void> };

export const createEmailSpy = (): EmailSpy => {
  const sentEmails: Email[] = [];
  return {
    sentEmails,
    send: async (to) => {
      sentEmails.push(to);
    },
  };
};

// assert on state, not on call sequences
expect(spy.sentEmails).toContain(email('user@example.com'));
```

### No `mock` from `bun:test` (absolute, enforced by lint)

The entire mocking surface of `bun:test` is banned: `mock()`, `mock.module()`, `spyOn`, and the `jest` and `vi` compatibility objects, plus every call-recording assertion (`.toHaveBeenCalled*`, `.toBeCalled*`, `.toHaveReturned*`) whatever built the spy. The canonical config in `references/bun-typescript.md` (`eslint.config.js`) enforces both halves: `MOCK_BAN` in `no-restricted-imports` names the four imports, and a `STYLE_BANS` selector rejects the assertions. A hand-written fake records what it received, and the test asserts on that record.

```ts
// BANNED
import { mock } from 'bun:test';
const mockSave = mock(async (_user: User): Promise<void> => {});
expect(mockSave).toHaveBeenCalledWith(expectedUser);

// BANNED (module substitution is process-global and leaks across test files)
mock.module('googleapis', () => ({ google: { drive: () => fakeApi } }));

// REQUIRED: fake the port for use-case tests
const repo = createInMemoryUserRepo();
await placeOrder(order, { repo });
expect(await repo.count()).toBe(1);

// REQUIRED: pass the API slice for infra adapter tests (see "Testing infra adapters")
const api: DriveApi = { files: { copy: async () => ({ data: { id: 'X' } }), /* ... */ } };
const drive = createDriveFromApi(api);
```

Why the absolute ban:

- **`mock.module` is process-global, not file-scoped.** Once set in any test file, every subsequent file the runner loads sees the substitution. This silently corrupted an unrelated `sleep.test.ts` in production use. There is no per-file restore; the leak is a feature of Bun's module cache.
- **`mock()` leaks without `mock.restore()` discipline.** Easy to forget; leak detection is best-effort.
- **Mocks test call sequences, not outcomes.** A mock passes when the right method is called with the right arguments, even if the production code does nothing useful afterwards. A fake passes only when the final state is correct, which is what the system is actually for.
- **Mocks couple tests to implementation.** Rename a method, split a call into two, extract a helper: the mock expectations break even though behaviour is unchanged. The fake keeps passing because the observable state is the same.
- **Mocks hide design pressure.** If you need a mock to test something, the contract is probably too fat (Interface Segregation), or the adapter is missing its `createXFromApi(api)` factory. Fix the design; do not reach for a mock.

`installFetchMock` (see "Testing infra adapters") and per-file `globalThis.setTimeout` swaps are **not** `mock.module`: they swap a global within a lifecycle hook (`afterEach`, `afterAll`) that always restores. The scope is bounded to the test file, not the process.

---

## What goes where

Domain collaborators are real; only secondary ports get fakes. The secondary ports are the ones that talk to the outside world (databases, HTTP, the clock, the filesystem, random sources) and they are the only things that need a double; everything else runs for real inside the test. This is the single most important property of the school: the domain can be refactored freely (rename an entity, split a domain service, merge three value objects, change the shape of an aggregate) and the tests keep passing, because they describe behaviour at the port, not structure inside.

| Kind | Role | Treatment in tests |
|:---|:---|:---|
| Entity | `Order`, `User`, `Subscription` | Real |
| Value object | `Money`, `Email`, `OrderId` | Real |
| Domain service | `pricingRules`, `discountPolicy` | Real |
| Aggregate root | `Order`, `Cart` | Real |
| Primary port | `placeOrder`, `registerUser`, `checkoutCart` | **The SUT** |
| Secondary port | `OrderRepo`, `EmailSender`, `Clock`, `TokenDecoder`, `PaymentGateway` | **Faked** (hand-written in-memory) |

### Unit tests: primary port as SUT (the default)

Most tests. The SUT is a use case, command handler, or application service. The domain runs real; secondary ports are faked.

```ts
describe('placeOrder', () => {
  it('when a premium customer buys a 100 EUR item, the order is saved with a 80 EUR total and a confirmation email is queued', async () => {
    const orders = createInMemoryOrderRepo();
    const emails = createEmailSpy();
    const customer = customerId('c-1');
    const customers = createInMemoryCustomerRepo({ [customer]: { tier: 'premium' } });

    await placeOrder(
      { customer, items: [{ sku: 'SKU-1', price: money(10_000, 'EUR') }] },
      { orders, customers, emails }
    );

    const [saved] = await orders.findByCustomer(customer);
    expect(saved.total).toEqual(money(8000, 'EUR'));
    expect(emails.sentEmails).toContain(email('c-1@example.com'));
  });
});
```

### Value-object / domain-service tests (the exception)

If a value object or a domain service has genuinely complex logic of its own (`Money.add` with currency rules, `PricingPolicy` with tier brackets, `DateRange.overlaps`), a handful of small direct tests is fine. They supplement the primary-port tests, they do not replace them. Keep them rare and only when the logic is non-trivial enough that discovering it through a use-case test would be confusing.

```ts
describe('Money.add', () => {
  it('adds two amounts with the same currency', () => {
    expect(addMoney(money(1000, 'EUR'), money(2000, 'EUR'))).toEqual(money(3000, 'EUR'));
  });

  it('refuses to add different currencies', () => {
    expect(() => addMoney(money(1000, 'EUR'), money(1000, 'USD'))).toThrow('CurrencyMismatch');
  });
});
```

A rough signal: if you find yourself writing more direct value-object tests than primary-port tests, something is off. The use case is where the business value lives; that is where most tests should point.

**The loop on a pure domain function**, which is what these exception tests look like when driven test-first. Build `calculateDiscount` (tier brackets over money; in the use-case that calls it, it runs real behind the primary port):

```ts
import { describe, expect, it } from 'bun:test';
import { calculateDiscount } from './calculate-discount';
import { money } from '../money/money';

describe('calculateDiscount', () => {
  it('when standard customer buys 100 EUR, returns 0', () => {
    const subtotal = money(10_000, 'EUR');
    const result = calculateDiscount(subtotal, 'standard');
    expect(result.cents).toBe(0);
  });

  it('when premium customer buys 100 EUR, returns 20 EUR', () => {
    const subtotal = money(10_000, 'EUR');
    const result = calculateDiscount(subtotal, 'premium');
    expect(result.cents).toBe(2000);
  });
});
```

RED: test 1 fails (no `calculate-discount.ts` yet). GREEN, fake it:

```ts
import type { Money } from '../money/money';
import { money, scaleMoney } from '../money/money';

type CustomerTier = 'standard' | 'premium';

export const calculateDiscount = (subtotal: Money, tier: CustomerTier): Money => money(0, subtotal.currency);
```

Test 1 passes, test 2 fails. Generalise:

```ts
export const calculateDiscount = (subtotal: Money, tier: CustomerTier): Money => {
  if (tier === 'premium') return scaleMoney(subtotal, 0.2);
  return money(0, subtotal.currency);
};
```

REFACTOR: no duplication to extract, names are clear, the function is four lines; move on to the next test. When a third tier appears (`vip`), resist extracting until after the third `if` branch exists (Rule of Three), then promote the logic to a dispatch record:

```ts
const tierRates: Record<CustomerTier, number> = {
  standard: 0,
  premium: 0.2,
  vip: 0.3,
};

export const calculateDiscount = (subtotal: Money, tier: CustomerTier): Money =>
  scaleMoney(subtotal, tierRates[tier]);
```

This is what "design happens during refactor" looks like.

### Branded types and `expect(...).toBe(raw)`, the test escape hatch

Bun's `expect(x).toBe(y)` matcher infers `y`'s type from `x`. When `x` has a branded type, `y` must be the same brand or TypeScript fails:

```ts
const tok = accessToken('eyJ...'); // accessToken: (s: string) => AccessToken
expect(tok).toBe('eyJ...');
//             ^^^^^^^^ Argument of type 'string' is not assignable to 'AccessToken'.
```

Three options. Use the third.

1. ❌ `as` the raw string: assertions are forbidden everywhere else, do not start in tests.
2. ❌ Run the value through the real factory in the assertion (`expect(tok).toBe(accessToken('eyJ...'))`): works, but the factory may have side effects (logging, parsing) that you don't want in a hot test loop.
3. ✅ **Export an `xxxUnsafe(raw): X` helper next to the factory. Use it only in tests.**

```ts
// src/domain/access-token.ts
export type AccessToken = string & { readonly __brand: 'AccessToken' };

export const accessToken = (value: string): AccessToken => {
  if (value.length === 0) throw new Error('AccessToken: empty');
  return value as AccessToken;
};

// Test escape hatch: bypasses validation. Naming convention: <factory>Unsafe.
// Production code MUST NOT import this; the only callers are *.test.ts files.
export const accessTokenUnsafe = (value: string): AccessToken => value as AccessToken;
```

```ts
// access-token.test.ts
import { accessToken, accessTokenUnsafe } from './access-token.ts';

it('round-trips through the factory', () => {
  expect(accessToken('eyJ...')).toBe(accessTokenUnsafe('eyJ...')); // both sides are AccessToken
});
```

**Naming convention.** `<factoryName>Unsafe`: `accessTokenUnsafe`, `envVarUnsafe`, `userIdUnsafe`, `safeUrlUnsafe`. The `Unsafe` suffix tells the next reader (and the next grep) exactly what they're looking at: a brand cast without validation, for tests only.

**Boundary.** Production code must not import any `*Unsafe` helper, and the shipped ESLint configs enforce it: `UNSAFE_BAN` (`{ group: ['**'], importNamePattern: 'Unsafe$' }`, ESLint 8.55 and later) rides inside every layer zone except `test-helpers`, whose fakes may build values with the helpers, and the zones already skip `*.test.ts`. It lives inside the zones, never in a block of its own: flat config replaces a rule's options, so a separate `no-restricted-imports` block for `src/**` would wipe the mock ban (rule 13) and the zone (rule 37) for every file it matches, or, placed before them, be wiped itself. The Bun smoke test plants a domain file importing `accessTokenUnsafe` and sees it red.

### Secondary-port integration tests

Integration tests prove the real adapter (Postgres, SendGrid, Redis) fulfils the contract its in-memory fake already satisfies. Run in a separate CI stage with real infrastructure.

```ts
describe('postgresOrderRepo', () => {
  let repo: OrderRepo;

  beforeAll(async () => {
    repo = createPostgresOrderRepo(testDb);
  });

  it('saves an order and finds it by customer', async () => {
    const order = buildOrder({ customer: customerId('c-1'), total: money(8000, 'EUR') });
    await repo.save(order);
    const [found] = await repo.findByCustomer(customerId('c-1'));
    expect(found).toEqual(order);
  });
});
```

Pair these with contract tests (below) so the fake and the real adapter cannot drift.

---

## Testing infra adapters

Infra adapters need their own playbook because their job is to translate a third-party library's contract into `Result<T, PortError>`. Three patterns cover every adapter shape (HTTP via `fetch`, external SDK, filesystem), plus a production-wiring smoke test and the silent-gotcha around fetch-mock handler ordering.

See `references/testing-infra.md` for the full treatment with worked examples.

---

## High-value integration tests

Focus integration tests on:

1. Boundaries | where systems meet.
2. Critical paths | money, security, core features.
3. Complex queries | database operations.

### Contract tests

Run the same test suite against every implementation of a contract.

```ts
const testUserRepoContract = (createRepo: () => UserRepo): void => {
  describe('UserRepo contract', () => {
    let repo: UserRepo;
    beforeEach(() => {
      repo = createRepo();
    });

    it('saves and retrieves a user', async () => {
      const user = buildUser({ name: name('Test'), email: email('t@t.com') });
      await repo.save(user);
      const found = await repo.findById(user.id);
      expect(found).toEqual(user);
    });

    it('returns null for a missing user', async () => {
      const found = await repo.findById(userId('nope'));
      expect(found).toBeNull();
    });
  });
};

// apply to all implementations
testUserRepoContract(() => createInMemoryUserRepo());
testUserRepoContract(() => createPostgresUserRepo(testDb));
```

This catches "I implemented the fake differently from the real one" bugs.

---

### Bypass tests (assert the refusal, not just the success)

A guard proves nothing until a test walks the forbidden path and is refused. The happy-path test ("admin can purge") would pass even if the role check were missing; the bypass test is the one that fails on the real defect:

```ts
test('non-admin purge is refused', async () => {
  const res = await app.request('/v1/admin/purge', authAs(staffUser));
  expect(res.status).toBe(403);
});

test('missing token is unauthorized', async () => {
  const res = await app.request('/v1/admin/purge');
  expect(res.status).toBe(401);
});
```

Three refusals every protected surface ships: the **lower-privilege role** (403), the **missing/invalid token** (401), and the **cross-tenant reach** (404, so existence is not disclosed; `references/isolation.md`). And test the seam a request actually travels: drive the real edge with a forged trust header (`x-org-id` set by the attacker) and assert it is inert, because the gap between two individually-correct systems is where real attacks live. A control nobody has tried to get around protects nothing (SKILL.md red flags; `references/workflow.md`, Verification discipline).

## Test builders

Create test records easily. A builder is just a factory function with sensible defaults.

```ts
export type OrderConfig = {
  readonly id?: OrderId;
  readonly customerId?: CustomerId;
  readonly items?: readonly Item[];
  readonly status?: OrderStatus;
};

export const buildOrder = (overrides: OrderConfig = {}): Order => ({
  id: overrides.id ?? orderId('ord-1'),
  customerId: overrides.customerId ?? customerId('cust-1'),
  items: overrides.items ?? [],
  status: overrides.status ?? 'pending',
});

// usage
const pending = buildOrder();
const paid = buildOrder({ status: 'paid' });
const withItems = buildOrder({ items: [item({ sku: 'ABC', price: money(10_000, 'EUR') })] });
```

---

## Random order: no test waits for another

The suite runs in a random order on every run, and every test is written so any order is green. The test script is `bun test --randomize` (Bun and Next.js variants), the same command in `package.json`, `ci.yml`, Stryker's command runner and the inner loop, so the shuffle is never something a developer opts into. A red run prints `--seed=<n>`; rerun with that seed to replay the order that failed, fix the dependency, then drop the seed. Java: `src/test/resources/junit-platform.properties` carries `junit.jupiter.testmethod.order.default=org.junit.jupiter.api.MethodOrderer$Random` and `junit.jupiter.testclass.order.default=org.junit.jupiter.api.ClassOrderer$Random`; `junit.jupiter.execution.order.random.seed=<n>` replays a failing order, and since JUnit logs its own seed below INFO, `ci-java.yml` picks the seed and prints it.

What the rule forbids in practice: a module-level `let` one test mutates and a later test reads; a fake shared across tests without a fresh instance per test; a file whose tests assume the previous file ran; `@Order` or `@TestMethodOrder(OrderAnnotation.class)` used to sequence a read after a write. What it asks for: fresh fakes per test (the `errors` knob pattern above), builders that create the state each test needs, and cleanup in the test that made the mess. State shared between tests is a defect the next shuffle exposes, never a flake to retry (hard rule 36).

---

## Common testing mistakes

| Mistake | Problem | Fix |
|:---|:---|:---|
| Writing code before the test | The fundamental inversion; the test then documents what was built, not what was wanted | RED first, always |
| Writing too much test, or too much code | The loop loses its grip; speculative branches arrive untested | Just enough test to fail, just enough code to pass, then refactor |
| Skipping refactor | Design never happens; duplication and poor names accrete | Refactor on every green; extract only on the third duplication |
| Abstract test names | Nothing to learn the product from | Concrete examples in domain language, one behaviour per test |
| Reaching for doubles too soon | Real collaborators would have caught the integration | Start real; if a double is needed, write a fake, never a mock |
| Testing implementation | Brittle tests | Test observable behaviour only |
| Using mocks | Tests prove call sequences instead of outcomes; break on refactor | Never use mocks: write a fake for the contract |
| Testing only the happy path of a guard | A missing role/tenant check still passes every test | Ship the refusal tests: 403 wrong role, 401 no token, 404 cross-tenant (Bypass tests above) |
| Fixing a bug without a test | The same defect returns unnoticed | Reproduce red first; keep it as a permanent `regression:` test |
| Shared state between tests | A test that only passes in one order; rule 36 makes it fail on the next shuffle | Fresh fakes and state per test; replay the printed seed, fix the dependency |
| No assertions | False confidence | Always assert something meaningful |
| Testing trivial code | Wasted effort | Focus on logic, edge cases, boundaries |
| Slow tests | Reduced feedback | Move integration concerns to integration tests |
