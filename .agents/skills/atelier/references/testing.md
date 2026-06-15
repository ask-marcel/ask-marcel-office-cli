# Testing Strategy

## The school: Outside-in classicist

The SUT of every unit test is a **primary port** — a use case, command handler, or application service at the hexagonal boundary. Inside the port, the full domain runs real: entities, value objects, domain services, aggregate roots. The only test doubles are **fakes** for secondary ports (repository, email sender, clock, token decoder, payment gateway, any adapter to the outside world).

Benefits:

- Refactoring the domain never breaks tests.
- Tests describe business scenarios, so they read as living documentation.
- The design pressure lands on the right boundary: when a test is hard to write, the port's contract is wrong, not the entity.

See `references/tdd.md` for the full treatment and Ian Cooper's context.

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
      { customer, items: [{ sku: 'SKU-1', price: money(100, 'EUR') }] },
      { orders, customers }
    );

    const [saved] = await orders.findByCustomer(customer);
    expect(saved.total).toEqual(money(80, 'EUR'));
  });
});
```

Notice what is **real**: the `placeOrder` use case, every domain function it calls, the `Money` value object, the `Order` entity, the pricing rules. What is **faked**: `orders` and `customers` — the two secondary ports.

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
    const order = buildOrder({ customer, total: money(80, 'EUR') });
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
    { customer, items: [{ sku: 'SKU-1', price: money(100, 'EUR') }] },
    { orders, customers }
  );

  // ASSERT - read state from the fake
  const [saved] = await orders.findByCustomer(customer);
  expect(saved.total).toEqual(money(80, 'EUR'));
});
```

### Writing AAA backwards

When stuck:
1. Assert first. What do you want to verify?
2. Act. What action produces that result?
3. Arrange. What setup is needed?

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

Three shapes are permitted: **dummy**, **stub**, **fake**. Hand-written spies (a fake that also records its inputs) are allowed when outcome assertions are not enough. Mocks from a mock library are banned — see the "No mocks" rule below.

### Dummy

A record passed but never used.

```ts
const dummyLogger = {} as Logger;
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
  };
};
```

Used in a test:

```ts
const sheets = createSheetsFake({
  tabs: { POST: [row] },
  errors: { appendOrUpdate: { kind: 'write-failed', message: 'quota' } },
});
```

Cast the error-map value to the port's discriminated-union shape with `as const` on the `kind` so TypeScript narrows to the right variant. See `references/result-type.md`.

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

`err(...)` from a batch use-case is reserved for prerequisites — the initial `sheets.readRows` fails, or credentials are missing. See `references/result-type.md` for the full rationale.

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

The entire `mock` namespace of `bun:test` is banned — `mock()`, `mock.module()`, `.toHaveBeenCalledWith`, `.toHaveBeenCalledTimes`. Enforced by `no-restricted-imports` in `eslint.config.js`:

```js
'no-restricted-imports': ['error', {
  paths: [{
    name: 'bun:test',
    importNames: ['mock'],
    message: 'Use fakes (in-memory implementations), hand-written spies, or the createXFromApi(api) two-constructor pattern for infra adapters. See references/testing.md.',
  }],
}],
```

```ts
// BANNED
import { mock } from 'bun:test';
const mockSave = mock(async (_user: User): Promise<void> => {});
expect(mockSave).toHaveBeenCalledWith(expectedUser);

// BANNED (module substitution is process-global and leaks across test files)
mock.module('googleapis', () => ({ google: { drive: () => fakeApi } }));

// REQUIRED — fake the port for use-case tests
const repo = createInMemoryUserRepo();
await placeOrder(order, { repo });
expect(await repo.count()).toBe(1);

// REQUIRED — pass the API slice for infra adapter tests (see "Testing infra adapters")
const api: DriveApi = { files: { copy: async () => ({ data: { id: 'X' } }), /* ... */ } };
const drive = createDriveFromApi(api);
```

Why the absolute ban:

- **`mock.module` is process-global, not file-scoped.** Once set in any test file, every subsequent file the runner loads sees the substitution. This silently corrupted an unrelated `sleep.test.ts` in production use. There is no per-file restore; the leak is a feature of Bun's module cache.
- **`mock()` leaks without `mock.restore()` discipline.** Easy to forget; leak detection is best-effort.
- **Mocks test call sequences, not outcomes.** A mock passes when the right method is called with the right arguments — even if the production code does nothing useful afterwards. A fake passes only when the final state is correct, which is what the system is actually for.
- **Mocks couple tests to implementation.** Rename a method, split a call into two, extract a helper: the mock expectations break even though behaviour is unchanged. The fake keeps passing because the observable state is the same.
- **Mocks hide design pressure.** If you need a mock to test something, the contract is probably too fat (Interface Segregation), or the adapter is missing its `createXFromApi(api)` factory. Fix the design; do not reach for a mock.

`installFetchMock` (see "Testing infra adapters") and per-file `globalThis.setTimeout` swaps are **not** `mock.module` — they swap a global within a lifecycle hook (`afterEach`, `afterAll`) that always restores. The scope is bounded to the test file, not the process.

---

## What goes where

### Unit tests — primary port as SUT (the default)

Most tests. The SUT is a use case, command handler, or application service. The domain runs real; secondary ports are faked.

```ts
describe('placeOrder', () => {
  it('when a premium customer buys a 100 EUR item, the order is saved with a 80 EUR total and a confirmation email is queued', async () => {
    const orders = createInMemoryOrderRepo();
    const emails = createEmailSpy();
    const customer = customerId('c-1');
    const customers = createInMemoryCustomerRepo({ [customer]: { tier: 'premium' } });

    await placeOrder(
      { customer, items: [{ sku: 'SKU-1', price: money(100, 'EUR') }] },
      { orders, customers, emails }
    );

    const [saved] = await orders.findByCustomer(customer);
    expect(saved.total).toEqual(money(80, 'EUR'));
    expect(emails.sentTo).toContain(email('c-1@example.com'));
  });
});
```

### Value-object / domain-service tests (the exception)

If a value object or a domain service has genuinely complex logic of its own — `Money.add` with currency rules, `PricingPolicy` with tier brackets, `DateRange.overlaps` — a handful of small direct tests is fine. They supplement the primary-port tests, they do not replace them. Keep them rare and only when the logic is non-trivial enough that discovering it through a use-case test would be confusing.

```ts
describe('Money.add', () => {
  it('adds two amounts with the same currency', () => {
    expect(addMoney(money(10, 'EUR'), money(20, 'EUR'))).toEqual(money(30, 'EUR'));
  });

  it('refuses to add different currencies', () => {
    expect(() => addMoney(money(10, 'EUR'), money(10, 'USD'))).toThrow('CurrencyMismatch');
  });
});
```

A rough signal: if you find yourself writing more direct value-object tests than primary-port tests, something is off. The use case is where the business value lives; that is where most tests should point.

### Branded types and `expect(...).toBe(raw)` — the test escape hatch

Bun's `expect(x).toBe(y)` matcher infers `y`'s type from `x`. When `x` has a branded type, `y` must be the same brand or TypeScript fails:

```ts
const tok = accessToken('eyJ...'); // accessToken: (s: string) => AccessToken
expect(tok).toBe('eyJ...');
//             ^^^^^^^^ Argument of type 'string' is not assignable to 'AccessToken'.
```

Three options. Use the third.

1. ❌ `as` the raw string — assertions are forbidden everywhere else, do not start in tests.
2. ❌ Run the value through the real factory in the assertion (`expect(tok).toBe(accessToken('eyJ...'))`) — works, but the factory may have side effects (logging, parsing) that you don't want in a hot test loop.
3. ✅ **Export an `xxxUnsafe(raw): X` helper next to the factory. Use it only in tests.**

```ts
// src/domain/access-token.ts
export type AccessToken = string & { readonly __brand: 'AccessToken' };

export const accessToken = (value: string): AccessToken => {
  if (value.length === 0) throw new Error('AccessToken: empty');
  return value as AccessToken;
};

// Test escape hatch — bypasses validation. Naming convention: <factory>Unsafe.
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

**Naming convention.** `<factoryName>Unsafe` — `accessTokenUnsafe`, `envVarUnsafe`, `userIdUnsafe`, `safeUrlUnsafe`. The `Unsafe` suffix tells the next reader (and the next grep) exactly what they're looking at: a brand cast without validation, for tests only.

**Boundary.** Production code must not import any `*Unsafe` helper. A simple lint rule (or a periodic grep) keeps it honest:

```js
// eslint.config.js
{
  files: ['src/**/!(*.test).ts'],
  rules: {
    'no-restricted-imports': ['error', {
      patterns: [{ group: ['*'], importNames: ['/Unsafe$/'], message: '*Unsafe helpers are test-only' }],
    }],
  },
}
```

(ESLint's `importNames` doesn't accept regex directly; in practice the rule is enforced by review or by a small custom rule. The convention is the load-bearing part.)

### Secondary-port integration tests

Integration tests prove the real adapter (Postgres, SendGrid, Redis) fulfils the contract its in-memory fake already satisfies. Run in a separate CI stage with real infrastructure.

```ts
describe('postgresOrderRepo', () => {
  let repo: OrderRepo;

  beforeAll(async () => {
    repo = createPostgresOrderRepo(testDb);
  });

  it('saves an order and finds it by customer', async () => {
    const order = buildOrder({ customer: customerId('c-1'), total: money(80, 'EUR') });
    await repo.save(order);
    const [found] = await repo.findByCustomer(customer('c-1'));
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
const withItems = buildOrder({ items: [item({ sku: 'ABC', price: money(100, 'EUR') })] });
```

---

## Common testing mistakes

| Mistake | Problem | Fix |
|:---|:---|:---|
| Testing implementation | Brittle tests | Test observable behaviour only |
| Using mocks | Tests prove call sequences instead of outcomes; break on refactor | Never use mocks — write a fake for the contract |
| Shared state between tests | Flaky tests | Isolate each test (fresh fakes per test) |
| No assertions | False confidence | Always assert something meaningful |
| Testing trivial code | Wasted effort | Focus on logic, edge cases, boundaries |
| Slow tests | Reduced feedback | Move integration concerns to integration tests |
