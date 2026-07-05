# Test-Driven Development

> **Note on examples.** Port and use-case signatures in this file are sometimes elided to `Promise<T>` (or throw on business failure) for brevity where error handling is not the lesson. In real code every IO port returns `Promise<Result<T, PortError>>` and every use-case returns `Promise<Result<Summary, StepError>>` — hard rule 16, see `references/result-type.md`.

## The core loop

```
RED -> GREEN -> REFACTOR -> RED -> ...
```

### RED phase

Propose a failing test that describes the behaviour you want, and get the user's confirmation before writing it (SKILL.md hard rule 24 — tests are confirmation-gated; never create, change, or delete one silently). The test should:

- Use domain language, not technical jargon.
- Describe WHAT, not HOW.
- Be a concrete example, not an abstract statement.

```ts
// BAD - abstract
it('can add numbers', () => { /* ... */ });

// GOOD - concrete example
it('when adding 2 + 3, returns 5', () => { /* ... */ });
```

### GREEN phase

Write the simplest possible code to make the test pass. Two strategies:

**Fake It** - return a hardcoded value.
```ts
export const add = (a: number, b: number): number => 5;
```

**Obvious Implementation** - if you know the solution.
```ts
export const add = (a: number, b: number): number => a + b;
```

Prefer Fake It when learning or unsure. Let more tests drive the real implementation.

### REFACTOR phase

This is where design happens. Look for:

- Duplication (wait for Rule of Three).
- Functions longer than 10 lines to extract.
- Poor names to improve.
- Complex conditions to simplify.
- Raw primitives in domain positions to promote to branded types.

## The Three Laws of TDD

1. No production code unless it makes a failing test pass.
2. No more test code than sufficient to fail (compilation failures count).
3. No more production code than sufficient to pass the one failing test.

## The Rule of Three

Only extract duplication when you see it THREE times.

Why? A wrong abstraction is more expensive to undo than duplication is to tolerate.

```
Duplication #1 - leave it.
Duplication #2 - note it, leave it.
Duplication #3 - NOW extract it.
```

## Triangulation

Each new test sculpts the solution toward a general implementation.

Think of degrees of freedom: a car needs forward/back, left/right, rotation. Each test carves out one degree of freedom until the implementation handles all cases.

Example. You implement `isPalindrome`:

1. Test `'mom'` -> fake it: `return true`.
2. Test `'hello'` -> now the fake fails. Generalise: compare halves.
3. Test `''` -> edge case forces explicit handling.
4. Test `'racecar'` -> confirms the general case.

## Transformation Priority Premise

When going from RED to GREEN, prefer simpler transformations:

| Priority | Transformation |
|:---|:---|
| 1 | `{}` to `null` |
| 2 | `null` to constant |
| 3 | constant to variable |
| 4 | unconditional to conditional |
| 5 | scalar to collection |
| 6 | statement to recursion |
| 7 | value to mutated value |

Higher priority = simpler. Do not jump to complex transformations too early.

## Arrange-Act-Assert

Every test follows the same shape: ARRANGE the world (real domain, faked secondary ports), ACT by calling the primary port, ASSERT on the returned result or the fake's final state. The worked port-level example lives in `references/testing.md` (§ Arrange-Act-Assert).

## Writing tests backwards

When stuck, flip the order:

1. Write the ASSERT first. What do you want to verify?
2. Write the ACT. What action produces that result?
3. Write the ARRANGE. What setup is needed?

## Test naming principles

- Use behaviour-driven names with domain language.
- Provide concrete examples, not abstract statements.
- One example per test for easy debugging.
- Do not leak implementation details.

```ts
// BAD - technical, implementation-focused
it('should set the data property to 1', () => { /* ... */ });

// GOOD - behaviour-focused, domain language
it('should recognise "mom" as a palindrome', () => { /* ... */ });
```

## Outside-in classicist TDD

Inspired by Ian Cooper's talk *TDD, Where Did It All Go Wrong?* (and the classic Detroit/Chicago school that predates it).

The school we follow has three rules. Each one is a direct response to a pattern of test pain that the industry has learned the hard way.

### 1. The SUT is the primary port

Tests target the **primary port** — the use case, command handler, or application service at the hexagonal boundary. Never an individual entity, value object, or domain service.

```ts
// BAD - testing the entity directly
describe('Order entity', () => {
  it('getDiscount returns 20 when tier is premium', () => {
    const order = createOrder(orderId('ord-1'), 'premium');
    expect(getDiscount(order)).toBe(20);
  });
});

// GOOD - testing the primary port; the entity is used, not tested
describe('placeOrder use-case', () => {
  it('when a premium customer buys 100 EUR, the order total is 80 EUR', async () => {
    const orders = createInMemoryOrderRepo();
    const customer = customerId('c-1');
    await placeOrder(
      { customer, items: [{ sku: 'SKU-1', price: money(100, 'EUR') }] },
      { orders, customers: createInMemoryCustomerRepo({ [customer]: { tier: 'premium' } }) }
    );
    const saved = await orders.findByCustomer(customer);
    expect(saved[0].total).toEqual(money(80, 'EUR'));
  });
});
```

### 2. Domain collaborators are real; only secondary ports get fakes

| Kind | Role | Treatment in tests |
|:---|:---|:---|
| Entity | `Order`, `User`, `Subscription` | Real |
| Value object | `Money`, `Email`, `OrderId` | Real |
| Domain service | `pricingRules`, `discountPolicy` | Real |
| Aggregate root | `Order`, `Cart` | Real |
| Primary port | `placeOrder`, `registerUser`, `checkoutCart` | **The SUT** |
| Secondary port | `OrderRepo`, `EmailSender`, `Clock`, `TokenDecoder`, `PaymentGateway` | **Faked** (hand-written in-memory) |

The secondary ports are the ones that talk to the outside world — databases, HTTP, the clock, the filesystem, random sources. They are the only things that need a double. Everything else runs for real inside the test.

This is the single most important property of the school: **the domain can be refactored freely.** Rename an entity, split a domain service into two, merge three value objects, change the shape of an aggregate, extract a helper, inline a helper — tests keep passing because they describe behaviour at the port, not structure inside.

### 3. No mocks, ever

Never import from the `mock` namespace of `bun:test` — `mock()`, `mock.module()`, `.toHaveBeenCalled*`. The entire namespace is banned and enforced by `no-restricted-imports`. Write a fake (a working in-memory implementation of the secondary-port contract) and assert on its final state; for infra adapters wrapping external SDKs, expose the two-constructor pattern (`createX` + `createXFromApi`) instead. The philosophical reason belongs here: a mock verifies the *sequence of internal calls*, so the test breaks on every innocent refactor and stops proving behaviour; a fake verifies the *final state*, which is what the system is for. The banned/required code pair, the full five-point rationale, and the permitted test-double shapes (dummy, stub, fake, hand-written spy) live in `references/testing.md` (§ No `mock` from `bun:test`).

### What this buys

- **Freedom to refactor the domain.** Internal restructurings do not break tests. Tests describe the port's behaviour; the port's behaviour is what stays stable.
- **Tests that survive for years.** Business scenarios are stable; code that implements them changes constantly.
- **Tests that read as specifications.** Every test name is a complete business scenario. A new team member learns the product by reading test titles.
- **Design pressure on the right boundary.** When a test is hard to write, it is telling you the primary port's contract is wrong — not that the entity needs a helper method.

## TDD for a pure arrow-function module

Example. Build a `calculateDiscount` function.

**Test file** `src/pricing/calculate-discount.test.ts`:

```ts
import { describe, expect, it } from 'bun:test';
import { calculateDiscount } from './calculate-discount';
import { money } from '../money/money';

describe('calculateDiscount', () => {
  it('when standard customer buys 100 EUR, returns 0', () => {
    const subtotal = money(100, 'EUR');
    const result = calculateDiscount(subtotal, 'standard');
    expect(result.amount).toBe(0);
  });

  it('when premium customer buys 100 EUR, returns 20', () => {
    const subtotal = money(100, 'EUR');
    const result = calculateDiscount(subtotal, 'premium');
    expect(result.amount).toBe(20);
  });
});
```

**RED.** Test 1 fails (no `calculate-discount.ts` yet).

**GREEN.** Fake it:

```ts
import type { Money } from '../money/money';
import { money } from '../money/money';

type CustomerTier = 'standard' | 'premium';

export const calculateDiscount = (subtotal: Money, tier: CustomerTier): Money => money(0, subtotal.currency);
```

Test 1 passes. Test 2 fails. Generalise:

```ts
export const calculateDiscount = (subtotal: Money, tier: CustomerTier): Money => {
  if (tier === 'premium') return money(subtotal.amount * 0.2, subtotal.currency);
  return money(0, subtotal.currency);
};
```

**REFACTOR.** At this point: no duplication to extract, names are clear, function is 4 lines. Nothing to clean. Move on to the next test. When a third tier appears (`vip`), resist extracting until after the third `if` branch exists (Rule of Three). Then promote the logic to a dispatch record:

```ts
const tierRates: Record<CustomerTier, number> = {
  standard: 0,
  premium: 0.2,
  vip: 0.3,
};

export const calculateDiscount = (subtotal: Money, tier: CustomerTier): Money =>
  money(subtotal.amount * tierRates[tier], subtotal.currency);
```

This is what "design happens during refactor" looks like.

## Common mistakes

1. Writing code before tests. Violates the fundamental principle.
2. Writing too much test. Just enough to fail.
3. Writing too much code. Just enough to pass.
4. Skipping refactor. Design lives here.
5. Testing implementation. Test behaviour, not how it is done.
6. Abstract test names. Use concrete examples.
7. Extracting too early. Wait for Rule of Three.
8. Reaching for doubles too soon. Start with real collaborators. If a double is needed, write a fake — never a mock.
9. Asserting on multiple unrelated behaviours in one test. One behaviour per test.
