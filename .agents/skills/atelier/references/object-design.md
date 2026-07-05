# Object-Oriented Design (class-free edition)

> **Note on examples.** Port and use-case signatures in this file are sometimes elided to `Promise<T>` (or throw on business failure) for brevity where error handling is not the lesson. In real code every IO port returns `Promise<Result<T, PortError>>` and every use-case returns `Promise<Result<Summary, StepError>>` — hard rule 16, see `references/result-type.md`.

## Responsibility-Driven Design

Modules are defined by their responsibilities, not their data.

### Finding modules

Start with:
1. Nouns in requirements | candidate modules (record types + operation functions).
2. Verbs | candidate functions.
3. Domain concepts | candidate value objects.

### Finding responsibilities

Each module answers:
- What does this module **know**?
- What does this module **do**?
- What does this module **decide**?

### Module stereotypes

Fit every module into one (or at most two) stereotype:

| Stereotype | Purpose | Example |
|:---|:---|:---|
| Information holder | Holds data, minimal behaviour | `User`, `Product`, `Address` |
| Structurer | Maintains relationships | `OrderItems`, `UserGroup` |
| Service provider | Performs stateless work | `paymentProcessor`, `emailSender` |
| Coordinator | Orchestrates workflow | `orderFulfillment` |
| Controller | Makes decisions, delegates | `checkoutController` |
| Interfacer | Transforms between systems | `userApiAdapter`, `dbMapper` |

### The two questions

For every module, ask:
1. "What stereotype is this?" (and if applicable, which design pattern?)
2. "Is it doing too much?" (check object calisthenics)

If you cannot answer either cleanly, refactor.

---

## Tell, don't ask

Command the module to do work. Do not interrogate its data and do the work elsewhere.

```ts
// BAD - ask, then do
if (account.balance >= amount) {
  account = { ...account, balance: account.balance - amount };
}

// GOOD - tell
const result = withdraw(account, amount);
if (result.kind === 'success') account = result.account;
```

The module that owns the data owns the behaviour.

---

## Design by Contract

Every function has:

- **Preconditions** | what must be true BEFORE calling.
- **Postconditions** | what will be true AFTER.
- **Invariants** | what is ALWAYS true about the module's state.

```ts
// INVARIANT - balance is never negative
// PRECONDITION - amount > 0
// POSTCONDITION - balance decreased by amount OR error returned
export type WithdrawResult =
  | { readonly kind: 'success'; readonly account: Account }
  | { readonly kind: 'invalidAmount' }
  | { readonly kind: 'insufficientFunds' };

export const withdraw = (account: Account, amount: Money): WithdrawResult => {
  if (amount.amount <= 0) return { kind: 'invalidAmount' };
  if (lessThanMoney(account.balance, amount)) return { kind: 'insufficientFunds' };
  return { kind: 'success', account: { ...account, balance: subMoney(account.balance, amount) } };
};
```

---

## Composition over inheritance

We have no `class`, so we have no inheritance by keyword. Good. Inheritance was always a fragile mechanism (tight coupling to a parent, fragile base class problem, forced "is-a" relationships).

Instead of inheritance, compose through spread:

```ts
// BAD - pseudo-inheritance via class extends (also banned in this repo)
// class PremiumUser extends User { getDiscount() { return 20; } }

// GOOD - composition via record spread
export type User = { readonly id: UserId; readonly email: Email };

export type DiscountPolicy = { readonly calculate: () => number };

export type DiscountedUser = User & { readonly discount: DiscountPolicy };

export const premiumDiscount: DiscountPolicy = { calculate: () => 20 };
export const standardDiscount: DiscountPolicy = { calculate: () => 0 };
export const noDiscount: DiscountPolicy = { calculate: () => 0 };

// pluggable discount behaviour at construction time
const upgradeToPremium = (user: User): DiscountedUser => ({ ...user, discount: premiumDiscount });
```

Discount is now a pluggable dependency, not a parent class.

---

## Law of Demeter (Principle of Least Knowledge)

Only talk to your immediate friends. A function should only call:

1. Functions on its own module.
2. Functions on its parameters.
3. Functions on records it constructs.

```ts
// BAD - reaching through
const city = order.customer.address.city;

// GOOD - ask the immediate friend
const city = orderShippingCity(order);
```

This reduces coupling: a change to `Address` does not ripple through every caller.

---

## Encapsulation

Hide internal details, expose behaviour. In a class-free codebase, encapsulation comes from:

1. `readonly` fields on record types.
2. Unexported helper functions in the same module.
3. Factory functions that return only the public contract.

```ts
// BAD - exposed mutable internals
export type Order = {
  items: OrderItem[];   // callers can push()
  total: number;        // callers can corrupt
};

// GOOD - encapsulated
export type Order = {
  readonly id: OrderId;
  readonly items: OrderItems;
  readonly total: Money;
};

export const addItemToOrder = (order: Order, item: OrderItem): Order => {
  const next = addToOrderItems(order.items, item);
  return { ...order, items: next, total: orderItemsTotal(next) };
};
```

The record is immutable by convention; operations return a new record.

When you must hide state behind a contract (e.g. a repository with a connection pool), expose it through a factory function that returns only the contract type:

```ts
export type UserRepo = { save: (u: User) => Promise<void>; findById: (id: UserId) => Promise<User | null> };

export const createPostgresUserRepo = (connection: PoolConnection): UserRepo => {
  // connection is a closure variable, invisible to callers
  return {
    save: async (u) => { /* uses connection */ },
    findById: async (id) => { /* uses connection */ },
  };
};
```

Callers see only `UserRepo`. The pool is encapsulated.

---

## Polymorphism (via dispatch)

Replace conditionals with dispatch records or function-type parameters.

```ts
// BAD - type checking
export const calculateShipping = (method: string, value: number): number => {
  if (method === 'standard') return value < 50 ? 5 : 0;
  if (method === 'express') return 15;
  if (method === 'overnight') return 25;
  throw new Error('unknown method');
};

// GOOD - polymorphism via function-type contract
export type ShippingMethod = { calculateCost: (orderValue: number) => number };

export const standardShipping: ShippingMethod = { calculateCost: (v) => (v < 50 ? 5 : 0) };
export const expressShipping: ShippingMethod = { calculateCost: () => 15 };
export const overnightShipping: ShippingMethod = { calculateCost: () => 25 };

export const calculateShipping = (method: ShippingMethod, value: number): number => method.calculateCost(value);
```

Or, when the set of variants is small and closed, use a discriminated union + exhaustive dispatch:

```ts
export type ShippingMethodKind = 'standard' | 'express' | 'overnight';

const shippingRates: Record<ShippingMethodKind, (value: number) => number> = {
  standard: (v) => (v < 50 ? 5 : 0),
  express: () => 15,
  overnight: () => 25,
};

export const calculateShipping = (kind: ShippingMethodKind, value: number): number => shippingRates[kind](value);
```

Both approaches avoid the growing-if-else smell.

---

## Value objects vs entities

### Value objects

Defined by attributes. No identity. Immutable. Compared by value. Examples: `Money`, `Email`, `Address`, `DateRange`.

```ts
export type Money = { readonly amount: number; readonly currency: string };

export const money = (amount: number, currency: string): Money => {
  if (!Number.isFinite(amount)) throw new Error('invalid Money.amount');
  return { amount, currency };
};

export const moneyEquals = (a: Money, b: Money): boolean =>
  a.amount === b.amount && a.currency === b.currency;

export const addMoney = (a: Money, b: Money): Money => {
  if (a.currency !== b.currency) throw new Error('CurrencyMismatch');
  return money(a.amount + b.amount, a.currency);
};
```

### Entities

Have identity that survives attribute changes. Mutable through transformation functions that return new records. Compared by identity.

```ts
export type User = {
  readonly id: UserId;
  readonly email: Email;
  readonly name: Name;
};

export const changeUserEmail = (user: User, newEmail: Email): User => ({ ...user, email: newEmail });

export const userEquals = (a: User, b: User): boolean => a.id === b.id;
```

The `id` is always the identity field. Two users with the same id are the same user, even if other fields differ.

---

## Aggregates

A cluster of records treated as a single unit for data changes.

- One record is the aggregate root (entry point).
- External code only references the root.
- Root transformations enforce invariants for the entire cluster.

```ts
// Order is the aggregate root
export type Order = {
  readonly id: OrderId;
  readonly items: OrderItems;
  readonly customerId: CustomerId;
};

const MAX_ORDER_VALUE = money(10000, 'EUR');

// All access through the root
export const addItemToOrder = (order: Order, product: Product, quantity: number): Order => {
  const item = orderItem(product, quantity);
  const nextItems = addToOrderItems(order.items, item);
  const next: Order = { ...order, items: nextItems };
  validateOrderInvariants(next);
  return next;
};

export const removeItemFromOrder = (order: Order, itemId: ItemId): Order => {
  const nextItems = removeFromOrderItems(order.items, itemId);
  return { ...order, items: nextItems };
};

const validateOrderInvariants = (order: Order): void => {
  if (greaterThanMoney(orderItemsTotal(order.items), MAX_ORDER_VALUE)) {
    throw new Error('OrderTotalExceeded');
  }
};

// BAD - bypassing the root, invariants not checked
// items.push(item);

// GOOD - through the root, invariants enforced
// const next = addItemToOrder(order, product, 2);
```

The root function (`addItemToOrder`) is the only public path to mutation. Anything inside `order.items` is an implementation detail.
