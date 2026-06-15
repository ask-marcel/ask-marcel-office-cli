# SOLID Principles (class-free edition)

SOLID helps structure software to be flexible, maintainable, and testable. These principles reduce coupling and increase cohesion. In a codebase with no classes or interfaces, each principle still applies. We express it through function-type aliases and modules of arrow functions.

> **Note on examples.** Port and use-case signatures in this file are sometimes elided to `Promise<T>` (or throw on business failure) for brevity where error handling is not the lesson. In real code every IO port returns `Promise<Result<T, PortError>>` and every use-case returns `Promise<Result<Summary, StepError>>` — hard rule 16, see `references/result-type.md`.

## S | Single Responsibility Principle

> "A module should have one, and only one, reason to change."

### Problem it solves

God modules that do everything: hard to test, hard to change, hard to understand.

### How to apply

Each module handles one responsibility. If describing what a module does requires "and", split it.

```ts
// BAD - three responsibilities in one module
export const processOrder = async (order: Order): Promise<void> => {
  if (order.items.length === 0) throw new Error('empty');           // validation
  const total = order.items.reduce((s, i) => s + i.price, 0);       // calculation
  await db.query('INSERT INTO orders ...');                         // persistence
  await mailer.send(order.customer.email, 'Confirmed');             // notification
};

// GOOD - one responsibility per module, composed at the use-case
// src/orders/services/validate-order.ts
export const validateOrder = (order: Order): void => {
  if (order.items.length === 0) throw new Error('empty order');
};

// src/orders/services/calculate-total.ts
export const calculateOrderTotal = (order: Order): Money =>
  order.items.reduce((sum, i) => addMoney(sum, i.price), money(0, 'EUR'));

// src/orders/services/save-order.ts
export type OrderRepo = { save: (order: Order, total: Money) => Promise<void> };

// src/orders/services/notify-customer.ts
export type OrderNotifier = { notifyConfirmation: (order: Order) => Promise<void> };

// src/orders/use-cases/place-order.ts (composition)
export const placeOrder = async (
  order: Order,
  repo: OrderRepo,
  notifier: OrderNotifier
): Promise<void> => {
  validateOrder(order);
  const total = calculateOrderTotal(order);
  await repo.save(order, total);
  await notifier.notifyConfirmation(order);
};
```

### Detection questions

- Does this module have multiple reasons to change?
- Can I describe it without using "and"?
- Would different stakeholders request changes to different parts?

---

## O | Open/Closed Principle

> "Software entities should be open for extension but closed for modification."

### Problem it solves

Modifying existing, tested code every time requirements change. Risk of breaking working features.

### How to apply

Design function-type contracts that allow new behaviour through new records or new modules, not edits to existing ones.

```ts
// BAD - must modify to add new shipping
type ShippingKind = 'standard' | 'express';

export const calculateShipping = (kind: ShippingKind, value: number): number => {
  if (kind === 'standard') return value < 50 ? 5 : 0;
  if (kind === 'express') return 15;
  throw new Error('unknown kind');
};

// GOOD - open for extension
export type ShippingMethod = { calculateCost: (orderValue: number) => number };

export const standardShipping: ShippingMethod = {
  calculateCost: (v) => (v < 50 ? 5 : 0),
};

export const expressShipping: ShippingMethod = {
  calculateCost: () => 15,
};

// Add new shipping by exporting a new const, not editing existing
export const sameDayShipping: ShippingMethod = {
  calculateCost: () => 25,
};
```

### Architectural insight

OCP at architecture level means: design the codebase so new features arrive as new modules, not as edits to existing ones. Expect diffs that only `A` files, never `M` them.

---

## L | Liskov Substitution Principle

> "Subtypes must be substitutable for their base types without altering program correctness."

### Problem it solves

Implementations that break caller expectations, forcing type checks and special cases.

### How to apply

Every implementation of a function-type contract must honour that contract. If the contract says "never returns negative", no implementation may return negative.

```ts
// The contract
export type DiscountPolicy = {
  getDiscount: (orderValue: number) => number; // contract: >= 0
};

// BAD - violates the contract
export const weirdDiscount: DiscountPolicy = {
  getDiscount: () => -5, // increases cost, breaks caller trust
};

// GOOD - factory enforces the contract at construction time
export const fixedDiscount = (amount: number): DiscountPolicy => {
  if (amount < 0) throw new Error('discount must be non-negative');
  return { getDiscount: () => amount };
};

export const percentDiscount = (percent: number): DiscountPolicy => {
  if (percent < 0 || percent > 100) throw new Error('percent out of range');
  return { getDiscount: (value) => value * (percent / 100) };
};
```

### Key insight

This is why you can swap `inMemoryUserRepo` for `postgresUserRepo` for a test-only fake. All implementations honour the `UserRepo` contract, so callers cannot tell them apart.

---

## I | Interface Segregation Principle

> "Clients should not be forced to depend on methods they do not use."

### Problem it solves

Fat contracts that force partial implementations, empty methods, or `throw new Error('not supported')`.

### How to apply

Split large function-type aliases into smaller, cohesive ones. Clients depend only on what they need.

```ts
// BAD - fat contract
export type WarehouseDevice = {
  printLabel: (orderId: OrderId) => void;
  scanBarcode: () => string;
  packageItem: (orderId: OrderId) => void;
};

// A printer that only prints has to implement scan and package too
export const basicPrinter: WarehouseDevice = {
  printLabel: (orderId) => { /* works */ },
  scanBarcode: () => { throw new Error('not supported'); }, // forced!
  packageItem: () => { throw new Error('not supported'); },
};

// GOOD - segregated contracts
export type LabelPrinter = { printLabel: (orderId: OrderId) => void };
export type BarcodeScanner = { scanBarcode: () => string };
export type ItemPackager = { packageItem: (orderId: OrderId) => void };

export const basicPrinter: LabelPrinter = {
  printLabel: (orderId) => { /* only what it does */ },
};

export const fullStation: LabelPrinter & BarcodeScanner & ItemPackager = {
  printLabel: (orderId) => { /* ... */ },
  scanBarcode: () => { /* ... */ },
  packageItem: (orderId) => { /* ... */ },
};
```

### Detection

If you see `throw new Error('not implemented')` or empty method bodies, the contract is too fat. Split it.

---

## D | Dependency Inversion Principle

> "High-level modules should not depend on low-level modules. Both should depend on abstractions."

### Problem it solves

Tight coupling to specific implementations (databases, APIs, frameworks). Hard to test, hard to swap.

### How to apply

Depend on function-type aliases, inject implementations through factory functions.

```ts
// BAD - direct dependency on a concrete module
import { sendGridSend } from '../infra/sendgrid';

export const confirmOrder = async (to: Email): Promise<void> => {
  await sendGridSend(to, 'Order confirmed'); // locked in
};

// GOOD - depend on a function-type contract
export type EmailSender = { send: (to: Email, message: string) => Promise<void> };

export type ConfirmOrder = (to: Email) => Promise<void>;
export const createConfirmOrder = (sender: EmailSender): ConfirmOrder =>
  async (to) => {
    await sender.send(to, 'Order confirmed');
  };

// Wire any implementation at composition time
// src/orders/infra/sendgrid-sender.ts
export const sendGridSender: EmailSender = {
  send: async (to, message) => { /* real SendGrid call */ },
};

// src/orders/infra/ses-sender.ts
export const sesSender: EmailSender = {
  send: async (to, message) => { /* real SES call */ },
};

// src/orders/infra/in-memory-sender.ts (for tests)
export const createInMemorySender = (): EmailSender & { sent: { to: Email; message: string }[] } => {
  const sent: { to: Email; message: string }[] = [];
  return {
    sent,
    send: async (to, message) => {
      sent.push({ to, message });
    },
  };
};

// composition
const confirm = createConfirmOrder(sendGridSender);
await confirm(email('alice@example.com'));
```

### The dependency rule

Source code dependencies point INWARD toward high-level policies (domain). They never point outward to infrastructure.

```
Infrastructure -> Application -> Domain
      outer          middle         inner

Dependencies flow: outer -> inner
Never:             inner -> outer
```

---

## Applying SOLID at architecture scale

These principles scale beyond modules:

| Principle | Architecture application |
|:---|:---|
| SRP | Each bounded context has one responsibility. |
| OCP | New features arrive as new modules, not edits to existing ones. |
| LSP | Swappable implementations (real repo, fake repo, in-memory repo) behave identically within the contract. |
| ISP | Thin function-type contracts at module boundaries. No fat "everything" contracts. |
| DIP | High-level policy modules know nothing about databases or frameworks. |

---

## Quick reference

| Principle | One-liner | Red flag |
|:---|:---|:---|
| SRP | One reason to change | "This module handles X and Y and Z" |
| OCP | Add, do not modify | `if/else` chains for types |
| LSP | Implementations are substitutable | Type-checking in calling code |
| ISP | Small, focused contracts | Empty or `throw` method bodies |
| DIP | Depend on contracts, not concretions | Importing a concrete implementation from domain code |
