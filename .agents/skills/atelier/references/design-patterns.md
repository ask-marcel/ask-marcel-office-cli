# Design Patterns (class-free edition)

Reusable solutions to common design problems. A shared vocabulary for discussing design. This file translates the classic GoF patterns into modules of arrow functions and typed records.

> **Note on examples.** Port and use-case signatures in this file are sometimes elided to `Promise<T>` (or throw on business failure) for brevity where error handling is not the lesson. In real code every IO port returns `Promise<Result<T, PortError>>` and every use-case returns `Promise<Result<Summary, StepError>>` — hard rule 16, see `references/result-type.md`.

## Warning first

Do not force patterns. Let them emerge from refactoring. Patterns solve problems you have, not problems you might have.

Use a pattern when:
1. You recognise the problem.
2. The pattern fits without bending.
3. It simplifies, not complicates.
4. The team understands it.

---

## Creational patterns

### Singleton

**Purpose.** Ensure only one instance exists.

**When to use.** Side-effect-free constants and frozen configuration records. Often overused. Dependency injection is usually better.

```ts
// Just export a module-level const. That is your singleton.
// src/domain/retry-policy.ts
export const retryPolicy = Object.freeze({ maxAttempts: 3, backoffMs: 200 });
```

No ceremony needed. ESM modules are singletons by default.

Stateful or IO-performing singletons (loggers, clients, pools) are NOT expressed this way — they are factories in `src/infra/**` injected at composition (hard rule 4).

### Factory

**Purpose.** Create records without the caller specifying the exact shape or variant.

**When to use.** Creation logic is complex or varies by type.

```ts
export type Notification = { send: (message: string) => Promise<void> };

export const emailNotification: Notification = {
  send: async (message) => { /* ... */ },
};
export const smsNotification: Notification = {
  send: async (message) => { /* ... */ },
};
export const pushNotification: Notification = {
  send: async (message) => { /* ... */ },
};

export type NotificationKind = 'email' | 'sms' | 'push';

export const createNotification = (kind: NotificationKind): Notification => {
  if (kind === 'email') return emailNotification;
  if (kind === 'sms') return smsNotification;
  return pushNotification;
};
```

Or with a dispatch record:

```ts
const notifications: Record<NotificationKind, Notification> = {
  email: emailNotification,
  sms: smsNotification,
  push: pushNotification,
};

export const createNotification = (kind: NotificationKind): Notification => notifications[kind];
```

### Builder

**Purpose.** Construct complex records step by step.

**When to use.** Records with many optional fields, test data creation.

```ts
export type User = {
  readonly name: Name;
  readonly email: Email;
  readonly age?: Age;
};

export type UserConfig = Partial<Omit<User, 'name' | 'email'>> & Pick<User, 'name' | 'email'>;

export const buildUser = (config: UserConfig): User => ({
  name: config.name,
  email: config.email,
  age: config.age,
});

// usage
const alice = buildUser({ name: name('Alice'), email: email('alice@example.com') });
const bob = buildUser({ name: name('Bob'), email: email('bob@example.com'), age: age(30) });
```

For test builders, see `references/testing.md` (section: Test builders).

### Prototype

**Purpose.** Create new records by cloning existing ones with variations.

Since our records are immutable, "clone with variations" is just spread:

```ts
export type Document = {
  readonly title: string;
  readonly content: string;
  readonly metadata: DocumentMetadata;
};

export const cloneDocument = (doc: Document, overrides: Partial<Document> = {}): Document => ({
  ...doc,
  ...overrides,
  metadata: { ...doc.metadata, ...overrides.metadata },
});
```

---

## Structural patterns

### Adapter

**Purpose.** Make an incompatible contract fit the contract we depend on.

**When to use.** Integrating a third-party library or legacy module.

```ts
// Third-party module with a different contract
type OldPaymentAPI = { makePayment: (cents: number) => boolean };

// Our contract
export type PaymentGateway = {
  charge: (amount: Money) => Promise<ChargeResult>;
};

// Adapter
export const createOldPaymentAdapter = (oldAPI: OldPaymentAPI): PaymentGateway => ({
  charge: async (amount) => {
    const cents = Math.round(amount.amount * 100);
    const success = oldAPI.makePayment(cents);
    return success ? chargeSuccess() : chargeFailed();
  },
});
```

### Decorator

**Purpose.** Add behaviour to a record dynamically.

**When to use.** Adding features without modifying existing code. Composing orthogonal behaviours.

```ts
export type Notifier = { send: (message: string) => Promise<void> };

export const emailNotifier: Notifier = {
  send: async (message) => { /* ... */ },
};

// Decorators are higher-order functions
export const withSms = (wrapped: Notifier): Notifier => ({
  send: async (message) => {
    await wrapped.send(message);
    await sendSms(message);
  },
});

export const withSlack = (wrapped: Notifier): Notifier => ({
  send: async (message) => {
    await wrapped.send(message);
    await postToSlack(message);
  },
});

// compose
const notifier = withSlack(withSms(emailNotifier));
await notifier.send('Alert!'); // sends to all three
```

### Proxy

**Purpose.** Control access to a resource.

**When to use.** Lazy loading, access control, logging, caching.

```ts
export type Image = { display: () => void };

export const createRealImage = (filename: string): Image => {
  loadFromDisk(filename); // expensive
  return { display: () => renderToScreen(filename) };
};

// Lazy-loading proxy
export const createImageProxy = (filename: string): Image => {
  let real: Image | null = null;
  return {
    display: () => {
      if (real === null) real = createRealImage(filename);
      real.display();
    },
  };
};
```

### Composite

**Purpose.** Treat individual records and compositions uniformly.

**When to use.** Tree structures, hierarchies (files/folders, UI components).

```ts
export type Component =
  | { readonly kind: 'product'; readonly price: Money }
  | { readonly kind: 'box'; readonly children: readonly Component[] };

export const componentPrice = (c: Component): Money => {
  if (c.kind === 'product') return c.price;
  return c.children.reduce((sum, child) => addMoney(sum, componentPrice(child)), money(0, 'EUR'));
};

// usage
const smallBox: Component = {
  kind: 'box',
  children: [
    { kind: 'product', price: money(10, 'EUR') },
    { kind: 'product', price: money(20, 'EUR') },
  ],
};

const bigBox: Component = {
  kind: 'box',
  children: [smallBox, { kind: 'product', price: money(50, 'EUR') }],
};

// componentPrice(bigBox) -> 80 EUR
```

---

## Behavioural patterns

### Strategy

**Purpose.** Define a family of algorithms, make them interchangeable.

**When to use.** Multiple ways to do the same thing, switchable at runtime.

```ts
export type PricingStrategy = { calculate: (basePrice: Money) => Money };

export const regularPricing: PricingStrategy = {
  calculate: (basePrice) => basePrice,
};

export const premiumDiscount: PricingStrategy = {
  calculate: (basePrice) => money(basePrice.amount * 0.8, basePrice.currency),
};

export const blackFriday: PricingStrategy = {
  calculate: (basePrice) => money(basePrice.amount * 0.5, basePrice.currency),
};

export const cartTotal = (items: readonly Item[], pricing: PricingStrategy): Money => {
  const base = items.reduce((sum, i) => addMoney(sum, i.price), money(0, 'EUR'));
  return pricing.calculate(base);
};
```

### Observer

**Purpose.** Notify multiple listeners about state changes.

**When to use.** Event systems, pub/sub, reactive updates.

```ts
export type Observer<T> = (event: T) => void;

export type Emitter<T> = {
  subscribe: (observer: Observer<T>) => () => void; // returns unsubscribe
  emit: (event: T) => void;
};

export const createEmitter = <T>(): Emitter<T> => {
  let observers: Observer<T>[] = [];
  return {
    subscribe: (observer) => {
      observers.push(observer);
      return (): void => {
        observers = observers.filter((o) => o !== observer);
      };
    },
    emit: (event) => observers.forEach((o) => o(event)),
  };
};

// usage
type OrderEvent = { kind: 'placed'; orderId: OrderId };
const orderEvents = createEmitter<OrderEvent>();

const placed: OrderId[] = [];
const unsubscribe = orderEvents.subscribe((event) => {
  if (event.kind === 'placed') placed.push(event.orderId);
});

orderEvents.emit({ kind: 'placed', orderId: orderId('ord-1') });
unsubscribe();
```

### Template Method

**Purpose.** Define an algorithm skeleton, let callers plug in the varying steps.

**When to use.** Common algorithm with varying steps.

```ts
// The skeleton as a higher-order function
export type DataExporterSteps = {
  format: (data: readonly Data[]) => string;
  write: (content: string) => Promise<void>;
};

export const runExport = async (data: readonly Data[], steps: DataExporterSteps): Promise<void> => {
  validateExportData(data);
  const formatted = steps.format(data);
  await steps.write(formatted);
  await notifyExportComplete();
};

// Variant steps
export const csvExporterSteps: DataExporterSteps = {
  format: (data) => data.map(dataToCsvRow).join('\n'),
  write: async (content) => Bun.write('export.csv', content).then(() => undefined),
};

export const jsonExporterSteps: DataExporterSteps = {
  format: (data) => JSON.stringify(data),
  write: async (content) => Bun.write('export.json', content).then(() => undefined),
};

// usage
await runExport(rows, csvExporterSteps);
await runExport(rows, jsonExporterSteps);
```

### Command

**Purpose.** Encapsulate a request as a record.

**When to use.** Undo/redo, queuing, logging actions.

```ts
export type Command = {
  execute: () => void;
  undo: () => void;
};

// addToCart/removeFromCart are immutable (they return new carts), so the
// command closes over a mutable holder; execute/undo swap the current cart.
export const addItemCommand = (cart: Cart, item: Item): Command & { getCart: () => Cart } => {
  let current = cart;
  return {
    execute: () => {
      current = addToCart(current, item);
    },
    undo: () => {
      current = removeFromCart(current, item);
    },
    getCart: () => current,
  };
};

export type CommandHistory = {
  execute: (command: Command) => void;
  undo: () => void;
};

export const createCommandHistory = (): CommandHistory => {
  const history: Command[] = [];
  return {
    execute: (command) => {
      command.execute();
      history.push(command);
    },
    undo: () => {
      const last = history.pop();
      if (last !== undefined) last.undo();
    },
  };
};
```

---

## Pattern awareness (the four-dimensional lens)

When reading unfamiliar code, ask:

1. What problem does this solve? (creational, structural, behavioural)
2. What scope? (record-level, module-level, system-level)
3. When is it applied? (compile-time, runtime)
4. How coupled? (tight, loose)

This lets you recognise patterns even when the code has no class in sight.

---

## Anti-patterns to avoid

| Anti-pattern | Problem | Solution |
|:---|:---|:---|
| God module | One file does everything | Split by responsibility (SRP) |
| Spaghetti | Tangled, no structure | Refactor to vertical slices and layers |
| Golden hammer | Using one pattern for everything | Match pattern to the problem |
| Premature optimisation | Optimising before measuring | YAGNI, profile first |
| Copy-paste | Duplication everywhere | Extract on Rule of Three |
| Pattern hunting | Applying patterns to look clever | Let patterns emerge from refactoring |
