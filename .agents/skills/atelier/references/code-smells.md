# Code Smells and Anti-Patterns (class-free edition)

> **Note on examples.** Port and use-case signatures in this file are sometimes elided to `Promise<T>` (or throw on business failure) for brevity where error handling is not the lesson. In real code every IO port returns `Promise<Result<T, PortError>>` and every use-case returns `Promise<Result<Summary, StepError>>` — hard rule 16, see `references/result-type.md`.

## What are code smells?

Indicators that something may be wrong. Not bugs, but design problems that make code hard to understand, change, or test.

## The five categories

### 1. Bloaters - code that has grown too large

| Smell | Symptom | Refactoring |
|:---|:---|:---|
| Long function | > 10 lines | Extract function |
| Large module | > 50 lines, multiple responsibilities | Extract module (SRP) |
| Long parameter list | > 3 parameters | Introduce parameter record |
| Data clumps | Same group of parameters appear together | Extract record type |
| Primitive obsession | Raw strings and numbers for domain concepts | Wrap in branded types |

### 2. OO-abusers in a class-free codebase

Many classical OO smells do not exist here because we do not have classes. But the underlying problems still occur:

| Smell | Symptom in our style | Refactoring |
|:---|:---|:---|
| Switch statements | Large `if/else` chain or `switch` on a kind field | Replace with dispatch record or strategy contract |
| Parallel modules | Adding a new variant requires updating several files | Move related variants to the same module |
| Type tagging abuse | Checking kind fields in many places | Consolidate dispatch in one place |

### 3. Change preventers

| Smell | Symptom | Refactoring |
|:---|:---|:---|
| Divergent change | One module changed for many unrelated reasons | Split modules (SRP) |
| Shotgun surgery | One change touches many modules | Move related code together |

### 4. Dispensables

| Smell | Symptom | Refactoring |
|:---|:---|:---|
| Comments | Explaining bad code instead of rewriting | Rename, extract function |
| Duplicate code | Copy-paste | Extract function after Rule of Three |
| Dead code | Unreachable or unused | Delete |
| Speculative generality | Abstractions "just in case" | Delete (YAGNI) |
| Lazy module | A module that only wraps one trivial function | Inline it |

### 5. Couplers

| Smell | Symptom | Refactoring |
|:---|:---|:---|
| Feature envy | A function uses another module's data more than its own | Move the function to the envied module |
| Inappropriate intimacy | Two modules know too much about each other's internals | Extract a boundary contract |
| Message chains | `a.b.c.d` | Hide the delegation behind a friend function |
| Middle man | A module that only forwards calls | Inline it |

---

## The seven most common smells, in our style

### 1. Long function

**Symptom.** A function doing multiple things.

```ts
// SMELL
export const processOrder = async (order: Order): Promise<void> => {
  if (order.items.length === 0) throw new Error('empty');
  if (!order.customer) throw new Error('no customer');

  let total = 0;
  for (const item of order.items) {
    total += item.price * item.quantity;
    if (item.discount) total -= item.discount;
  }

  const taxRate = getTaxRate(order.customer.state);
  total = total * (1 + taxRate);

  await db.insert({ ...order, total });
  await mailer.send(order.customer.email, 'Order confirmed');
};

// REFACTORED
export const processOrder = async (order: Order, deps: ProcessOrderDeps): Promise<void> => {
  validateOrder(order);
  const total = calculateOrderTotal(order);
  await deps.repo.save(order, total);
  await deps.notifier.notifyConfirmation(order);
};
```

### 2. Large module

**Symptom.** A module with many responsibilities.

```ts
// SMELL - god module
// src/domain/user.ts
export const createUser = () => { /* ... */ };
export const login = () => { /* ... */ };
export const logout = () => { /* ... */ };
export const resetPassword = () => { /* ... */ };
export const setTheme = () => { /* ... */ };
export const setLanguage = () => { /* ... */ };
export const sendEmail = () => { /* ... */ };
export const sendSms = () => { /* ... */ };
export const charge = () => { /* ... */ };
export const refund = () => { /* ... */ };

// REFACTORED - split by responsibility, fitted to the Clean Architecture layout
// src/domain/user.ts                    - createUser, updateUser
// src/use-cases/auth/login.ts           - login, logout, resetPassword (with port deps)
// src/domain/preferences.ts             - setTheme, setLanguage
// src/use-cases/notifications/send.ts   - sendEmail, sendSms (port deps)
// src/use-cases/billing/charge.ts       - charge, refund (port deps)
```

### 3. Feature envy

**Symptom.** A function uses another module's data more than its own.

```ts
// SMELL - shipping logic lives in Order but uses only Customer data
// src/orders/order.ts
export const calculateShipping = (customer: Customer): Money => {
  if (customer.country === 'US') {
    if (customer.state === 'CA') return money(10, 'USD');
    return money(15, 'USD');
  }
  return money(25, 'USD');
};

// REFACTORED - move to customer module
// src/customers/customer-shipping.ts
export const customerShippingCost = (customer: Customer): Money => {
  if (customer.country === 'US') {
    if (customer.state === 'CA') return money(10, 'USD');
    return money(15, 'USD');
  }
  return money(25, 'USD');
};

// src/orders/order.ts - now just asks the friend
export const orderShippingCost = (order: Order): Money => customerShippingCost(order.customer);
```

### 4. Primitive obsession

**Symptom.** Raw strings and numbers for domain concepts.

```ts
// SMELL
export const createUser = (email: string, age: number, zipCode: string): User => {
  if (!email.includes('@')) throw new Error();
  if (age < 0) throw new Error();
  return { email, age, zipCode };
};

// REFACTORED - branded types catch invalid data at construction time
export type Email = string & { readonly __brand: 'Email' };
export const email = (value: string): Email => {
  if (!value.includes('@')) throw new Error('invalid Email');
  return value as Email;
};

export type Age = number & { readonly __brand: 'Age' };
export const age = (value: number): Age => {
  if (value < 0 || value > 150) throw new Error('invalid Age');
  return value as Age;
};

export type ZipCode = string & { readonly __brand: 'ZipCode' };
export const zipCode = (value: string): ZipCode => {
  if (!/^\d{5}(-\d{4})?$/.test(value)) throw new Error('invalid ZipCode');
  return value as ZipCode;
};

export const createUser = (e: Email, a: Age, z: ZipCode): User => ({ email: e, age: a, zipCode: z });
```

### 5. Switch statements / growing if-chains

**Symptom.** Type-checking on a kind field, repeated across the codebase.

```ts
// SMELL
export const getArea = (shape: Shape): number => {
  if (shape.kind === 'circle') return Math.PI * shape.radius ** 2;
  if (shape.kind === 'rectangle') return shape.width * shape.height;
  if (shape.kind === 'triangle') return 0.5 * shape.base * shape.height;
  throw new Error('unknown');
};

export const getPerimeter = (shape: Shape): number => {
  if (shape.kind === 'circle') return 2 * Math.PI * shape.radius;
  // ... same switch
  throw new Error('unknown');
};

// REFACTORED - dispatch record, add a new shape in one place
export type Shape = { area: () => number; perimeter: () => number };

export const circle = (radius: number): Shape => ({
  area: () => Math.PI * radius ** 2,
  perimeter: () => 2 * Math.PI * radius,
});

export const rectangle = (width: number, height: number): Shape => ({
  area: () => width * height,
  perimeter: () => 2 * (width + height),
});

export const triangle = (base: number, height: number, sides: readonly [number, number, number]): Shape => ({
  area: () => 0.5 * base * height,
  perimeter: () => sides[0] + sides[1] + sides[2],
});

// callers
export const getArea = (shape: Shape): number => shape.area();
export const getPerimeter = (shape: Shape): number => shape.perimeter();
```

### 6. Inappropriate intimacy

**Symptom.** Modules know too much about each other's internals.

```ts
// SMELL - order reaches into inventory internals
export const processOrder = (order: Order, inventory: Inventory): void => {
  for (const item of order.items) {
    const stock = inventory.stockLevels[item.sku];
    if (stock.quantity < item.quantity) throw new Error('out of stock');
    inventory.stockLevels[item.sku].quantity -= item.quantity;
  }
};

// REFACTORED - tell, do not ask
export type ReserveResult =
  | { readonly kind: 'success' }
  | { readonly kind: 'outOfStock'; readonly item: OrderItem };

export type Inventory = {
  reserve: (items: readonly OrderItem[]) => ReserveResult;
};

export const processOrder = (order: Order, inventory: Inventory): void => {
  const result = inventory.reserve(order.items);
  if (result.kind === 'outOfStock') throw new Error(`out of stock: ${result.item.sku}`);
};
```

### 7. Speculative generality

**Symptom.** Abstractions for hypothetical needs.

```ts
// SMELL - over-engineered for imagined future requirements
export type PaymentProcessor = {
  process: () => void;
  rollback: () => void;
  audit: () => void;
  generateReport: () => void;
  scheduleRecurring: () => void;
};

export const stripeProcessor: PaymentProcessor = {
  process: () => { /* real code */ },
  rollback: () => { throw new Error('not implemented'); },
  audit: () => { throw new Error('not implemented'); },
  generateReport: () => { throw new Error('not implemented'); },
  scheduleRecurring: () => { throw new Error('not implemented'); },
};

// REFACTORED - YAGNI
export type PaymentProcessor = { process: () => void };

export const stripeProcessor: PaymentProcessor = {
  process: () => { /* real code */ },
};

// Add rollback when the first real use case appears
```

---

## Prevention strategies

1. Follow object calisthenics. The rules prevent most smells.
2. Practice TDD. Tests reveal design problems early.
3. Review in pairs. Fresh eyes catch smells.
4. Refactor continuously. Do not let smells accumulate.
5. Apply SOLID. Prevents structural smells.
6. Use static analysis (ESLint + `typescript-eslint`). Catches many common issues automatically.

---

## When you find a smell

1. Confirm it is a problem. Not every smell needs fixing right now.
2. Ensure test coverage before refactoring.
3. Refactor in small steps. Keep tests green at each step.
4. Commit frequently. Easy to revert.
