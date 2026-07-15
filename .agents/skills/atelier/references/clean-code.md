# Clean Code (class-free edition)

## What is clean code?

Code that is:
- Easy to understand (reveals intent clearly).
- Easy to change (modifications are localised).
- Easy to test (dependencies are injectable through function-type contracts).
- Simple (no unnecessary complexity).

## The human-centred approach

Code has three consumers:
1. Users | get their needs met.
2. Customers | make or save money.
3. Developers | must maintain it.

Design for all three. Remember: developers read code 10x more than they write it.

---

## Naming principles (priority order)

The first five are SKILL.md's canonical priority order (Consistency → Understandability → Specificity → Brevity → Searchability). Pronounceability (6) and Austerity (7) are supplementary tie-breakers, not new top-level priorities.

### 1. Consistency and uniqueness (highest priority)

One concept, one name, everywhere.

```ts
// BAD - inconsistent names for the same concept
export const getUserById = (id: UserId): Promise<User | null> => { /* ... */ };
export const fetchCustomerById = (id: CustomerId): Promise<Customer | null> => { /* ... */ };
export const retrieveClientById = (id: ClientId): Promise<Client | null> => { /* ... */ };

// GOOD - consistent
export const getUser = (id: UserId): Promise<User | null> => { /* ... */ };
export const getOrder = (id: OrderId): Promise<Order | null> => { /* ... */ };
export const getProduct = (id: ProductId): Promise<Product | null> => { /* ... */ };
```

### 2. Understandability

Domain language, not technical jargon.

```ts
// BAD - technical
const arr = users.filter((u) => u.isActive);

// GOOD - domain language
const activeCustomers = users.filter((user) => user.isActive);
```

### 3. Specificity

Ban vague names: `data`, `info`, `manager`, `handler`, `processor`, `utils` as primary names. Use them only when nothing else fits (rare).

```ts
// BAD - vague
export const processData = (data: unknown): unknown => { /* ... */ };

// GOOD - specific
export const validatePayment = (payment: Payment): ValidationResult => { /* ... */ };
```

### 4. Brevity (but not at the cost of clarity)

Short names are good only if meaning is preserved.

```ts
// BAD - too cryptic
const usrLst = getUsrs();

// BAD - unnecessarily long
const listOfAllActiveUsersInTheSystem = getActiveUsers();

// GOOD - brief but clear
const activeUsers = getActiveUsers();
```

### 5. Searchability

Names should be unique enough to grep.

```ts
// BAD - common word, hard to search
const data = fetch();

// GOOD - unique, searchable
const orderSummary = fetchOrderSummary(orderId);
```

### 6. Pronounceability

You should be able to say it in conversation.

```ts
// BAD
const genymdhms = generateYearMonthDayHourMinuteSecond();

// GOOD
const timestamp = generateTimestamp();
```

### 7. Austerity

Avoid unnecessary filler words.

```ts
// BAD - redundant
const userData = user;
type UserType = { /* ... */ };

// GOOD
const user = /* ... */;
type User = { /* ... */ };
```

---

## Object Calisthenics (translated to a class-free codebase)

These exercises originally targeted OO code. They translate cleanly to our typed-record / arrow-function style.

### 1. One level of indentation per function

```ts
// BAD - multiple levels
export const process = (orders: Order[]): void => {
  for (const order of orders) {
    if (isValidOrder(order)) {
      for (const item of order.items) {
        if (item.inStock) {
          processItem(item);
        }
      }
    }
  }
};

// GOOD - extract
export const shipValidOrders = (orders: Order[]): void => {
  orders.filter(isValidOrder).forEach(processOrder);
};

export const processOrder = (order: Order): void => {
  order.items.filter((item) => item.inStock).forEach(processItem);
};
```

### 2. Do not use the `else` keyword

Early returns, guard clauses, or dispatch records.

```ts
// BAD - else
export const getDiscount = (user: User): number => {
  if (user.isPremium) {
    return 20;
  } else {
    return 0;
  }
};

// GOOD - early return
export const getDiscount = (user: User): number => {
  if (user.isPremium) return 20;
  return 0;
};
```

### 3. Wrap all primitives and strings

Primitives that carry domain meaning become branded types with validating factories.

```ts
// BAD - primitive obsession
export const createUser = (email: string, age: number): User => {
  if (!email.includes('@')) throw new Error();
  if (age < 0) throw new Error();
  return { email, age };
};

// GOOD - branded types
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

export const createUser = (e: Email, a: Age): User => ({ email: e, age: a });
```

Two primitives get a named callout because their failure mode is silent (hard rule 12; `references/reliability.md`, Money and time):

```ts
// Money holds integer minor units, never a float: 0.1 + 0.2 !== 0.3, and the rounding lands on an invoice
export type Money = { readonly cents: number; readonly currency: 'EUR' | 'USD' };
export const money = (cents: number, currency: Money['currency']): Money => {
  if (!Number.isSafeInteger(cents)) throw new Error('invalid Money.cents');
  return { cents, currency };
};
// arithmetic lives with the type and refuses a currency mismatch; display formatting divides at the presentation edge
```

Instants are UTC in the domain (epoch milliseconds or a branded ISO instant); a timezone is a display concern applied only at the presentation edge, never stored inside the domain value.

### 4. First-class collections

Any module that holds a collection with domain meaning should have no other fields. Extract the collection as its own module.

```ts
// BAD - collection mixed with other state
export type Order = {
  readonly id: OrderId;
  readonly items: OrderItem[];
  readonly customerId: CustomerId;
  readonly total: Money;
};

// GOOD - collection gets its own module
// src/orders/order-items.ts
export type OrderItems = { readonly items: readonly OrderItem[] };
export const emptyOrderItems = (): OrderItems => ({ items: [] });
export const addToOrderItems = (items: OrderItems, item: OrderItem): OrderItems => ({ items: [...items.items, item] });
export const orderItemsTotal = (items: OrderItems): Money =>
  items.items.reduce((sum, i) => addMoney(sum, i.price), money(0, 'EUR'));
export const isOrderItemsEmpty = (items: OrderItems): boolean => items.items.length === 0;

// src/orders/order.ts
export type Order = {
  readonly id: OrderId;
  readonly items: OrderItems;
  readonly customerId: CustomerId;
};
```

### 5. One dot per line (Law of Demeter)

Do not chain through object graphs. Expose a behaviour function that does the walk internally.

```ts
// BAD - train wreck
const city = order.customer.address.city;

// GOOD - ask the immediate friend
const city = orderShippingCity(order);

// where
export const orderShippingCity = (order: Order): City => customerShippingCity(order.customer);
export const customerShippingCity = (customer: Customer): City => addressCity(customer.address);
```

### 6. Do not abbreviate

If a name is too long to type, the module is doing too much.

```ts
// BAD
const custRepo = createCustRepo();
const ord = createOrd();

// GOOD
const customerRepo = createCustomerRepo();
const order = createOrder(orderId('ord-1'));
```

### 7. Keep all entities small

- Functions < 10 lines.
- Modules < 50 lines.
- Files < 100 lines.

If larger, it is probably doing too much. Split it.

### 8. Small record shapes

A record with many fields usually mixes concerns. Push back on god records by composing smaller typed records. Two fields is a useful aspiration, but three or four can be fine when every field genuinely belongs to the same domain concept. Treat this as a design smell prompt, not a hard lint rule.

```ts
// BAD - too many fields, mixed concerns
export type Order = {
  readonly id: OrderId;
  readonly customerId: CustomerId;
  readonly items: OrderItem[];
  readonly total: Money;
  readonly status: OrderStatus;
};

// GOOD - compose smaller records
export type OrderDetails = {
  readonly customer: CustomerId;
  readonly items: OrderItems;
};

export type Order = {
  readonly id: OrderId;
  readonly details: OrderDetails;
};
```

### 9. No getters or setters

Records expose behaviour functions, not raw reads and writes. In our style this is automatic: we have no classes, so there are no getters to write. Access fields directly on the record where appropriate, but for domain operations always expose a verb function.

```ts
// BAD - caller does the work
if (account.balance >= amount) {
  account = { ...account, balance: account.balance - amount };
}

// GOOD - behaviour-rich function
export type WithdrawResult =
  | { readonly kind: 'success'; readonly account: Account }
  | { readonly kind: 'insufficientFunds' };

export const withdraw = (account: Account, amount: Money): WithdrawResult => {
  if (!canWithdraw(account, amount)) return { kind: 'insufficientFunds' };
  return { kind: 'success', account: { ...account, balance: subMoney(account.balance, amount) } };
};

// caller tells, module decides
const result = withdraw(account, amount);
if (result.kind === 'success') account = result.account;
```

---

## Comments

### When to write a comment

Only to explain WHY, never WHAT or HOW. Code explains what and how. Comments explain business reasons, non-obvious decisions, or warnings.

```ts
// BAD - explains what (redundant)
// Add 1 to counter
counter += 1;

// GOOD - explains why
// Compensate for 0-based indexing in legacy API
counter += 1;
```

### Prefer self-documenting code

Rename to make intent clear instead of adding a comment.

```ts
// BAD - comment needed
// Check if user can access premium features
if (user.subscriptionLevel >= 2 && !user.isBanned) { /* ... */ }

// GOOD - self-documenting
if (canAccessPremiumFeatures(user)) { /* ... */ }
```

---

## Formatting

### Vertical spacing

- Related code stays together.
- Blank lines between concepts.
- Public / most-important API at the top.

### Horizontal spacing

- Consistent indentation (2 spaces).
- Space around operators.
- Max line length 180 (our Prettier config).

### Storytelling

Code should read top-to-bottom like a story. High-level at top, details below.

```ts
// Public API first
export const processOrder = (order: Order, deps: ProcessOrderDeps): ProcessResult => {
  validate(order);
  const total = calculateTotal(order);
  return save(order, total, deps);
};

// Supporting functions below, in order of appearance
const validate = (order: Order): void => { /* ... */ };
const calculateTotal = (order: Order): Money => { /* ... */ };
const save = (order: Order, total: Money, deps: ProcessOrderDeps): ProcessResult => { /* ... */ };
```

Non-exported helpers go at the bottom or in a sibling module. Exported API stays at the top, easy to scan.
