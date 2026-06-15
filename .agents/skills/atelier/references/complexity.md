# Managing Complexity

## The two kinds of complexity

### Essential complexity

Inherent to the problem domain. Cannot be removed, only managed. Business rules, domain logic, user requirements all fit here.

### Accidental complexity

Introduced by our solutions. CAN and SHOULD be minimised. Poor abstractions, unnecessary indirection, framework ceremony, technical debt.

**Goal.** Minimise accidental complexity while expressing essential complexity clearly.

---

## Detecting complexity

### 1. Change amplification

Small changes require touching many files.

**Symptom.** "To add this field, I need to update 15 files."

**Cause.** Scattered responsibilities, poor abstraction boundaries, data copied across layers instead of flowing through a single record.

### 2. Cognitive load

Code is hard to understand, requires holding too much in memory.

**Symptom.** "I need to understand 10 other modules to understand this one."

**Cause.** Tight coupling, hidden dependencies, unclear naming, missing branded types.

### 3. Unknown unknowns

Behaviour is surprising, side effects are hidden.

**Symptom.** "I changed this, and something completely unrelated broke."

**Cause.** Global mutable state, hidden dependencies, implicit contracts, missing tests.

---

## The XP values for fighting complexity

From Extreme Programming.

1. **Communication.** Code should communicate clearly. Names, structure, tests all contribute.
2. **Simplicity.** Do the simplest thing that could possibly work.
3. **Feedback.** Fast feedback loops catch complexity early (TDD, CI, code review).
4. **Courage.** Refactor aggressively. Do not let complexity accumulate.
5. **Respect.** Respect future readers, including future-you.

---

## KISS | Keep It Simple

> "The simplest solution that works is usually the best."

How to apply:
1. Start with the obvious solution.
2. Only add complexity when REQUIRED.
3. Prefer boring, well-understood approaches.
4. Question every abstraction.

```ts
// Over-engineered
export const createUserServiceFactoryProvider = (): UserServiceFactoryProvider => {
  // singleton + factory + provider. Three patterns for one need.
};

// KISS - just the function
export const getUser = async (repo: UserRepo, id: UserId): Promise<User | null> => repo.findById(id);
```

---

## YAGNI | You Aren't Gonna Need It

> "Do not build features until they are actually needed."

Warning signs in a requirement or review comment:
- "We might need this later."
- "It would be nice to have."
- "Just in case."
- "For future extensibility."

The cost of YAGNI violations:
1. Development time building unused features.
2. Maintenance burden on code that must be kept alive.
3. Cognitive load on everyone who reads the code.
4. Wrong abstractions that are expensive to undo.

```ts
// YAGNI violation
export type User = {
  readonly name: Name;
  readonly email: Email;
  readonly middleName?: Name;
  readonly secondaryEmail?: Email;
  readonly faxNumber?: PhoneNumber;
  readonly linkedinProfile?: Url;
  readonly twitterHandle?: Handle;
};

// YAGNI - only what is needed NOW
export type User = {
  readonly name: Name;
  readonly email: Email;
};
```

---

## DRY with Rule of Three

> "Every piece of knowledge should have a single, unambiguous representation."

### BUT | do not extract until the third occurrence

A wrong abstraction costs more to undo than duplication does to tolerate.

```
Duplication #1 - leave it.
Duplication #2 - note it, leave it.
Duplication #3 - NOW extract it.
```

### Example

```ts
// First time - leave it
export const processUserOrder = (order: Order): void => {
  validate(order);
  applyTax(order);
  save(order);
};

// Second time - note the similarity, leave it
export const processGuestOrder = (order: Order): void => {
  validate(order);
  applyTax(order);
  save(order);
  sendGuestEmail(order);
};

// Third time - NOW extract
export const processCorporateOrder = (order: Order): void => {
  validate(order);
  applyTax(order);
  save(order);
  applyCorporateDiscount(order);
};

// After three, extract the common spine
export type OrderPostProcess = (order: Order) => void;

export const processOrder = (order: Order, postProcess: OrderPostProcess = () => {}): void => {
  validate(order);
  applyTax(order);
  save(order);
  postProcess(order);
};
```

Now `processGuestOrder` and `processCorporateOrder` become one-liners that pass their specific post-processor.

---

## Separation of Concerns

> "Each module should address a single concern."

Concerns to separate:
- Business logic vs infrastructure.
- What (policy) vs how (mechanism).
- Input vs processing vs output.
- Data vs behaviour.

```ts
// BAD - mixed concerns
export const processOrder = async (order: Order): Promise<void> => {
  if (order.items.length === 0) throw new Error('empty');    // validation
  let total = 0;
  for (const item of order.items) total += item.price * item.quantity; // business logic
  await db.query(`INSERT INTO orders ...`);                  // persistence
  await mailer.send(order.customer.email, 'confirmed');      // notification
};

// GOOD - separated
export type ProcessOrderDeps = {
  validator: OrderValidator;
  calculator: OrderCalculator;
  repo: OrderRepo;
  notifier: OrderNotifier;
};

export const processOrder = async (order: Order, deps: ProcessOrderDeps): Promise<ProcessResult> => {
  deps.validator.validate(order);
  const total = deps.calculator.calculateTotal(order);
  const saved = await deps.repo.save(order, total);
  await deps.notifier.notifyConfirmation(saved);
  return { kind: 'success', order: saved };
};
```

---

## Managing technical debt

### Kinds of technical debt

1. Deliberate | conscious trade-off for speed.
2. Accidental | mistakes, lack of knowledge.
3. Bit rot | code degrades as the world around it changes.

### Boy Scout Rule

> "Leave the code better than you found it."

Every time you touch code:
- Improve one small thing.
- Fix one naming issue.
- Extract one function.
- Add one missing test.

### When to pay down debt

- When it is in your path (you are already editing the area).
- When it is blocking new features.
- When it is causing bugs.
- During dedicated refactoring time.

### When NOT to refactor

- Code that works and will not change.
- Code being replaced soon.
- When you do not have tests to protect you.

---

## The four elements of simple design (priority order)

1. **Runs all the tests.** If it does not work, nothing else matters.
2. **Expresses intent.** Clear names, obvious structure, code tells the story.
3. **No duplication.** DRY after Rule of Three.
4. **Minimal.** Fewest modules and functions possible. Remove anything unnecessary.

If all four are true, the design is simple enough. Stop polishing and ship.
