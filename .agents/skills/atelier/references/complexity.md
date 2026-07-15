# Managing Complexity

> **Note on examples.** Port and use-case signatures in this file are sometimes elided to `Promise<T>` (or throw on business failure) for brevity where error handling is not the lesson. In real code every IO port returns `Promise<Result<T, PortError>>` and every use-case returns `Promise<Result<Summary, StepError>>` — hard rule 16, see `references/result-type.md`.

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

## The lazy ladder

KISS and YAGNI answer *how simple*; the lazy ladder turns them into a procedure you run *before writing code*. The cheapest, most correct, fastest-to-review code is the code you never wrote — so stop at the first rung that solves the problem:

1. **Does it need to exist?** YAGNI at the top of the ladder. If no current requirement forces it, the rung is "don't build it." A deleted requirement beats an elegant implementation.
2. **Standard library or language feature?** `Array`/`Map`/`Set`, `structuredClone`, `Intl`, `Object.groupBy`, optional chaining, etc. Reach for the language before hand-rolling.
3. **Native runtime capability?** In a Bun repo the runtime replaces whole categories of dependency: `Bun.file`/`Bun.write` for file IO (hard rule 20), `crypto.subtle` and `crypto.randomUUID()` for crypto, the global `fetch`, `URL`/`URLSearchParams`, `Bun.password` for hashing. Prefer the platform before `bun add`.
4. **A dependency already in `package.json`?** If something installed already does the job, use it — do not add a second library for the same capability (and never `bun add` a near-duplicate; hard rule 19).
5. **One clear line?** If the whole thing collapses to one readable expression, that is the implementation.
6. **Only then** write the minimum that works — the absolute smallest correct version, no speculative abstractions (what "no speculative seams" does and does not mean: see *Defer the build, not the seam* below).

**Tiebreaker.** When two stdlib approaches are the same size, pick the one that is edge-case-correct and more efficient (e.g. `re.exec` in a loop over `.matchAll` spread only when you need the perf; the *correct* boundary handling always wins over the cute one-liner). Two more reflexes: **delete before adding** (a diff that removes code is usually the better fix), and **boring over clever** (the next reader, possibly you in six months, pays for cleverness).

### Defer the build, not the seam

YAGNI applies to implementations, never to the thin contracts that keep a later swap cheap. When a future need is likely (a heavier mailer, a real cache, a second provider), put the boundary in now and defer the machinery:

- The **port** (one function-type alias) and the **smallest real adapter** ship today. In this standard that seam is structural anyway: every side-effectful dependency sits behind a port with a test seam from day one (hard rule 13), so the interface costs nothing extra.
- The **heavy implementation** (retries beyond `retryOnErr`, batching, sharding, a circuit breaker, a pooled cache) waits until the need is real. A dependency earns its circuit breaker; it is not born with one (`references/reliability.md`).

So the lazy ladder's "no speculative abstractions" bans inventing *domain* abstractions and heavyweight adapters for imagined futures. It never bans the port itself: deleting the seam to save one type alias buys nothing today and makes the eventual swap a rewrite.

### Simplicity is not negligence

The ladder trims *speculation*, never *safety*. Five things stay off the chopping block no matter how lazy the rung:

- **Trust-boundary validation** — branded value objects with validating factories at every source→sink crossing (see `references/security.md`, `references/object-design.md`).
- **`Result` error handling at IO boundaries** — real failure modes belong in the type (hard rule 16). "No error handling for impossible scenarios" means skip the *impossible* ones; a network call failing is not impossible.
- **Security** — the source-to-sink discipline is never "simplified away."
- **Accessibility** — semantic elements, focus states, ARIA on interactive controls in the design system (see `references/atomic-design.md`).
- **Anything the user explicitly requested** — laziness applies to *unrequested* scope, never to the actual ask.

The test: would a reviewer call this *lazy* (good — minimal, used the platform) or *negligent* (bad — dropped a checkpoint)? If the latter, you cut the wrong thing.

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

### Boy Scout Rule (bounded by Surgical Changes)

> "Leave the code better than you found it."

This does **not** license drive-by edits. It is subordinate to SKILL.md Behavioural Guideline #3 (Surgical Changes): touch only what the task requires, match local style, and never "improve" adjacent code, reformat, or refactor things that are not broken. Unrelated dead code or naming you notice in passing gets *mentioned*, not fixed, unless the user asks.

The Boy Scout impulse applies **only to lines your change already touches**, and only when tests protect it:
- Improve one small thing on a line you were editing anyway.
- Fix a name that your own change made inaccurate.
- Extract a function when your change created the third duplication (Rule of Three).
- Add a missing test for behaviour you just touched (proposed and confirmed first — rule 24).

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
