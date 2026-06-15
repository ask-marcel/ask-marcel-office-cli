# Software Architecture

## The goal

Enable the team to:

1. **Add** features with minimal friction.
2. **Change** existing features safely.
3. **Remove** features cleanly.
4. **Test** features in isolation.
5. **Deploy** independently when possible.

## Architectural principles

### 1. Vertical slices (feature-first)

Organise by feature, not by technical layer.

```
BAD - layer-first
src/
  controllers/
    userController.ts
    orderController.ts
  services/
    userService.ts
    orderService.ts
  repositories/
    userRepository.ts
    orderRepository.ts

GOOD - feature-first
src/
  users/
    user-controller.ts
    user-service.ts
    user-repository.ts
  orders/
    order-controller.ts
    order-service.ts
    order-repository.ts
```

**Why.** Changes to the "users" feature stay in `users/`. High cohesion within features, low coupling between them.

### 2. Horizontal boundaries (layers)

Separate concerns into layers with clear dependencies.

```
+--------------------------------------+
|           Presentation               |  UI, controllers, CLI entry
+--------------------------------------+
|           Application                |  Use cases, orchestration
+--------------------------------------+
|             Domain                   |  Business logic, value objects, entities
+--------------------------------------+
|          Infrastructure              |  Database, APIs, external integrations
+--------------------------------------+
```

### 3. The dependency rule

**Dependencies point INWARD.**

```
Infrastructure -> Application -> Domain
      outer          middle         inner
```

- Inner layers know NOTHING about outer layers.
- Domain has zero dependencies on infrastructure.
- Use function-type contracts to invert dependencies.

```ts
// Domain defines the contract (inner)
export type UserRepo = {
  save: (user: User) => Promise<void>;
  findById: (id: UserId) => Promise<User | null>;
};

// Infrastructure implements it (outer)
export const createPostgresUserRepo = (db: Database): UserRepo => ({
  save: async (user) => {
    /* SQL here */
  },
  findById: async (id) => {
    /* SQL here */
  },
});

// Domain use-case depends on the contract, never on the postgres implementation
export const createGetUser = (repo: UserRepo) => async (id: UserId): Promise<User | null> => repo.findById(id);
```

### 4. Contracts

Function-type aliases define boundaries between components.

```ts
// The contract
export type PaymentGateway = {
  charge: (amount: Money, card: CardDetails) => Promise<ChargeResult>;
  refund: (chargeId: ChargeId) => Promise<RefundResult>;
};

// Multiple implementations possible
export const stripeGateway: PaymentGateway = { /* ... */ };
export const payPalGateway: PaymentGateway = { /* ... */ };
export const fakeGateway: PaymentGateway = { /* ... */ }; // in-memory fake for tests
```

### 5. Cross-cutting concerns

Concerns that span multiple features: logging, auth, validation, error handling.

Options in our style:
- Middleware / interceptors.
- Higher-order functions that wrap other functions.
- Decorator functions (from `references/design-patterns.md`).

```ts
// Higher-order function wraps a handler with logging
export type Handler<Req, Res> = (request: Req) => Promise<Res>;

export const withLogging = <Req extends { path: string }, Res extends { status: number }>(
  handler: Handler<Req, Res>
): Handler<Req, Res> =>
  async (request) => {
    logger.info('request', { path: request.path });
    const response = await handler(request);
    logger.info('response', { status: response.status });
    return response;
  };
```

### 6. Conway's Law

> "Organisations design systems that mirror their communication structure."

**Implication.** Team structure affects architecture. Align both intentionally.

---

## Common architectural styles

### Layered architecture

Traditional layers: Presentation -> Business -> Persistence.

**Pros.** Simple, well-understood.
**Cons.** Can become a "big ball of mud" without discipline. No clear story about dependency direction.

### Hexagonal architecture (Ports and Adapters)

Domain at the centre, adapters around the edges.

```
        +---------------------+
        |    HTTP adapter     |
        +----------+----------+
                   |
+------------------v------------------+
|              DOMAIN                  |
|   +--------------------------+      |
|   |     business logic        |      |
|   |     use cases             |      |
|   +--------------------------+      |
+------------------+------------------+
                   |
        +----------v----------+
        |  database adapter   |
        +---------------------+
```

- **Ports** | function-type contracts defined by the domain.
- **Adapters** | concrete implementations that connect to the outside world.

### Clean architecture

Similar to hexagonal, with explicit layers:

1. **Entities** | enterprise business rules.
2. **Use cases** | application business rules.
3. **Interface adapters** | controllers, presenters, gateways.
4. **Frameworks and drivers** | web, DB, external interfaces.

---

## Feature-driven structure (frontend)

Aligns with the Next.js variant:

```
src/
  components/
    atoms/      | (Atomic Design: no internal composition)
    molecules/  | (import atoms only)
    organisms/  | (import atoms + molecules)
  page/         | page shells
  features/
    auth/
      components/
      hooks/
      services/
      types/
      index.ts  | public API
    checkout/
      components/
      hooks/
      services/
      types/
      index.ts
  lib/          | truly shared
    hooks/
    i18n/
    utils/
```

See `references/nextjs-monorepo.md` for the full layout and rules.

---

## Clean Architecture layout (backend, canonical)

For any non-trivial Bun backend — pipelines, batch jobs, CLIs with real integrations — use this strict six-folder layout. Files land where they belong based on what they depend on, not on which feature they serve.

```
src/
├── domain/                     # branded value objects, Zod schemas, pure utilities, FlowConfig builder
│   ├── ids.ts                  #   branded IDs (UserId, OrderId, ...)
│   ├── urls.ts                 #   SafeUrl, canonicalUrl helpers
│   ├── schemas/                #   Zod shape definitions (no IO)
│   ├── utilities/              #   split-text, retry-on-err, rss-parser, format-error
│   ├── result.ts               #   Result<T, E> + helpers
│   └── flow.ts                 #   pure FlowConfig builder
├── use-cases/                  # coordinators + the port interfaces they depend on
│   ├── ports/                  #   type-only interfaces for every side-effectful dependency
│   │   ├── sheets.ts
│   │   ├── llm.ts
│   │   ├── telegram.ts
│   │   ├── rss-fetcher.ts
│   │   ├── prompt-loader.ts
│   │   ├── logger.ts
│   │   └── step-error.ts
│   ├── select-news.ts
│   ├── post-telegram.ts
│   └── run-pipeline.ts
├── infra/                      # concrete adapters that implement the ports
│   ├── google-auth.ts
│   ├── sheets-google.ts
│   ├── gemini-llm.ts
│   ├── telegram-http.ts
│   ├── rss-fetcher-http.ts
│   ├── prompt-loader-fs.ts
│   └── logger.ts
├── presenter/                  # CLI argv parsing, usage text, output formatting
│   └── cli.ts
├── composition/                # the composition root: env parser + buildPipelineDeps
│   ├── env.ts
│   └── build-deps.ts           #   the ONLY place infra/ meets use-cases/
├── test-helpers/               # in-memory fakes for every port + test data builders
│   ├── sheets-fake.ts
│   ├── llm-fake.ts
│   ├── telegram-fake.ts
│   ├── logger-fake.ts
│   ├── capture-rejection.ts
│   └── test-flow.ts
└── main.ts                     # thin entry: argv → presenter → composition → use-case
```

### Dependency rule (strict, inward-only)

| Folder | Depends on |
|:---|:---|
| `domain/` | nothing inside `src/` |
| `use-cases/` | `domain/` + its own `ports/` (types only) |
| `infra/` | `domain/` + the ports it implements + third-party SDKs |
| `presenter/` | `domain/` only |
| `composition/` | everything (this is the only place where concrete `infra/` meets use-case deps) |
| `test-helpers/` | `domain/` + ports (no production code depends on test-helpers) |
| `main.ts` | `composition/` + `presenter/` + `infra/` (for top-level error notification only) |

Invariants the layout protects:

- The domain is zero-dependency on anything in `src/` except shared `domain/*`. `grep -rn "from '.*infra" src/domain src/use-cases` must return nothing.
- Ports are type-only modules: they declare interfaces, never implementations.
- The composition root is the only place where you may import both an adapter and a use-case.
- Tests instantiate fakes; no production code imports from `test-helpers/`.

### Adding a new external service

1. Define the port under `src/use-cases/ports/<service>.ts` — type only, returns `Promise<Result<T, <Service>Error>>` where the error is a discriminated union.
2. Create the in-memory fake under `src/test-helpers/<service>-fake.ts` with an optional `errors` config so tests can inject `err(...)`.
3. Implement the real adapter under `src/infra/<service>-<protocol>.ts` (e.g. `sheets-google.ts`, `tmdb-http.ts`). The adapter is the only place `try/catch` wraps the SDK call.
4. Wire it into `PipelineDeps` and `src/composition/build-deps.ts`.
5. Write use-case tests that inject the fake and pattern-match on `Result.ok`.

### Framework vs configuration

Domain-specific data — brand lists, tenant slugs, feature flags, tier-discount rates, per-environment API endpoints — is **configuration**, not framework code. It lives in env vars, JSON files, or an external source loaded at runtime. The framework code never contains string-literal unions of brand slugs, hardcoded record maps of brands, or `if (brand === 'acme') ...` branches.

Signal: if a new tenant requires editing a union type or a switch statement, the code is fused with the data. Refactor to drive the behaviour from config.

### Composition root testability (no skip lists)

`src/composition/build-deps.ts` is **not** a coverage-skip. It is fully unit-testable when two ergonomic switches are in place:

1. Every "where do I read state from" point — file path, env var, system clock, random source — is parameterisable.
2. Every "what do I write to / log to" sink can be injected as a port (Logger, EmailSender, Clock).

The pattern is an optional `BuildDepsConfig` argument with sensible defaults that preserve production behaviour:

```ts
// src/composition/build-deps.ts
export type BuildDepsConfig = {
  readonly tokenStorePath?: string;
  readonly logger?: Logger;
};

export const buildPipelineDeps = async (
  env: Env,
  config: BuildDepsConfig = {}
): Promise<PipelineDeps> => {
  const logger = config.logger ?? createWinstonLogger();
  const tokenStore = createTokenStoreFs({ path: config.tokenStorePath ?? '.tokens.json' });
  // ... rest unchanged
};
```

Production callers (just `src/main.ts`) call `buildPipelineDeps(env)` with no second argument; behaviour is identical. Tests pass `{ tokenStorePath: tmpDir + '/tokens.json', logger: createLoggerFake() }`. With the token store empty and `staleAfterMs` set so refresh paths short-circuit, end-to-end execution is offline and the wiring covers itself.

Also export the otherwise-private helpers (`overlayToken`, `buildEnrichmentPlugin`, etc.) so individual branches can be tested in isolation rather than only through the composed `buildPipelineDeps` call.

The earlier policy that left `build-deps.ts` in the coverage skip list as "verified live, not via units" was hedging. With the two switches above, the file goes from "skipped" to 100%. The same logic applies to any composition or wiring file that feels untestable: parameterise the inputs, inject the outputs, and the test seam appears.

## Feature-driven structure (simpler alternative, for small scripts)

For throwaway scripts, one-off CLIs, or pre-pipeline prototypes, a simpler feature-first layout is fine. Skip the port/adapter split until the repo genuinely needs it.

```
src/
  <feature>/
    domain.ts
    use-case.ts
    infra.ts
  utils/
    logger.ts
```

Graduate to the Clean Architecture layout above when: the script gains a second external service, needs tests with fakes, or grows past ~500 lines. See `references/bun-typescript.md` for the small-script tsconfig / eslint setup.

---

## The walking skeleton

Start with a minimal end-to-end slice:

1. Thinnest possible feature that touches all layers.
2. Deployable from day one.
3. Proves the architecture works.

Example walking skeleton for e-commerce:
- User can view ONE product (hardcoded).
- User can add it to a cart.
- User can "checkout" (just logs the attempt).

From there, flesh out each feature fully with TDD.

---

## Testing architecture

```
+--------------------------------------------+
|       E2E / acceptance tests               |  few, slow, high confidence
+--------------------------------------------+
|       Integration tests                    |  some, medium speed
+--------------------------------------------+
|       Unit tests                           |  many, fast, isolated
+--------------------------------------------+
```

Test by layer:
- **Domain** | unit tests (most tests here).
- **Application** | integration tests with faked infrastructure.
- **Infrastructure** | integration tests with real dependencies.
- **E2E** | critical paths only.

See `references/testing.md` for the full strategy.

---

## Architecture Decision Records (ADRs)

Document significant decisions:

```markdown
# ADR 001 | Use PostgreSQL for persistence

## Status
Accepted

## Context
We need a database. Options: PostgreSQL, MongoDB, MySQL.

## Decision
PostgreSQL, because:
- ACID compliance
- Team familiarity
- JSON support for flexibility

## Consequences
- Need PostgreSQL expertise.
- Schema migrations required.
- Excellent query capabilities.
```

Store ADRs under `docs/adr/` in the repo. One file per decision, numbered.

---

## Red flags in architecture

- Circular dependencies between modules.
- Domain depending on infrastructure.
- Framework code in business logic.
- No clear boundaries between features.
- Shared mutable state across modules.
- "utils" or "common" packages that grow forever.
- Database schema driving the domain model (domain should drive the schema, not the other way).
