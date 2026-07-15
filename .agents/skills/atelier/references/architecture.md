# Software Architecture

> Code samples here are simplified to show *structure*. Some port signatures elide the mandatory `Result<T, E>` wrapper — e.g. a use-case shown returning `Result<User | null, RepoError>` would, in real code, aggregate to `Result<Summary, StepError>` (hard rule 16) — and some inner arrows omit explicit return types. The layout and dependency direction are the point; apply the hard rules in full when writing the real thing.

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
export type RepoError = { type: 'io'; message: string };

export type UserRepo = {
  save: (user: User) => Promise<Result<void, RepoError>>;
  findById: (id: UserId) => Promise<Result<User | null, RepoError>>;
};

// Infrastructure implements it (outer)
export const createPostgresUserRepo = (db: Database): UserRepo => ({
  save: async (user) => {
    /* SQL here, wrapped in ok()/err() */
  },
  findById: async (id) => {
    /* SQL here, wrapped in ok()/err() */
  },
});

// Domain use-case depends on the contract, never on the postgres implementation
export const createGetUser = (repo: UserRepo) => async (id: UserId): Promise<Result<User | null, RepoError>> => repo.findById(id);
```

IO ports always return `Promise<Result<T, PortError>>`, never bare `Promise<T>` — see `references/result-type.md`.

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
// Higher-order function wraps a handler with logging.
// The logger is a parameter, not a module-level singleton (hard rule 4).
export type Handler<Req, Res> = (request: Req) => Promise<Res>;

export const withLogging = <Req extends { path: string }, Res extends { status: number }>(
  handler: Handler<Req, Res>,
  logger: Logger
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

- **Layered** (Presentation → Business → Persistence): simple and well-understood, but without a dependency-direction rule it decays into a big ball of mud.
- **Hexagonal (Ports and Adapters)**: the domain at the centre; **ports** are function-type contracts defined by the domain, **adapters** the concrete implementations that connect to the outside world. The layer diagram at the top of this file *is* this style.
- **Clean architecture**: hexagonal with named rings — entities (enterprise rules), use cases (application rules), interface adapters (controllers/presenters/gateways), frameworks and drivers (web, DB, external interfaces).

This standard's `src/{domain,use-cases,infra,presenter,composition}` layout (below) is the concrete expression of the last two.

---

## Feature-driven structure (frontend)

In the Next.js variant there is no `features/` folder. A vertical slice is expressed as one organism per page section plus a page shell per route, with the slice's logic in `src/lib/<feature>/` (e.g. `src/lib/guides/`). Components never live outside the design system (`src/components/{atoms,molecules,organisms}` — hard rules 21-22), and state lives in `src/lib/hooks/`, wired by the page shells.

```
src/
  components/
    atoms/      | no internal composition
    molecules/  | import atoms only
    organisms/  | import atoms + molecules; one per page section
  page/         | page shells; one per route, wire organisms to lib state
  lib/
    guides/     | feature logic for the "guides" slice
    hooks/      | state, consumed by page shells
    i18n/
```

See `references/nextjs-monorepo.md` for the full layout and rules, and `references/atomic-design.md` for the component layer rules.

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
│   ├── http/                   # inbound Bun.serve adapter (server archetype)
│   │   ├── server.ts           #   route table + the one request-level try/catch
│   │   └── to-response.ts      #   pure Result → Response mapper
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
| `infra/` | `domain/` + the ports it implements (+ the use-case `Result`/error types an inbound adapter maps) + third-party SDKs |
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

### Inbound HTTP (server archetype)

The canonical archetype is a CLI/batch job that runs and `process.exit`s. When the entry instead serves HTTP, **the server is an `infra/` adapter — the inbound mirror of an outbound one — not a new layer and not a `presenter/` file.** An outbound adapter (`telegram-http.ts`) turns a thrown SDK error *into* a `Result`; the inbound adapter turns a use-case `Result` *into* a `Response`. It reuses infra's existing `try/catch` quarantine slot and 80% coverage tier, and **adds no new layer** — it lives under `infra/`, covered by that row's inbound-adapter clause.

Why not `presenter/`: the `Result → Response` mapper must read `Summary` and `StepError`, which live in `use-cases/ports/`. The `presenter → domain/ only` rule forbids that import, so a "presenter mapper" silently breaks the dependency table. The `infra/` row explicitly covers the use-case `Result`/error types an inbound adapter maps — so the mapper belongs there.

```ts
// src/infra/http/to-response.ts — pure, total: a use-case Result → an HTTP Response.
import type { Result } from '../../domain/result.ts';
import type { Summary, StepError } from '../../use-cases/ports/step-error.ts';

export const toResponse = (result: Result<Summary, StepError>): Response => {
  if (result.ok) return Response.json(result.value, { status: 200 });
  const { step, cause, message } = result.error;
  // The use-case flatten already stringified the port `kind` into `cause: string`, so there is no
  // typed discriminant to switch on here — a use-case failure is a 500 by default. Precise client
  // errors (400) are decided upstream at the branded request checkpoint, before this runs.
  return Response.json({ step, error: cause, message }, { status: 500 });
};
```

```ts
// src/infra/http/server.ts — inbound adapter; the one request-level try/catch lives here.
import type { Result } from '../../domain/result.ts';
import type { OrderInput } from '../../domain/order-input.ts';
import type { Summary, StepError } from '../../use-cases/ports/step-error.ts';
import { parseOrderBody } from '../../domain/order-input.ts'; // branded checkpoint (rule 12)
import { formatError } from '../../domain/utilities/format-error.ts';
import { toResponse } from './to-response.ts';

type HttpDeps = { readonly placeOrder: (input: OrderInput) => Promise<Result<Summary, StepError>> };

export const createHttpServer = (deps: HttpDeps): { readonly fetch: (req: Request) => Promise<Response> } => ({
  fetch: async (req) => {
    try {
      const parsed = parseOrderBody(await req.text());
      if (!parsed.ok) return Response.json({ error: parsed.error.message }, { status: 400 });
      return toResponse(await deps.placeOrder(parsed.value));
    } catch (e) {
      return Response.json({ error: formatError(e) }, { status: 500 });
    }
  },
});
```

`src/main.ts` stays the single thin entry with its one top-level catch: env → `buildPipelineDeps(env)` → `createHttpServer(deps)` → `Bun.serve({ port: env.port, fetch: server.fetch })`. Do not add a second `main-web.ts`; if the app is server-shaped, `main.ts` *is* the serving entry, and the Dockerfile gains `EXPOSE <port>` (`references/bun-typescript.md` § Containerization).

Three rules this archetype leans on:

1. **The body is an untrusted source — brand it (rule 12).** `parseOrderBody` is the validating checkpoint; precise `400`s are decided here, where the type is still narrow.
2. **A use-case error defaults to `500`.** The flatten destroys the port `kind` (see `references/result-type.md`), so never switch on `StepError.cause` to fabricate `401`/`404`/`429` — that non-exhaustive lookup rots silently. Honoring typed statuses is a real upgrade: carry a status number through the flatten where TS still enforces totality over the `kind` union — adopt it when a requirement lands, not speculatively.
3. **Register each new `src/infra/http/*.ts` in `scripts/coverage-preload.ts` in the same commit** — or the 80% gate passes trivially on uncovered files.

No router until the third route (Rule of Three) — a `switch (new URL(req.url).pathname)` covers one or two endpoints. No framework; that choice stays out of scope.

### API shape and the three model boundaries

Four rules keep the layers separable when the app exposes or consumes an API. They are the same idea applied at three boundaries: each side of a boundary owns its model, and one mapping function translates.

**The backend is a client-agnostic API (resource-shaped, never screen-shaped).** One resource-shaped API that every client (web, iOS, Android, a third-party integration) consumes the same way. A screen-shaped endpoint (`GET /mobile-home-screen` returning banner + greeting + widgets) recouples the backend to one interface and changes every time that screen does; resource endpoints (`/orders`, `/promotions`) let a new client compose what it needs from what already exists.

**The domain model is not the database model.** The two serve different purposes, so they get different shapes: the domain record carries behaviour and invariants; the row carries storage concerns (nullable columns, foreign keys). The repository adapter is the one mapping point (`row -> toDomain(row)`), and the DB shape stops there. The business drives the schema, never the schema the business (Red flags, below).

**The internal model is not the wire model.** Wire DTOs (request bodies, API responses) never cross into use-cases in either direction: the inbound adapter parses the body through the branded checkpoint into a domain type (rule 12), and the outbound edge maps domain to a response record shaped for the API. Serialising a persistence entity straight to JSON welds all three models together; then every rename breaks clients.

**The frontend depends on a contract it controls (the gateway).** Agree the API contract first (the schema that also generates the docs, `references/governance.md`), then reach data only through a gateway port with two implementations, the real client and a fake returning canned data in the same shape, chosen at one wiring point. The gateway returns `Result`, so a failure is a state the UI renders rather than a throw that blanks the screen (rule 16), and it maps the wire DTO into the frontend's own model, so an API rename is absorbed in one `toOrder(dto)` function instead of rippling through components.

```ts
// the frontend gateway: same port discipline as any backend adapter
export type OrderGateway = { readonly list: () => Promise<Result<Order[], LoadError>> };
const toOrder = (d: ApiOrder): Order => ({ id: d.order_id, total: money(d.total_cents, 'EUR'), customerName: d.customer.first_name ?? 'Guest' });
export const httpOrders = (api: Api): OrderGateway => ({
  list: async () => {
    const dto = await api.get<ApiOrder[]>('/orders');
    return dto.ok ? ok(dto.value.map(toOrder)) : dto;
  },
});
export const fakeOrders: OrderGateway = { list: async () => ok([{ id: '1', total: money(8000, 'EUR'), customerName: 'Ada' }]) };
```

The UI is then built, demoed, and tested in parallel with the backend, and the flip from fake to real is one wiring line. (In the Next.js static variant, build-time data loading plays this role; the gateway applies to the server-app sub-variant and any client fetching at runtime: `references/nextjs-monorepo.md`.)

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

Test by layer, most tests at the bottom of the pyramid:
- **Domain** | unit tests through the primary port (most tests here).
- **Application** | integration tests with faked infrastructure.
- **Infrastructure** | integration tests with real dependencies.
- **E2E** | critical paths only.
- **Performance** | load-test gates on routes with a latency budget (`references/reliability.md`).

See `references/testing.md` for the full strategy.

---

## Recording architectural decisions

Two tiers (see `references/governance.md`, Decision records). Every significant decision is a `[decision]` entry in `.claude/LESSONS.md`: append-only, one line, superseded by a newer `[decision]` when it changes (see `references/lessons.md`). Decisions whose rejected alternatives and reversal path are worth keeping (a vendor, a storage engine, a deliberate lock-in) additionally get a full ADR under `docs/adr/`, committed in the same change as the code it explains. For decisions worth pressure-testing *before* they are made, the atelier-grill-me companion skill walks the decision tree and its output is the natural ADR draft.

---

## Red flags in architecture

- Circular dependencies between modules.
- Domain depending on infrastructure.
- Framework code in business logic.
- No clear boundaries between features.
- Shared mutable state across modules.
- "utils" or "common" packages that grow forever.
- Database schema driving the domain model (domain should drive the schema, not the other way).
- A screen-shaped endpoint (`/mobile-home-screen`) instead of resource endpoints clients compose.
- A wire DTO or persistence entity crossing into use-cases, or serialised straight to the client.
- A frontend component calling `fetch` directly instead of going through its gateway port.
