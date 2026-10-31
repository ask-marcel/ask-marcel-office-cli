# 0004: Layer zones in the lint (rule 37)

- Status: accepted (2026-10-30)
- Date: 2026-10-30

## Context

Atelier rule 37 says that dependencies point inward, and that the lint enforces this for each
layer. `domain/` imports nothing else in `src/`. `use-cases/` sees `domain/` and its own `ports/`.
`infra/` sees `domain/` and the ports that it implements. `presenter/` sees `domain/`.
`composition/` sees everything. Production code never imports `test-helpers/`.

On main 3eb2167 about 195 production imports break the rule:

- Use-cases import types from infra: `GraphError` (72 in read), `ReadGraph`, `WriteGraph`,
  `FetchFn`, `AuthManager` and `AuthError`.
- About 25 imports in read use-cases are runtime imports of converter adapters: turndown, mammoth,
  the OOXML zip reader, pdf, msg, eml, SheetJS and the line diff.
- `core/src/infra/read-graph.ts` imports `postIfReadOnly` from `use-cases/commands/read-only-post.ts`.
- The core presenter imports the name fill functions from `use-cases/commands/bin-names.ts`, and
  the `Logger` type from `use-cases/ports/logger.ts`.

## Decision

1. `GraphError` moves to `core/src/domain/graph-error.ts`. It is pure data, and every layer may
   import domain. `read-only-post.ts` is pure apart from `GraphError`, so it moves to domain too.
   This removes the infra -> use-cases import.
2. `ReadGraph` and `WriteGraph` move to `core/src/use-cases/ports/graph.ts`. `FetchFn` stays in
   infra, because outside infra only tests use it. In auth, `login` and `logout` take an
   `AuthSession` port (`auth/src/use-cases/ports/auth-session.ts`): the five methods they call,
   with `AuthError` and `CachedTierInfo`. The full `AuthManager` stays in infra and fits the port.
   Infra imports these types from the ports and implements them.
3. Each old module re-exports what moved out of it: infra re-exports the moved types,
   `use-cases/commands/bin-names.ts` re-exports the fill functions, and `command-directory.ts`
   re-exports `PACKAGE_ROLES`. Tests are not in the zones, so no existing test changes (rule 24).
4. `PACKAGE_ROLES` moves to `core/src/domain/package-role.ts`, and the pure name fill functions
   move to `core/src/domain/bin-names.ts`, with `NameSource` as a structural type. The presenter
   types its logger as a local type with the one method it calls, so it needs no port import.
5. The converter adapters reach the read commands through a port. Until 2026-10-31 they were
   runtime imports on an allowlist in the read use-cases zone. Then `use-cases/ports/converters.ts`
   declared the `Converters` bag, infra built the real one (`CONVERTERS`), the registry became
   `createReadCommands(converters)`, and the helpers took the bag as their first parameter. The
   allowlist is gone, and the read use-cases zone bans every infra import.
6. The zones are one `no-restricted-imports` block per layer in `eslint.base.js`. Each block repeats
   the mock ban and the test-helpers ban, because ESLint replaces the options of a rule when two
   blocks match a file. The patterns match the relative paths and the
   `@ask-marcel/office-core/<layer>/` paths. `check-lint-fixtures.ts` proves each zone in each folder.

The work lands in four pull requests: core, then auth and write, then read, then the lint.

## Options considered

- Inject the converter adapters now, through a `Converters` port passed to `execute`. This gives the
  full rule in one series, but it changes the `ReadCommand` contract and many existing tests.
  Deferred at first, then done on 2026-10-31 through factories in read (`createExecute`), so
  core's `execute` contract did not change.
- Let use-cases import types from infra (`allowTypeImports`). It is cheap, but use-cases still name
  infra modules, so it is weaker than rule 37. Rejected.
- Move the converter adapters into domain. Domain would then import npm libraries. Rejected.
- Let the presenter see `use-cases/ports`. It is a deviation that a local type makes unnecessary.
  Rejected.

## Consequences

- A use-case that imports infra, or a presenter that imports use-cases, fails lint in every folder.
- The converter adapters reach a read command only through the `Converters` port, so a test can
  pass any bag in place of the real one.
- The old infra modules keep re-exports of the moved types for the tests.

## Reversal

Remove the zone blocks from `eslint.base.js` and their fixtures. The moved types can stay where
they are, because each layer may import domain and ports.
