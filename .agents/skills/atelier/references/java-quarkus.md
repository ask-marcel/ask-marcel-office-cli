# Java Variant (Quarkus)

The atelier standard translated to Java. Same commitments (TDD with hand-written fakes, Clean Architecture with inward dependencies, errors as values, value types at trust boundaries, executable gates), expressed in Java 21+ idioms: records, sealed interfaces, constructor injection. The framework flavour is Quarkus (JAX-RS resources, Panache, Flyway, MicroProfile config); Spring translates one-to-one if a repo demands it, but do not mix the two.

Pick this variant when the repo has a `pom.xml` (or `build.gradle`) and Java sources. The hard rules apply as translated below; rules 21 and 22 (design system) do not apply, a Java backend has no UI layer.

## Runtime and toolchain

- **Java 21+ LTS**, records, sealed interfaces, pattern matching for switch.
- **Maven through the wrapper, always**: `./mvnw`, never a locally installed `mvn` whose version drifts (the rule 5 analogue). Commit the wrapper.
- **Quarkus** for anything HTTP-shaped; a plain `main` for CLIs and batch jobs.
- Formatting is machine-owned: **Spotless with google-java-format**, one plugin version pinned in the parent pom, `./mvnw spotless:apply` locally, `spotless:check` in the gate (rule 8 analogue).

## `pom.xml` conventions (rule 19 translated)

- **Exact versions only.** Never a version range (`[1.0,)`) and never a `-SNAPSHOT` dependency in `main`. Maven resolves ranges to whatever is newest that day, which is the `"latest"` footgun with different syntax.
- All versions live in `<properties>` or the parent pom / BOM; children declare nothing loose. One formatter version, one runtime BOM, inherited everywhere (the one-committed-config rule).
- **maven-enforcer-plugin** makes it executable: `requireJavaVersion` (a bare `21` means "at least 21"; avoid the `[21,)` range form, which `check-pom.sh` would flag as a version range), `requireReleaseDeps` (no `-SNAPSHOT` dependencies), `requireUpperBoundDeps`.
- Renovate (or equivalent) keeps pins current so a pinned version never rots into a known-vulnerable one; **OWASP dependency-check** (or the platform's scanner) runs in CI and fails on high CVSS, the `bun audit` analogue: daily schedule plus a PR run scoped to `pom.xml`.
- **google-java-format on JDK 16+** needs the `jdk.compiler` exports: commit a one-line `.mvn/jvm.config` (shown under the canonical pom below). Harmless where unneeded.

### Canonical `pom.xml`

The gate skeleton every atelier Java repo carries, framework-free: a Quarkus service adds the pinned Quarkus BOM in `<dependencyManagement>` and its extensions on top. This block is extracted verbatim by `scripts/smoke-test-java.sh` in the skill repo's CI, so drift here fails a build, not a user.

```xml
<?xml version="1.0" encoding="UTF-8"?>
<project xmlns="http://maven.apache.org/POM/4.0.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://maven.apache.org/POM/4.0.0 http://maven.apache.org/xsd/maven-4.0.0.xsd">
  <modelVersion>4.0.0</modelVersion>
  <groupId>com.example</groupId>
  <artifactId>app</artifactId>
  <version>0.1.0</version>
  <packaging>jar</packaging>

  <!-- Exact pins only (rule 19): no ranges, no -SNAPSHOT dependencies. -->
  <properties>
    <maven.compiler.release>21</maven.compiler.release>
    <project.build.sourceEncoding>UTF-8</project.build.sourceEncoding>
    <junit.version>5.14.4</junit.version>
    <compiler.plugin.version>3.15.0</compiler.plugin.version>
    <surefire.plugin.version>3.5.6</surefire.plugin.version>
    <spotless.plugin.version>3.8.0</spotless.plugin.version>
    <google-java-format.version>1.35.0</google-java-format.version>
    <jacoco.plugin.version>0.8.15</jacoco.plugin.version>
    <pitest.plugin.version>1.25.7</pitest.plugin.version>
    <pitest-junit5.plugin.version>1.2.3</pitest-junit5.plugin.version>
    <pitest.threads>4</pitest.threads>
    <enforcer.plugin.version>3.6.3</enforcer.plugin.version>
  </properties>

  <dependencies>
    <dependency>
      <groupId>org.junit.jupiter</groupId>
      <artifactId>junit-jupiter</artifactId>
      <version>${junit.version}</version>
      <scope>test</scope>
    </dependency>
  </dependencies>

  <build>
    <plugins>
      <!-- rule 15: warnings are errors -->
      <plugin>
        <groupId>org.apache.maven.plugins</groupId>
        <artifactId>maven-compiler-plugin</artifactId>
        <version>${compiler.plugin.version}</version>
        <configuration>
          <compilerArgs>
            <arg>-Xlint:all</arg>
            <arg>-Werror</arg>
          </compilerArgs>
        </configuration>
      </plugin>
      <plugin>
        <groupId>org.apache.maven.plugins</groupId>
        <artifactId>maven-surefire-plugin</artifactId>
        <version>${surefire.plugin.version}</version>
      </plugin>
      <!-- rule 8: one committed formatter, machine-owned -->
      <plugin>
        <groupId>com.diffplug.spotless</groupId>
        <artifactId>spotless-maven-plugin</artifactId>
        <version>${spotless.plugin.version}</version>
        <configuration>
          <java>
            <googleJavaFormat>
              <version>${google-java-format.version}</version>
            </googleJavaFormat>
          </java>
        </configuration>
      </plugin>
      <!-- coverage tiers: 100 on domain+usecases, 80 on infra/api/composition -->
      <plugin>
        <groupId>org.jacoco</groupId>
        <artifactId>jacoco-maven-plugin</artifactId>
        <version>${jacoco.plugin.version}</version>
        <executions>
          <execution>
            <goals><goal>prepare-agent</goal></goals>
          </execution>
          <execution>
            <id>check-tiers</id>
            <phase>verify</phase>
            <goals><goal>check</goal></goals>
            <configuration>
              <rules>
                <rule>
                  <element>PACKAGE</element>
                  <includes>
                    <include>com.example.app.domain*</include>
                    <include>com.example.app.usecases*</include>
                  </includes>
                  <limits>
                    <limit><counter>LINE</counter><value>COVEREDRATIO</value><minimum>1.00</minimum></limit>
                  </limits>
                </rule>
                <rule>
                  <element>PACKAGE</element>
                  <includes>
                    <include>com.example.app.infra*</include>
                    <include>com.example.app.api*</include>
                    <include>com.example.app.composition*</include>
                  </includes>
                  <limits>
                    <limit><counter>LINE</counter><value>COVEREDRATIO</value><minimum>0.80</minimum></limit>
                  </limits>
                </rule>
              </rules>
            </configuration>
          </execution>
        </executions>
      </plugin>
      <!-- mutation gate: invoked by the hook and CI, not bound to verify (the mutate:staged analogue) -->
      <plugin>
        <groupId>org.pitest</groupId>
        <artifactId>pitest-maven</artifactId>
        <version>${pitest.plugin.version}</version>
        <dependencies>
          <dependency>
            <groupId>org.pitest</groupId>
            <artifactId>pitest-junit5-plugin</artifactId>
            <version>${pitest-junit5.plugin.version}</version>
          </dependency>
        </dependencies>
        <configuration>
          <targetClasses>
            <param>com.example.app.domain.*</param>
            <param>com.example.app.usecases.*</param>
          </targetClasses>
          <targetTests>
            <param>com.example.app.*</param>
          </targetTests>
          <mutationThreshold>90</mutationThreshold>
          <timestampedReports>false</timestampedReports>
          <!-- Free parallelism: mutation results are thread-independent. Tune to cores.
               Incremental history is NOT a free lever here (see the prose below). -->
          <threads>${pitest.threads}</threads>
        </configuration>
      </plugin>
      <!-- rule 19 at build time: JDK floor, no snapshot deps, converging versions -->
      <plugin>
        <groupId>org.apache.maven.plugins</groupId>
        <artifactId>maven-enforcer-plugin</artifactId>
        <version>${enforcer.plugin.version}</version>
        <executions>
          <execution>
            <id>enforce</id>
            <goals><goal>enforce</goal></goals>
            <configuration>
              <rules>
                <requireJavaVersion>
                  <version>21</version>
                </requireJavaVersion>
                <requireReleaseDeps>
                  <message>No -SNAPSHOT dependencies (rule 19)</message>
                </requireReleaseDeps>
                <requireUpperBoundDeps />
              </rules>
            </configuration>
          </execution>
        </executions>
      </plugin>
    </plugins>
  </build>
</project>
```

`.mvn/jvm.config` (one line, committed):

```text
--add-exports jdk.compiler/com.sun.tools.javac.api=ALL-UNNAMED --add-exports jdk.compiler/com.sun.tools.javac.code=ALL-UNNAMED --add-exports jdk.compiler/com.sun.tools.javac.file=ALL-UNNAMED --add-exports jdk.compiler/com.sun.tools.javac.parser=ALL-UNNAMED --add-exports jdk.compiler/com.sun.tools.javac.tree=ALL-UNNAMED --add-exports jdk.compiler/com.sun.tools.javac.util=ALL-UNNAMED
```

## Source architecture

Same rings, package-per-layer inside a feature-first root where the repo is big enough (`references/architecture.md` governs; this is the Java expression):

```
src/main/java/com/example/app/
├── domain/            # value records, pure logic, Result: zero framework imports
├── usecases/          # application services + the port interfaces they depend on
│   └── ports/         #   interfaces only, returning Result
├── infra/             # adapters: JPA repositories, HTTP clients, LLM adapter, logger config
├── api/               # inbound JAX-RS resources + wire DTOs (the inbound adapter ring)
└── composition/       # CDI wiring: @Produces methods, config records
src/main/resources/
├── application.properties
└── db/migration/      # Flyway V*__*.sql, versioned, expand-contract (rule 30)
src/test/java/...      # tests mirror the tree; fakes in a shared testsupport package
```

Dependency rule unchanged: `domain` imports nothing from the framework; `usecases` sees domain + its own ports; `infra` and `api` implement/consume them; only `composition` (and the CDI container) sees everything. The check is mechanical: `grep -rn "import jakarta.ws.rs\|import io.quarkus" src/main/java/com/example/app/domain src/main/java/com/example/app/usecases` returns nothing (framework annotations on use-case classes are tolerated only for `@ApplicationScoped`; prefer producing them from `composition` when practical).

## The hard rules, translated

| Rule (TS) | Java expression |
|:---|:---|
| 1 no `class` | Inverted mechanism, same intent: **records** for data, **sealed interfaces** for unions, small **final classes** with constructor injection for behaviour. No inheritance for reuse (composition only), no static mutable state, no field injection (`@Inject` on fields hides the contract; constructors state it) |
| 2 no `function` decl | n/a |
| 3 no `interface` | Inverted: interfaces ARE the port mechanism. Keep them small and role-shaped (ISP); one capability per interface |
| 4 no `console.*` | No `System.out`/`System.err`/`printStackTrace`. Inject a `Logger` (JBoss/SLF4J) through the constructor; redaction configured once (below) |
| 5 Bun only | `./mvnw` only; the Quarkus CLI is sugar over it |
| 6 explicit return types | Native. Avoid `var` on any public or port surface; locals may use it when the right side names the type |
| 7-9 imports/style/ESM | Spotless owns style; no wildcard imports |
| 10 no custom error classes | Business failures are `Err` values, never bespoke exception types. Exceptions are for bugs and framework edges; never use checked exceptions on domain surfaces |
| 11 TDD | Unchanged (JUnit 5) |
| 12 branded types | Value **records with validating compact constructors** plus a `parse(...)` factory returning `Result` (below) |
| 13 no `mock` | **No Mockito, no EasyMock, no `@InjectMock`.** Hand-written fakes implement the port interface; enforce by keeping mock libraries out of the pom entirely |
| 14 outside-in classicist | Unchanged: the SUT is the application service; domain runs real; only secondary ports get fakes |
| 15 zero warnings, no inline ignores | No `@SuppressWarnings`, ever. Compile with `-Xlint:all -Werror`; SonarJava/Error Prone severities change at project level with a comment |
| 16 `Result` at IO boundaries | `Result<T, E>` as a sealed interface (below); every port method returns it |
| 17 `try/catch` quarantine | Adapters in `infra/` catch SDK/JPA exceptions and translate to `Err`; use-cases pattern-match with `switch`; one top-level handler at the entry point |
| 18 no curried chains | n/a |
| 19 no `latest` | Exact versions, no ranges, no SNAPSHOT deps (above) |
| 20 Bun file API | `java.nio.file.Files` in `infra/` only; never `java.io.File` gymnastics in domain code |
| 21-22 design system | n/a (no UI) |
| 23-26 commits, tests, identity | Unchanged: same hooks, same confirmation gates |
| 27-34 production disciplines | Unchanged; Java expressions in their references and below |

## `Result` in Java (rule 16)

Shipped as copyable assets (`assets/java/Result.java`, `Ok.java`, `Err.java`); copy them at bootstrap rather than retyping. Sealed, so a `switch` is exhaustive and the compiler owns totality:

```java
// domain/Result.java
public sealed interface Result<T, E> permits Ok, Err {}
public record Ok<T, E>(T value) implements Result<T, E> {}
public record Err<T, E>(E error) implements Result<T, E> {}
```

Per-port errors are sealed unions too, the discriminated-union analogue:

```java
public sealed interface RepoError permits RepoError.Io, RepoError.NotFound {
  record Io(String message) implements RepoError {}
  record NotFound(String id) implements RepoError {}
}

public interface InvoiceRepo {
  Result<Invoice, RepoError> find(InvoiceId id);
}
```

Use-cases consume with pattern matching; no `try/catch` in a use-case:

```java
return switch (repo.find(id)) {
  case Ok<Invoice, RepoError>(var invoice) -> process(invoice);
  case Err<Invoice, RepoError>(var e) -> new Err<>(StepError.from("find-invoice", e));
};
```

## Value records at trust boundaries (rule 12)

The compact constructor is the guard (constructing an invalid instance is a bug, so it throws); the static `parse` is the boundary factory returning `Result` (expected-invalid input is a value). Shipped as `assets/java/Email.java`, the exemplar to copy for `Money` (integer minor units), `UserId`, `IsoCountryCode`, and every other domain primitive:

```java
public record Email(String value) {
  public Email { if (!value.matches("^[^@\\s]+@[^@\\s]+$")) throw new IllegalArgumentException("email"); }
  public static Result<Email, String> parse(String raw) {
    return raw.matches("^[^@\\s]+@[^@\\s]+$") ? new Ok<>(new Email(raw)) : new Err<>("invalid_email");
  }
}

public record Money(long cents, Currency currency) {}   // integer minor units, never double
// instants are java.time.Instant (UTC) in the domain; ZoneId only at the presentation edge
```

At the HTTP edge, Bean Validation covers shape (`@Valid`, `@NotNull`, `@Positive` on wire DTOs); domain invariants live in the records. Wire DTOs never cross into use-cases: map DTO to domain at the resource, domain to response record on the way out (`references/architecture.md`, The internal model is yours).

## Ports, fakes, and the logger (rules 4, 13)

```java
// usecases/ports/Blobs.java
public interface Blobs { Result<Void, BlobError> put(String key, byte[] body); }

// test: src/test/java/.../testsupport/MemoryBlobs.java, a hand-written fake
public final class MemoryBlobs implements Blobs {
  public final Map<String, byte[]> store = new ConcurrentHashMap<>();
  private final BlobError failWith; // optional errors knob, like the TS fakes
  public MemoryBlobs() { this(null); }
  public MemoryBlobs(BlobError failWith) { this.failWith = failWith; }
  public Result<Void, BlobError> put(String key, byte[] body) {
    if (failWith != null) return new Err<>(failWith);
    store.put(key, body);
    return new Ok<>(null);
  }
}
```

Constructor injection everywhere; the CDI container is the composition root. Tests never boot the container for unit work: `new PlaceOrder(new MemoryBlobs(), new FakeClock(), recordingLogger)` and assert on outcomes.

Logging: JBoss/SLF4J injected via constructor, JSON output in production, and redaction configured once at the logging layer (a rewrite filter for `password`, `token`, `authorization`, plus the natural identifiers of rule 27: `email`, `phone`, `name`). Never `System.out`, never a secret or natural identifier in a message (`references/privacy.md`).

## Persistence (rules 30, 31; `references/reliability.md`)

- **Flyway owns the schema**: every change a `V*__*.sql`, expand-contract for anything shipped, never a hand ALTER, never `hibernate.hbm2ddl.auto=update` outside a throwaway spike.
- **Writes through the ORM, hot reads explicit**: Panache/JPA persists; list endpoints use explicit projection queries you can EXPLAIN; no lazy-relation walks on a list path (the N+1).
- **Keyset pagination**, never OFFSET: `find("createdAt > ?1 or (createdAt = ?1 and id > ?2) order by createdAt, id", ...)` over a composite index; stream large exports (`try (var s = find(...).stream())`).
- **Optimistic locking**: `@Version` on every mutable entity; map `OptimisticLockException` to 409 plus the current state at the resource.
- **Soft delete**: `deletedAt` stamp plus `@SQLRestriction("deleted_at IS NULL")` on the entity so reads stay honest by default; subject erasure is the privacy exception (`references/privacy.md`).
- **The entity is not the domain model**: Panache entities stay inside `infra/`; repositories map entity to domain record at the boundary (`references/architecture.md`, The domain model is not the database model).
- **Outbox**: persist the `OutboxEntry` in the same `@Transactional` method as the state change; a `@Scheduled` worker delivers with retries, idempotent on `(topic, aggregateId)`.
- **Tenant isolation**: owner from `SecurityIdentity`, `set_config` for RLS inside the same `@Transactional` block, runtime datasource role `NOBYPASSRLS` (`references/isolation.md`).

## Inbound resources (`api/`)

- **Authenticated by default**: `quarkus.http.auth.permission.default.policy=authenticated` in `application.properties`; permit-all is the explicit, justified exception (health probes). Identity comes from OIDC (rule 33: never hand-rolled auth); rate limits and TLS are baseline, not per-route memory (`references/security.md`, One baseline).
- Resource-shaped endpoints, not screen-shaped (`references/architecture.md`, The backend is a client-agnostic API).
- The resource maps `Result` to HTTP: `Ok` to 200/201, domain-expected failures to their status, use-case `StepError` to 500 with a generic body (internals stay in the log with the trace id).
- **OpenAPI from the code**: MicroProfile OpenAPI annotations (`@Operation`, `@APIResponse`, example objects) so the published spec cannot drift (`references/governance.md`).
- Every network client the app opens has connect and per-request timeouts, bounded jittered retries (`@Retry(maxRetries = 3, jitter = 200)` on the adapter), and an `Idempotency-Key` where the operation is not naturally safe to repeat (rule 29).
- Personal data never in a `@QueryParam` or a log line (rule 27): user-typed search terms arrive in a `@Valid` POST body.

## Observability

Quarkus ships OpenTelemetry: enable it, add `@WithSpan` on application services (or wrap at composition), Micrometer counters/timers tagged by outcome for behaviour metrics, JSON logs carrying the trace id. Alerting and SLO discipline as in `references/observability.md`.

## Testing (rules 11, 13, 14; gates)

- **Unit tests are plain JUnit 5**, no `@QuarkusTest`, no container, sub-millisecond: the SUT is the application service with fakes injected by hand. `@QuarkusTest` boots the app and belongs to the integration ring only.
- **Integration tests** use `@QuarkusTest` + REST Assured against the real edge: the happy path, the bypass (`references/testing.md`, Test the bypass: wrong role is 403, missing token 401), the cross-tenant 404, and the forged-header seam (`references/isolation.md`). Testcontainers (or dev services) provide a real database; fixtures are synthetic (rule 34).
- **Test names are business scenarios**: `premiumCustomerGets20PercentOff`, `crossTenantReadIsNotFound`, `regressionEmptyCartTotalsToZero`.
- **Coverage tiers with JaCoCo**: 100% line on `domain` + `usecases`, 80% on `infra` + `api` + `composition`, enforced by per-package `<rule>` limits in the JaCoCo check goal so the build fails loudly, untested classes included in the denominator (the coverage-preload principle is native here: JaCoCo counts all classes in the module).
- **Mutation testing with PIT**: `mutationThreshold=90` on `domain` + `usecases` packages. Incremental history is NOT free in current PIT: 1.25.7 errors `History has been enabled but no history plugin has been installed/activated` for BOTH `withHistory` and explicit `historyInputFile`/`historyOutputFile` (verified via `smoke-test-java`), and the only history plugin is Arcmutate's commercial `+arcmutate_history`. So the free speed levers are the narrow target scope (`targetClasses`/`targetTests`), parallel `threads` (set in the pom, mutation results are thread-independent), and running the gate only when staged files touch it (the hook does exactly that); in a multi-module repo, scope PIT per module. If incremental speed becomes a hard requirement at scale, Arcmutate is the only supplier, which makes it a licence decision, not a library swap. Same policy as Stryker: no per-file exclusions because tests feel awkward; tighten the test or refactor.
- **PIT on Quarkus**: no Quarkus-specific mutation tool exists; PIT plus `pitest-junit5-plugin` is the whole story, and the plugin (1.2.3+, needs Quarkus 3.22.x+) is the only Quarkus-aware piece. It auto-disables Quarkus's JaCoCo extension, the classic thing that broke PIT there. The scoping above is also what keeps this healthy: because `domain`/`usecases` are covered by plain JUnit 5 (not `@QuarkusTest`), PIT never runs over a container-boot test, so Quarkus's build-time augmentation never triggers the `tests did not pass without mutation` failure and no Quarkus container stands up per mutant. If you widen PIT onto `@QuarkusTest` classes, expect both that failure (patch with `avoidCallsTo` on `io.quarkus.*` plus test excludes on older plugin versions) and the per-mutant container cost; the atelier design avoids both by construction. Pin the Quarkus BOM at or above 3.22.x.
- **Evals for any LLM hole** gate the merge like PIT does (`references/ai.md`).

## Gates and hooks

Same two git hooks as the Bun variant, shell only, wired with `git config core.hooksPath .githooks`. All four artifacts ship in the skill's `assets/`; copy them, never hand-write:

- `assets/commit-msg`: the shipped Conventional Commits validator, unchanged (rule 23; it is dependency-free shell).
- `assets/pre-commit-java`: six gates, cost-ascending like the Bun hook's eight: commit size (`scripts/check-commit-size.sh`, shared with the Bun variant, ≤10 files / ≤300 lines) → pom sanity (`scripts/check-pom.sh`: no version ranges anywhere, no `-SNAPSHOT` in `<parent>`/`<dependencies>`/`<plugins>`; the project's own dev version may be a SNAPSHOT) → `gitleaks protect --staged` → `./mvnw -q spotless:check` → `./mvnw -q verify` (compile with `-Werror`, unit + integration tests, JaCoCo tier check) → PIT with history, run only when staged files touch the `domain`/`usecases` mutation scope.

```bash
cp <skill>/assets/pre-commit-java        .githooks/pre-commit
cp <skill>/assets/commit-msg             .githooks/commit-msg
cp <skill>/assets/check-commit-size.sh   scripts/check-commit-size.sh
cp <skill>/assets/check-pom.sh           scripts/check-pom.sh
chmod +x .githooks/pre-commit .githooks/commit-msg scripts/*.sh
git config core.hooksPath .githooks
```

CI runs the identical chain plus the scheduled dependency scan and, where the repo deploys, the compose portability gate and deployment events (`references/delivery.md`).

## Bootstrap checklist (fresh Java repo)

1. `quarkus create app com.example:app` (or the Maven archetype); commit the wrapper; delete sample code.
2. Parent pom: start from the canonical `pom.xml` above (compiler `-Werror`, Spotless + google-java-format, JaCoCo tier rules, PIT `mutationThreshold=90` scoped to `domain`/`usecases`, enforcer with `requireJavaVersion`/`requireReleaseDeps`/`requireUpperBoundDeps`), add the pinned Quarkus BOM and extensions, commit `.mvn/jvm.config`. Exact versions everywhere.
3. Scaffold packages: `domain`, `usecases/ports`, `infra`, `api`, `composition`; copy the shipped domain assets into `domain` rather than hand-writing them, then rename their package to your own groupId:
   ```bash
   cp <skill>/assets/java/{Result,Ok,Err,Email}.java src/main/java/<pkg>/domain/
   # Result/Ok/Err are the sealed Result union (rule 16); Email is the value-record
   # exemplar (rule 12) to copy for Money, UserId, and every other domain primitive.
   ```
4. `application.properties`: authenticated-by-default policy, OIDC config placeholders, OTel enabled, JSON logging with the redaction filter, datasource for the constrained runtime role.
5. Flyway: `src/main/resources/db/migration/V1__init.sql`; dev services or Testcontainers for the integration ring.
6. Test support: `testsupport` package with the first hand-written fakes (logger recorder, clock); **no Mockito in the pom**.
7. Hooks: copy the four assets as above (`pre-commit-java`, `commit-msg`, `check-commit-size.sh`, `check-pom.sh`); `git config core.hooksPath .githooks`; optional `gitleaks` install. Verify the pom gate once: `bash scripts/check-pom.sh`.
8. Walking skeleton: one use-case returning `Ok` through its port, its value record, its JUnit test (propose the test first, rule 24), one resource with its REST Assured test including the 401 case.
9. Verify green: `./mvnw spotless:check verify`, PIT on the skeleton, hooks reject a junk message and an oversized commit.
10. `.claude/LESSONS.md` header; choose the commit identity (rule 26); stage and propose the first commit (rule 25).

## Red flags (Java-specific)

- A Mockito import, `@InjectMock`, or a mock-library dependency in the pom (rule 13).
- `@SuppressWarnings` anywhere; a warning "fixed" by silencing (rule 15).
- A version range or `-SNAPSHOT` dependency (rule 19); `hbm2ddl.auto=update` outside a spike (rule 30).
- Business failure thrown as an exception across a port instead of returned as `Err` (rules 10, 16).
- A Panache entity or wire DTO crossing into `usecases/` (the internal model is yours).
- Field injection on anything with behaviour; static mutable state; a singleton holding per-user state (`references/reliability.md`, Stateless by default).
- `@QueryParam` carrying an email, a name, or user-typed text (rule 27); a `System.out` anywhere (rule 4).
- A list endpoint walking lazy relations or paging with OFFSET (`references/reliability.md`).
- An entity two actors edit with no `@Version` (rule 31); a repository read that ignores `deletedAt` by hand-rolled query while the entity carries the restriction (rule 30).
