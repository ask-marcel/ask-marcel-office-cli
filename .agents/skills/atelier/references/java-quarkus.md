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
- **maven-enforcer-plugin** makes it executable: `requireJavaVersion` (a bare `21` means "at least 21"; avoid the `[21,)` range form, which `check-pom.sh` would flag as a version range), `requireReleaseDeps` (no `-SNAPSHOT` dependencies), `requireUpperBoundDeps` (framework-free only: under a platform BOM the BOM converges versions, The Quarkus delta below), `banDynamicVersions` (no range, `LATEST` or `RELEASE`, including one held in a version property), and `bannedDependencies` for the mock libraries (rule 13: `org.mockito`, `org.easymock`, `org.powermock`, `org.jmockit`, `org.jmock`, `quarkus-junit5-mockito` and its current name `quarkus-junit-mockito`, `quarkus-panache-mock`; enforcer 3.x walks the whole tree, so a transitive Mockito is caught too).
- Renovate (or equivalent) keeps pins current so a pinned version never rots into a known-vulnerable one; **OWASP dependency-check** (or the platform's scanner) runs in CI and fails on high CVSS, the `bun audit` analogue: daily schedule plus a PR run scoped to `pom.xml`. The canonical pom pins the plugin in `<pluginManagement>` with `failBuildOnCVSS` 7: its default, 11, never fails the build. It reads the NVD API key from `NVD_API_KEY`, which the workflow passes from a repository secret (request a free key from NVD; without one the database download is throttled to hours).
- **google-java-format on JDK 16+** needs the `jdk.compiler` exports: commit a one-line `.mvn/jvm.config` (shown under the canonical pom below). Harmless where unneeded.

### Canonical `pom.xml`

The gate skeleton every atelier Java repo carries, framework-free: a Quarkus service applies the Quarkus delta below on top of it. This block is extracted verbatim by `scripts/smoke-test-java.sh` in the skill repo's CI, so drift here fails a build, not a user.

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
    <!-- the mutation scope; scripts/pit-changed.sh narrows it to the changed classes with -Dpitest.targetClasses= -->
    <pitest.targetClasses>com.example.app.domain.*,com.example.app.usecases.*</pitest.targetClasses>
    <enforcer.plugin.version>3.6.3</enforcer.plugin.version>
    <pmd.plugin.version>3.28.0</pmd.plugin.version>
    <archunit.version>1.5.0</archunit.version>
    <dependency-check.plugin.version>13.0.0</dependency-check.plugin.version>
  </properties>

  <dependencies>
    <dependency>
      <groupId>org.junit.jupiter</groupId>
      <artifactId>junit-jupiter</artifactId>
      <version>${junit.version}</version>
      <scope>test</scope>
    </dependency>
    <dependency>
      <!-- the dependency rule as a test (rule 37): assets/java/LayerRulesTest.java -->
      <groupId>com.tngtech.archunit</groupId>
      <artifactId>archunit-junit5</artifactId>
      <version>${archunit.version}</version>
      <scope>test</scope>
    </dependency>
  </dependencies>

  <build>
    <pluginManagement>
      <plugins>
        <!-- the CVE watchdog (assets/audit-java.yml runs dependency-check:check on a schedule,
             never in verify). The default failBuildOnCVSS is 11, which never fails; 7 fails on a
             high or critical CVE. The NVD key comes from the environment, never the pom -->
        <plugin>
          <groupId>org.owasp</groupId>
          <artifactId>dependency-check-maven</artifactId>
          <version>${dependency-check.plugin.version}</version>
          <configuration>
            <failBuildOnCVSS>7</failBuildOnCVSS>
            <nvdApiKeyEnvironmentVariable>NVD_API_KEY</nvdApiKeyEnvironmentVariable>
          </configuration>
        </plugin>
      </plugins>
    </pluginManagement>
    <plugins>
      <!-- rule 15: warnings are errors. -classfile is the one lint off, project-level: it
           reports defects inside third-party class files, not in this source (the
           MicroProfile Config API behind every @ConfigProperty trips it), and no
           source-level fix or suppression can reach it -->
      <plugin>
        <groupId>org.apache.maven.plugins</groupId>
        <artifactId>maven-compiler-plugin</artifactId>
        <version>${compiler.plugin.version}</version>
        <configuration>
          <compilerArgs>
            <arg>-Xlint:all,-classfile</arg>
            <arg>-Werror</arg>
          </compilerArgs>
        </configuration>
      </plugin>
      <plugin>
        <groupId>org.apache.maven.plugins</groupId>
        <artifactId>maven-surefire-plugin</artifactId>
        <version>${surefire.plugin.version}</version>
        <configuration>
          <!-- no test run is a red build, never a green one: with no tests JaCoCo skips its
               check ("missing execution data file") and the coverage tiers pass vacuously -->
          <failIfNoTests>true</failIfNoTests>
        </configuration>
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
            <!-- "no wildcard imports" as a check, not a habit: google-java-format keeps them -->
            <forbidWildcardImports />
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
      <!-- mutation gate: CI-only (assets/ci-java.yml runs the changed classes, assets/mutation-java.yml the daily sweep), never in the hook; not bound to verify -->
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
          <targetClasses>${pitest.targetClasses}</targetClasses>
          <targetTests>
            <param>com.example.app.*</param>
          </targetTests>
          <!-- PIT's coverage pass runs every target test once, so exclude the ones that
               kill no domain or use-case mutant and cost the most: the resource tests
               (@QuarkusTest boots the application) and the ArchUnit test (reads bytecode,
               executes nothing) -->
          <excludedTestClasses>
            <param>com.example.app.api.*</param>
            <param>com.example.app.architecture.*</param>
          </excludedTestClasses>
          <mutationThreshold>90</mutationThreshold>
          <!-- XML feeds scripts/check-pit-score.py, which holds every class to the threshold:
               PIT's own mutationThreshold judges the whole run's score -->
          <outputFormats>
            <param>XML</param>
            <param>HTML</param>
          </outputFormats>
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
                <!-- a range, LATEST or RELEASE held in a version property is dynamic too; check-pom.sh is its fast echo -->
                <banDynamicVersions />
                <!-- rule 13: hand-written fakes implement the ports; no mock library, direct or transitive -->
                <bannedDependencies>
                  <message>No mock library in the pom, hand-written fakes implement the ports (rule 13)</message>
                  <excludes>
                    <exclude>org.mockito:*</exclude>
                    <exclude>org.easymock:*</exclude>
                    <exclude>org.powermock:*</exclude>
                    <exclude>org.jmockit:*</exclude>
                    <exclude>org.jmock:*</exclude>
                    <exclude>io.quarkus:quarkus-junit5-mockito</exclude>
                    <exclude>io.quarkus:quarkus-junit-mockito</exclude>
                    <exclude>io.quarkus:quarkus-panache-mock</exclude>
                  </excludes>
                </bannedDependencies>
              </rules>
            </configuration>
          </execution>
        </executions>
      </plugin>
      <!-- rule 35: cyclomatic complexity at most 10 per method; the ruleset ships in assets/java -->
      <plugin>
        <groupId>org.apache.maven.plugins</groupId>
        <artifactId>maven-pmd-plugin</artifactId>
        <version>${pmd.plugin.version}</version>
        <configuration>
          <rulesets>
            <ruleset>pmd-ruleset.xml</ruleset>
          </rulesets>
          <failOnViolation>true</failOnViolation>
          <printFailingErrors>true</printFailingErrors>
          <includeTests>false</includeTests>
          <linkXRef>false</linkXRef>
          <!-- rule 15: no inline ignore; an impossible marker makes every // NOPMD comment inert -->
          <suppressMarker>ATELIER-NEVER-SUPPRESS</suppressMarker>
        </configuration>
        <executions>
          <execution>
            <id>complexity</id>
            <phase>verify</phase>
            <goals><goal>check</goal></goals>
          </execution>
        </executions>
      </plugin>
    </plugins>
  </build>
</project>
```

### The Quarkus delta

A Quarkus service imports the platform BOM and declares its extensions without versions. Add the property and the `<dependencyManagement>` block, put the extensions first in `<dependencies>`, and make two edits to the canonical pom:

```xml
<!-- in <properties> -->
<quarkus.platform.version>3.39.5</quarkus.platform.version>

<!-- between </properties> and <dependencies> -->
<dependencyManagement>
  <dependencies>
    <dependency>
      <groupId>io.quarkus.platform</groupId>
      <artifactId>quarkus-bom</artifactId>
      <version>${quarkus.platform.version}</version>
      <type>pom</type>
      <scope>import</scope>
    </dependency>
  </dependencies>
</dependencyManagement>

<!-- first in <dependencies>: the extensions, no versions (the BOM pins them) -->
<dependency>
  <groupId>io.quarkus</groupId>
  <artifactId>quarkus-rest</artifactId>
</dependency>
<dependency>
  <groupId>io.quarkus</groupId>
  <artifactId>quarkus-junit</artifactId>
  <scope>test</scope>
</dependency>
```

- **Drop `<requireUpperBoundDeps />` from the enforcer.** The BOM is the convergence authority: it pins every version its extensions use, and its own tree does not pass the rule (on 3.39.5, `jctools-core` is managed at 4.0.5 while an extension asks for 4.0.6), so the first `validate` fails. `requireJavaVersion`, `requireReleaseDeps` and the mock ban stay.
- **Drop the JUnit pin** (`<junit.version>` and the `<version>` on `junit-jupiter`). The BOM manages JUnit, 6.x on 3.39, and a kept pin mixes a 5.x aggregator with 6.x modules. ArchUnit, PIT and its JUnit plugin run unchanged on it.

`scripts/smoke-test-java.sh` applies exactly this delta to the extracted canonical pom and proves it: the unmodified pom under the BOM fails `validate` on `RequireUpperBoundDeps`, a `@ConfigProperty` bean fails the compile under a plain `-Xlint:all`, and the delta compiles and passes `LayerRulesTest` and the tests. Bump `quarkus.platform.version` with the platform's releases; nothing else in the delta carries a version.

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
└── architecture/LayerRulesTest.java   # the dependency rule as an ArchUnit test (rule 37), a shipped asset
src/test/resources/
└── junit-platform.properties   # random method and class order (rule 36)
```

Dependency rule unchanged: `domain` imports nothing from the framework; `usecases` sees domain + its own ports; `infra` and `api` implement/consume them; only `composition` (and the CDI container) sees everything. The check is a test (hard rule 37): `assets/java/LayerRulesTest.java`, copied into `src/test/java/<pkg>/architecture/`, runs ArchUnit's layered architecture over the five packages plus two framework bans (domain sees no `jakarta`, `io.quarkus`, `org.hibernate`, `org.jboss`, `org.eclipse.microprofile`, `io.smallrye` or `io.vertx` class; use-cases see none of `jakarta.ws.rs`, `jakarta.persistence`, `io.quarkus`, `org.hibernate`, `org.eclipse.microprofile`, `io.smallrye`, `io.vertx`, so `@ApplicationScoped` on a use-case stays tolerated; prefer producing beans from `composition` when practical), test classes excluded, empty layers allowed for a walking skeleton. It runs in every `mvn test`, so `verify` and CI carry it. `grep -rn "import jakarta.ws.rs\|import io.quarkus" src/main/java/com/example/app/domain src/main/java/com/example/app/usecases` stays the adopt-mode audit for a tree that has no test yet.

## The hard rules, translated

| Rule (TS) | Java expression |
|:---|:---|
| 1 no `class` | Inverted mechanism, same intent: **records** for data, **sealed interfaces** for unions, small **final classes** with constructor injection for behaviour. No inheritance for reuse (composition only), no static mutable state, no field injection (`@Inject` on fields hides the contract; constructors state it) |
| 2 no `function` decl | n/a |
| 3 no `interface` | Inverted: interfaces ARE the port mechanism. Keep them small and role-shaped (ISP); one capability per interface |
| 4 no `console.*` | No `System.out`/`System.err`/`printStackTrace`. Inject a `Logger` (JBoss/SLF4J) through the constructor; redaction configured once (below). Gated: PMD `SystemPrintln` and the shipped `NoPrintStackTrace` XPath rule in `verify` |
| 5 Bun only | `./mvnw` only; the Quarkus CLI is sugar over it |
| 6 explicit return types | Native. Avoid `var` on any public or port surface; locals may use it when the right side names the type |
| 7-9 imports/style/ESM | Spotless owns style; no wildcard imports (`forbidWildcardImports` in the canonical pom, so `spotless:check` rejects one) |
| 10 no custom error classes | Business failures are `Err` values, never bespoke exception types. Exceptions are for bugs and framework edges; never use checked exceptions on domain surfaces |
| 11 TDD | Unchanged (JUnit 5) |
| 12 branded types | Value **records with validating compact constructors** plus a `parse(...)` factory returning `Result` (below) |
| 13 no `mock` | **No Mockito, no EasyMock, no `@InjectMock`.** Hand-written fakes implement the port interface; two gates keep the libraries out of the pom: the enforcer's `bannedDependencies` (direct or transitive, in every `mvn` run) and `check-pom.sh`'s third check in the fast hook (a declared mock coordinate) |
| 14 outside-in classicist | Unchanged: the SUT is the application service; domain runs real; only secondary ports get fakes |
| 15 zero warnings, no inline ignores | No `@SuppressWarnings`, ever. Compile with `-Xlint:all,-classfile -Werror` (the one lint off reports third-party class files, not this source); SonarJava/Error Prone severities change at project level with a comment. Three gates since 2026-09-08: `check-no-suppressions.sh` rejects `@SuppressWarnings`, `@SuppressFBWarnings`, `NOPMD`, `NOSONAR`, `CHECKSTYLE:OFF` and `noinspection` in staged Java (the fast hook) and across the tree (CI); the `NoSuppressWarnings` XPath rule in `pmd-ruleset.xml` flags the annotation in `verify` (it cannot see `@SuppressWarnings("PMD")`, which suppresses its own report, hence the grep); the canonical pom's `suppressMarker` is an impossible token, so a `// NOPMD` comment is inert |
| 16 `Result` at IO boundaries | `Result<T, E>` as a sealed interface (below); every port method returns it |
| 17 `try/catch` quarantine | Adapters in `infra/` catch SDK/JPA exceptions and translate to `Err`; use-cases pattern-match with `switch`; one top-level handler at the entry point |
| 18 no curried chains | n/a |
| 19 no `latest` | Exact versions, no ranges, no SNAPSHOT deps (above) |
| 20 Bun file API | `java.nio.file.Files` in `infra/` only; never `java.io.File` gymnastics in domain code. Gated: `LayerRulesTest` keeps `java.nio.file` and the `java.io` File classes out of `domain` and `usecases` |
| 21-22 design system | n/a (no UI) |
| 23-26 commits, tests, identity | Unchanged: same hooks, same confirmation gates |
| 27-34 production disciplines | Unchanged; Java expressions in their references and below |
| 35 cyclomatic complexity at most 10 | PMD `CyclomaticComplexity` with `methodReportLevel` 11 (PMD flags at or above the level) through `maven-pmd-plugin` bound to `verify`; the ruleset is `assets/java/pmd-ruleset.xml`, copied to the repo root. Never raise the level: split the method or dispatch on a map |
| 36 tests run in random order | `src/test/resources/junit-platform.properties` sets `MethodOrderer$Random` and `ClassOrderer$Random`, so every `mvn test` shuffles methods and classes; JUnit logs its seed below INFO, so `ci-java.yml` picks a seed, prints it, and passes it as `-Djunit.jupiter.execution.order.random.seed=<n>`, the same flag that replays a red order locally. No `@Order`, no `@TestMethodOrder(OrderAnnotation.class)`, no static state read across tests (Testing, Random order; the Java smoke test proves an order-dependent chain red) |
| 37 dependencies point inward | The shipped `LayerRulesTest` (ArchUnit, `archunit-junit5` in the canonical pom, test scope): a layered architecture over `domain`, `usecases`, `infra`, `api`, `composition` plus the two framework bans (domain sees no `jakarta`, `io.quarkus`, `org.hibernate`, `org.jboss`, `org.eclipse.microprofile`, `io.smallrye`, `io.vertx` class; use-cases none of `jakarta.ws.rs`, `jakarta.persistence`, `io.quarkus`, `org.hibernate`, `org.eclipse.microprofile`, `io.smallrye`, `io.vertx`), test classes excluded, empty layers allowed; red on the first import against the table in every `mvn test`, so `verify` and CI carry it |

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
  private static final Pattern SHAPE = Pattern.compile("^[^@\\s]+@[^@\\s]+$");

  public Email { if (!SHAPE.matcher(value).matches()) throw new IllegalArgumentException("email"); }

  public enum Error { MALFORMED }

  public static Result<Email, Error> parse(String raw) {
    return SHAPE.matcher(raw).matches() ? new Ok<>(new Email(raw)) : new Err<>(Error.MALFORMED);
  }
}

public record Money(long cents, Currency currency) {}   // integer minor units, never double
// instants are java.time.Instant (UTC) in the domain; ZoneId only at the presentation edge
```

At the HTTP edge, Bean Validation covers shape (`@Valid`, `@NotNull`, `@Positive` on wire DTOs); domain invariants live in the records. Wire DTOs never cross into use-cases: map DTO to domain at the resource, domain to response record on the way out (`references/architecture.md`, API shape and the three model boundaries).

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

- **Authenticated by default**: `quarkus.security.jaxrs.default-roles-allowed=**` in `application.properties` (`**` is any authenticated identity), so a REST endpoint without a security annotation refuses an anonymous caller; `@PermitAll` is the explicit, justified exception, and the health probes under `/q/health` are not REST resources, so they stay reachable. A permission set named `default` with only a `policy` does nothing: Quarkus registers a path permission only with `paths`, so the path-based form is `quarkus.http.auth.permission.authenticated.paths=/*` with `.policy=authenticated`, plus a `permit` set for `/q/health/*` (the longest path wins). Identity comes from OIDC (rule 33: never hand-rolled auth); rate limits and TLS are baseline, not per-route memory (`references/security.md`, One baseline).
- Resource-shaped endpoints, not screen-shaped (`references/architecture.md`, The backend is a client-agnostic API).
- The resource maps `Result` to HTTP: `Ok` to 200/201, domain-expected failures to their status, use-case `StepError` to 500 with a generic body (internals stay in the log with the trace id).
- **OpenAPI from the code**: MicroProfile OpenAPI annotations (`@Operation`, `@APIResponse`, example objects) so the published spec cannot drift (`references/governance.md`).
- Every network client the app opens has connect and per-request timeouts, bounded jittered retries (`@Retry(maxRetries = 3, jitter = 200, retryOn = IOException.class)` on the method that throws: a separate CDI bean wrapping the raw client call, which the adapter calls and translates to `Err` (rule 17). MicroProfile Fault Tolerance retries only on a thrown exception and only through the bean's proxy, so `@Retry` on the adapter method that returns `Err` never fires, the trap `references/result-type.md` names for a throw-based retry around a port that returns `Result`), and an `Idempotency-Key` where the operation is not naturally safe to repeat (rule 29).
- Personal data never in a `@QueryParam` or a log line (rule 27): user-typed search terms arrive in a `@Valid` POST body.

## Observability

Quarkus ships OpenTelemetry: enable it, add `@WithSpan` on application services (or wrap at composition), Micrometer counters/timers tagged by outcome for behaviour metrics, JSON logs carrying the trace id. Alerting and SLO discipline as in `references/observability.md`.

## Testing (rules 11, 13, 14; gates)

- **Unit tests are plain JUnit 5**, no `@QuarkusTest`, no container, sub-millisecond: the SUT is the application service with fakes injected by hand. `@QuarkusTest` boots the app and belongs to the integration ring only.
- **Integration tests** use `@QuarkusTest` + REST Assured against the real edge: the happy path, the bypass (`references/testing.md`, Bypass tests: wrong role is 403, missing token 401), the cross-tenant 404, and the forged-header seam (`references/isolation.md`). Testcontainers (or dev services) provide a real database; fixtures are synthetic (rule 34).
- **Test names are business scenarios**: `premiumCustomerGets20PercentOff`, `crossTenantReadIsNotFound`, `regressionEmptyCartTotalsToZero`.
- **Coverage tiers with JaCoCo**: 100% line on `domain` + `usecases`, 80% on `infra` + `api` + `composition`, enforced by per-package `<rule>` limits in the JaCoCo check goal so the build fails loudly, untested classes included in the denominator (the coverage-preload principle is native here: JaCoCo counts all classes in the module).
- **Mutation testing with PIT**: `mutationThreshold=90` on `domain` + `usecases` packages. CI runs it on the changed classes only (`scripts/pit-changed.sh`, every pull request and push, `-Dpitest.targetClasses=` narrowing the pom's default scope); the full sweep is the daily `assets/mutation-java.yml`, never a commit gate. Incremental history is NOT free in current PIT: 1.25.7 errors `History has been enabled but no history plugin has been installed/activated` for BOTH `withHistory` and explicit `historyInputFile`/`historyOutputFile` (verified via `smoke-test-java`), and the only history plugin is Arcmutate's commercial `+arcmutate_history`. So the free speed levers are the narrow target scope (`targetClasses`/`targetTests`), parallel `threads` (set in the pom, mutation results are thread-independent), and the narrow scope itself keeps the CI run cheap (the hook never runs PIT; the gate is CI-only by design); in a multi-module repo, scope PIT per module. If incremental speed becomes a hard requirement at scale, Arcmutate is the only supplier, which makes it a licence decision, not a library swap. PIT's `mutationThreshold` judges the whole run's score, so a weak class passed beside strong ones; `pit-changed.sh` and `mutation-java.yml` end with `scripts/check-pit-score.py`, which reads the XML report (the pom asks PIT for XML and HTML) and holds every class, its nested classes folded in, to the pom's threshold (2026-10-23). Same policy as Stryker: no per-file exclusions because tests feel awkward; tighten the test or refactor.
- **PIT on Quarkus**: no Quarkus-specific mutation tool exists; PIT plus `pitest-junit5-plugin` is the whole story, and the plugin (1.2.3+, needs Quarkus 3.22.x+) is the only Quarkus-aware piece. It auto-disables Quarkus's JaCoCo extension, the classic thing that broke PIT there. The scoping above is also what keeps this healthy: because `domain`/`usecases` are covered by plain JUnit 5 (not `@QuarkusTest`) and the canonical pom's `excludedTestClasses` keeps the `api` resource tests and the ArchUnit test out of PIT's coverage pass (which otherwise runs every target test once), PIT never runs over a container-boot test, so Quarkus's build-time augmentation never triggers the `tests did not pass without mutation` failure and no Quarkus container stands up per mutant. If you widen PIT onto `@QuarkusTest` classes, expect both that failure (patch with `avoidCallsTo` on `io.quarkus.*` plus test excludes on older plugin versions) and the per-mutant container cost; the atelier design avoids both by construction. Pin the Quarkus BOM at or above 3.22.x.
- **Evals for any LLM hole** gate the merge like PIT does (`references/ai.md`).

**Random order (rule 36).** `src/test/resources/junit-platform.properties`:

```properties
junit.jupiter.testmethod.order.default=org.junit.jupiter.api.MethodOrderer$Random
junit.jupiter.testclass.order.default=org.junit.jupiter.api.ClassOrderer$Random
```

Every `mvn test` then shuffles methods and classes. JUnit logs its seed at a level no default run shows, so the shipped `ci-java.yml` picks the seed, prints it, and passes it with `-Djunit.jupiter.execution.order.random.seed=<n>`; the same flag on a local `./mvnw verify` replays a failing order. No `@Order`, no `@TestMethodOrder(OrderAnnotation.class)`, no static state read across tests: each test builds its own fixture.

## Gates and hooks

Same git hooks as the Bun variant, shell only, wired with `git config core.hooksPath .githooks`, plus the CI workflow. Every artifact below ships in the skill's `assets/`; copy them with the block after the list, never hand-write:

- `assets/commit-msg`: the shipped Conventional Commits validator, unchanged (rule 23; it is dependency-free shell).
- `assets/pre-commit-java`: the fast gates only, seven of them: commit size (`scripts/check-commit-size.sh`, shared with the Bun variant, ≤10 files / ≤300 lines) → pom sanity (`scripts/check-pom.sh`: no version ranges anywhere, no `-SNAPSHOT` in `<parent>`/`<dependencies>`/`<plugins>` or in a version property, a single-module project's own dev version may be a SNAPSHOT (in a multi-module repo a child's `<parent>` would be that SNAPSHOT, so keep release versions there); no mock library declared, rule 13) → no inline suppression (`scripts/check-no-suppressions.sh`, rule 15: `@SuppressWarnings`, `NOPMD`, `NOSONAR` and the rest as text) → `gitleaks git --staged` → identity (`scripts/check-identity.sh`, rule 26) → the discipline tripwires (`scripts/check-disciplines.sh`, rules 27, 29, 30) → `./mvnw -q spotless:check`. A multi-minute hook trains `--no-verify` (canon 15.1, and 15.3), so `./mvnw verify` and PIT do not live here.
- `assets/check-disciplines.sh` and the three guards it runs (`check-pii-channels.sh`, `check-io-deadlines.sh`, `check-data-lifecycle.sh`; rules 27, 29, 30): core gates, hook gate 6 on the staged lines and `--all` in CI, all Java-aware (`@QueryParam`, `HttpClient` timeouts, hard deletes and destructive DDL). The isolation guard (`check-isolation-tests.sh`, rule 28) is opt-in where tenants or owners exist, since it demands a 404 test of every new `api/` route (`references/workflow.md`, Discipline tripwires).
- `assets/audit-java.yml`: the two watchdogs that are not gate material, the OWASP CVE scan and `check-skill-pin.sh` (a vendored standard is a dependency, `references/governance.md`; the workflow's `SKILL_PIN_UPSTREAM` env names the repository the whole vendored tree is compared against), on a daily schedule plus the pull requests that touch a pom or the vendored skill.
- `assets/pit-changed.sh`: the mutation step of CI, PIT on the classes that changed in the event's range and their nested classes (a module's sources are refused loudly: run PIT per module) (the pull request's base, or `github.event.before..HEAD` on a push, which the workflow exports; an unknown base fails loudly, no change in scope exits 0), plus uncommitted and untracked sources locally, where the default base is the branch's `@{push}` (the classes changed since your last push; `origin/main` on a branch never pushed), an optional check before CI's run over the pull request.
- `assets/mutation-java.yml`: the daily full PIT sweep over `domain` and `usecases` (`workflow_dispatch` on demand), the only run that measures the whole tree; a red run is a task, not a blocked merge.
- `assets/ci-java.yml`: the authoritative gate set, run on every push and pull request as the required merge check. Its first step re-runs the commit-msg validator over the pushed range (`scripts/check-commit-messages.sh`, so `--no-verify` cannot slip a message past the local hook), then the pom gate, a full-history `gitleaks git` (CI installs its own pinned copy; every call ignores inline allow comments), plus `./mvnw verify` (compile with `-Werror`, unit + integration tests, JaCoCo tier check, the PMD complexity cap of rule 35), and PIT mutation (≥90 on `domain`/`usecases`). The commit-size range check runs here too; the CVE scan does not.
- `assets/java/LayerRulesTest.java`: the rule 37 test (ArchUnit); copied into `src/test/java/<pkg>/architecture/` with its package and `@AnalyzeClasses` root renamed to yours; needs `archunit-junit5` in the pom, which the canonical pom carries. Since 2026-09-19 it carries rule 20 too: `domain` and `usecases` depend on nothing in `java.nio.file` and on none of the `java.io` File classes; five rules in all.
- `assets/check-no-suppressions.sh`: the rule 15 tripwire for Java, staged lines in the hook and `--all` in CI; the forms it rejects are listed in its header.
- `assets/check-identity.sh`: the rule 26 tripwire, hook gate 5 on the staged lines and `--all` in CI: a multi-word git name in both orders and the email, every author and committer in the history under `--all`, `IDENTITY_DENYLIST` for employers and clients; one-word names and noreply addresses are handles, CODEOWNERS and `.mailmap` exempt. A private repository commits `.github/visibility` reading `private`, and the gate then warns instead of blocking.
- `assets/java/pmd-ruleset.xml`: the rule 35 ruleset and the rule 15 `NoSuppressWarnings` XPath rule (`CyclomaticComplexity`, `methodReportLevel` 11, so complexity 11 and above fails and 10 passes, the same boundary as the TypeScript `complexity: ['error', 10]`), copied to the repository root where the canonical pom's `maven-pmd-plugin` reads it in `verify`. `smoke-test-java.sh` plants a complexity-11 method and sees `pmd:check` red, and a complexity-10 one green. Since 2026-09-19 it carries rule 4 too: PMD's `SystemPrintln` and the `NoPrintStackTrace` XPath rule (PMD's own `AvoidPrintStackTrace` is silent on 7.17).
- `assets/check-reply.py` and `assets/claude-settings.json`: the reply gate, a Claude Code Stop hook rather than a git hook or a CI step, that has the agent restate a reply breaking the Interaction section's mechanical rules (`references/workflow.md`, Reply gate); it needs python3.
- `assets/land-on-green.sh`: the `land` job of `ci-java.yml`, which rebase-merges a green pull request pinned to the tested head and refuses when `main` moved since CI ran (`references/workflow.md`, Landing on green).
- `assets/check-branches.sh`, `assets/check-host.py`, `assets/host-ruleset.json` and `assets/branches.yml`: the daily watchdog of hard rule 38, a branch whose work already landed or one more than a day off `main`, and a host whose default branch is not `main`, whose `main` is unprotected, or whose merges are not rebase-only with head-branch deletion; after the first push the owner runs `python3 scripts/check-host.py --apply` once, with an admin token (`references/workflow.md`, Branch lifecycle).

```bash
mkdir -p .githooks scripts .github/workflows
cp <skill>/assets/pre-commit-java        .githooks/pre-commit
cp <skill>/assets/commit-msg             .githooks/commit-msg
cp <skill>/assets/java/pmd-ruleset.xml   pmd-ruleset.xml
cp <skill>/assets/gitattributes          .gitattributes
cp <skill>/assets/check-commit-size.sh   scripts/check-commit-size.sh
cp <skill>/assets/check-pom.sh           scripts/check-pom.sh
cp <skill>/assets/check-no-suppressions.sh scripts/check-no-suppressions.sh
cp <skill>/assets/check-identity.sh       scripts/check-identity.sh
cp <skill>/assets/check-disciplines.sh    scripts/check-disciplines.sh
cp <skill>/assets/check-commit-messages.sh scripts/check-commit-messages.sh
cp <skill>/assets/check-commit-range.sh    scripts/check-commit-range.sh
cp <skill>/assets/land-on-green.sh        scripts/land-on-green.sh
cp <skill>/assets/check-branches.sh       scripts/check-branches.sh
cp <skill>/assets/check-host.py           scripts/check-host.py
cp <skill>/assets/host-ruleset.json       scripts/host-ruleset.json
cp <skill>/assets/branches.yml            .github/workflows/branches.yml
cp <skill>/assets/pit-changed.sh         scripts/pit-changed.sh
cp <skill>/assets/check-pit-score.py     scripts/check-pit-score.py
cp <skill>/assets/check-docs.sh          scripts/check-docs.sh
cp <skill>/assets/ci-java.yml            .github/workflows/ci.yml
cp <skill>/assets/mutation-java.yml      .github/workflows/mutation-java.yml
cp <skill>/assets/audit-java.yml         .github/workflows/audit-java.yml
cp <skill>/assets/check-skill-pin.sh     scripts/check-skill-pin.sh
cp <skill>/assets/check-reply.py         scripts/check-reply.py
mkdir -p .claude && cp <skill>/assets/claude-settings.json .claude/settings.json   # the reply gate; merge its Stop entry if the file exists

# Discipline tripwires. Three are core gates run by check-disciplines.sh (hook gate 6,
# CI --all): rules 27, 29, 30. The isolation guard (rule 28) is opt-in where tenants or
# owners exist: call it beside the wrapper in the hook and CI. All four ship Java detection.
cp <skill>/assets/check-pii-channels.sh   scripts/check-pii-channels.sh
cp <skill>/assets/check-io-deadlines.sh   scripts/check-io-deadlines.sh
cp <skill>/assets/check-data-lifecycle.sh scripts/check-data-lifecycle.sh
cp <skill>/assets/check-isolation-tests.sh scripts/check-isolation-tests.sh
chmod +x .githooks/pre-commit .githooks/commit-msg scripts/*.sh
git config core.hooksPath .githooks
```

CI (`assets/ci-java.yml`) re-runs the commit-message and pom gates, scans the full history with `gitleaks git`, then runs `spotless:check`, `./mvnw verify`, and PIT on the changed classes (`scripts/pit-changed.sh`); the full PIT sweep runs daily from `assets/mutation-java.yml`. The CVE scan and the vendored-standard check moved out to the scheduled `assets/audit-java.yml`, since both change independently of your diff, and where the repo deploys, the compose portability gate and deployment events (`references/delivery.md`).

## Bootstrap checklist (fresh Java repo)

1. `quarkus create app com.example:app` (or the Maven archetype); commit the wrapper; delete sample code.
2. Parent pom: start from the canonical `pom.xml` above (compiler `-Werror`, Spotless + google-java-format, JaCoCo tier rules, PIT `mutationThreshold=90` scoped to `domain`/`usecases`, enforcer with `requireJavaVersion`/`requireReleaseDeps`/`requireUpperBoundDeps`), apply The Quarkus delta (the pinned BOM and its extensions, less `requireUpperBoundDeps` and the JUnit pin), commit `.mvn/jvm.config`. Exact versions everywhere.
3. Scaffold packages: `domain`, `usecases/ports`, `infra`, `api`, `composition`; copy the shipped domain assets into `domain` rather than hand-writing them, then rename their package to your own groupId:
   ```bash
   mkdir -p src/main/java/<pkg>/domain src/test/java/<pkg>/architecture
   cp <skill>/assets/java/{Result,Ok,Err,Email}.java src/main/java/<pkg>/domain/
   cp <skill>/assets/java/LayerRulesTest.java src/test/java/<pkg>/architecture/
   # LayerRulesTest is the dependency rule as a test (rule 37): rename its package and @AnalyzeClasses root.
   # Result/Ok/Err are the sealed Result union (rule 16); Email is the value-record
   # exemplar (rule 12) to copy for Money, UserId, and every other domain primitive.
   ```
4. `src/test/resources/junit-platform.properties` with the two random orderers (rule 36), then `application.properties`: authenticated-by-default policy, OIDC config placeholders, OTel enabled, JSON logging with the redaction filter, datasource for the constrained runtime role.
5. Flyway: `src/main/resources/db/migration/V1__init.sql`; dev services or Testcontainers for the integration ring.
6. Test support: `testsupport` package with the first hand-written fakes (logger recorder, clock); **no Mockito in the pom** (the enforcer's `bannedDependencies` and `check-pom.sh` keep it out).
7. Hooks, CI and their scripts: copy the whole block in Gates and hooks above, every line of it. The hook runs identity and the discipline wrapper on every commit and CI runs them with `--all`, so none of the three core tripwires is optional (a missing script fails every commit); only calling `check-isolation-tests.sh` (rule 28) is opt-in, where tenants or owners exist. `pmd-ruleset.xml` goes to the repository root, where `verify` reads it. Then `git config core.hooksPath .githooks`; optional local `gitleaks` install (CI installs its own). Verify the pom gate once: `bash scripts/check-pom.sh`.
8. Walking skeleton: one use-case returning `Ok` through its port, its value record, its JUnit test (propose the test first, rule 24), one resource with its REST Assured test including the 401 case.
9. Verify green: `./mvnw spotless:check verify`, PIT on the skeleton, hooks reject a junk message and an oversized commit, and a planted domain class that imports a use-case fails `./mvnw test` on `LayerRulesTest` (rule 37); revert the plant.
10. `.claude/LESSONS.md` header; verify no scaffolded file names a person, an employer, or a client (rule 26); stage and propose the first commit (rule 25).

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
