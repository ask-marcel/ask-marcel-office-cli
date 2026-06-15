# Bun TypeScript Script Variant

Applies to repos that are plain Bun + TypeScript (no Next.js, no React, no Tailwind). Used for CLIs, integration scripts, Firebase Admin jobs, CSV/PDF processing, batch jobs.

Identifiable by `"module": "src/main.ts"` in `package.json` and the Clean Architecture layout `src/{domain,use-cases,infra,presenter,composition,test-helpers}` (see `references/architecture.md`).

## Runtime

- **Runtime**: Bun (`bun init`, Bun v1.2.2 or newer).
- **Package manager**: Bun. `bun.lock` is committed.
- **Module system**: ESM. `"type": "module"` and `moduleDetection: "force"`.
- **Entry point**: `src/main.ts` (`"module": "src/main.ts"` in package.json).
- **Install**: `bun install`.
- **Run**: `bun run src/main.ts`.
- **TypeScript**: peer dep `^5.0.0`.

Never call `node`, `tsc`, `ts-node`, `vite`, `npm`, `pnpm`, or `yarn`.

## `package.json`

Minimal skeleton:

```json
{
  "name": "<project>",
  "module": "src/main.ts",
  "type": "module",
  "scripts": {
    "start": "bun run src/main.ts",
    "lint": "eslint --cache",
    "lint:strict": "LINT_STRICT=1 eslint",
    "typecheck": "tsc --noEmit",
    "coverage": "bun run scripts/check-coverage.ts",
    "mutate": "stryker run",
    "mutate:changed": "bash scripts/mutate-changed.sh",
    "mutate:staged": "bash scripts/mutate-staged.sh"
  },
  "devDependencies": {
    "@eslint/js": "^9.28.0",
    "@stryker-mutator/core": "^9.6.1",
    "@types/bun": "^1.2.0",
    "eslint": "^9.28.0",
    "eslint-plugin-prettier": "^5.4.1",
    "eslint-plugin-security": "^3.0.1",
    "eslint-plugin-sonarjs": "^4.0.3",
    "eslint-plugin-unicorn": "^59.0.1",
    "typescript-eslint": "^8.33.1"
  },
  "peerDependencies": {
    "typescript": "^5.0.0"
  }
}
```

The version ranges above are sample pins as of writing. **Never use `"latest"` or `"*"`** (atelier hard rule 19, enforced by `scripts/check-package-json.sh` in pre-commit gate 2). To get the actual current latest of every dep on a fresh repo, run `bun install` first (resolving the ranges above), then `bun update` (which rewrites the `^X.Y.Z` ranges to the latest matching version) and commit the lockfile change. After that, every new package goes in via `bun add <pkg>` (runtime) or `bun add -d <pkg>` (dev) — Bun pins it to `^X.Y.Z` automatically. Hand-editing `package.json` to add a dep is a smell.

Common runtime deps in this class of repo (install on demand with `bun add <name>`):
`@google/genai`, `axios`, `canvas`, `chardet`, `csv-writer`, `firebase-admin`, `iconv-lite`, `jsonwebtoken` (+ `@types/jsonwebtoken`), `papaparse`, `pdf-extract-image`, `pdf-to-png-converter`, `pdfjs-dist`, `winston`, `xlsx`.

## `tsconfig.json`

```jsonc
{
  "compilerOptions": {
    "lib": ["ESNext", "DOM"],
    "target": "ESNext",
    "module": "ESNext",
    "moduleDetection": "force",
    "jsx": "react-jsx",
    "allowJs": true,
    "moduleResolution": "bundler",
    "allowImportingTsExtensions": true,
    "verbatimModuleSyntax": true,
    "noEmit": true,
    "strict": true,
    "skipLibCheck": true,
    "noFallthroughCasesInSwitch": true,
    "noUnusedLocals": false,
    "noUnusedParameters": false,
    "noPropertyAccessFromIndexSignature": false,
    "types": ["bun"]
  }
}
```

`"types": ["bun"]` is required so VS Code's TypeScript server resolves `import ... from 'bun:test'`. The CLI `tsc --noEmit` works either way through type-acquisition heuristics, but the editor needs the explicit list. After adding this to an existing project, restart the TS server in VS Code (Cmd/Ctrl + Shift + P → "TypeScript: Restart TS Server").

Notes:

- `strict: true` covers `noImplicitAny`, `strictNullChecks`, `strictFunctionTypes`, `strictPropertyInitialization`, `alwaysStrict`, `useUnknownInCatchVariables`.
- `verbatimModuleSyntax: true` forces `import type` for type-only imports.
- `allowImportingTsExtensions: true` + `moduleResolution: "bundler"` allow `import ... from './foo.ts'` directly (idiomatic in Bun).
- `noEmit: true`: TypeScript is type-check only; Bun handles execution.

## `eslint.config.js`

Flat config, ESM, filename is `.js` (not `.mjs`) in this variant.

```js
import pluginJs from '@eslint/js';
import prettier from 'eslint-plugin-prettier';
import securityPlugin from 'eslint-plugin-security';
import sonarjsPlugin from 'eslint-plugin-sonarjs';
import unicornPlugin from 'eslint-plugin-unicorn';
import globals from 'globals';
import tsPlugin from 'typescript-eslint';

/** @type {import('eslint').Linter.Config[]} */
export default [
  pluginJs.configs.recommended,
  ...tsPlugin.configs.recommended,
  securityPlugin.configs.recommended,
  {
    files: ['**/*.ts'],
    languageOptions: { globals: globals.node },
    rules: {
      'func-style': ['error', 'expression'],
      'no-console': ['error'],
      'prefer-template': 'error',
      quotes: ['error', 'single', { avoidEscape: true }],
      // `mock` from `bun:test` is process-global once installed and leaks into
      // every other test file the runner loads. Use dependency injection
      // (createXFromApi or installFetchMock) instead. See references/testing-infra.md.
      'no-restricted-imports': ['error', {
        paths: [{
          name: 'bun:test',
          importNames: ['mock'],
          message:
            '`mock` from bun:test is forbidden — it leaks across test files. Use dependency injection: refactor the production code to accept the SDK as a parameter, then pass a fake at construction.',
        }],
      }],
      '@typescript-eslint/explicit-function-return-type': ['error', { allowExpressions: true, allowTypedFunctionExpressions: true }],
      '@typescript-eslint/consistent-type-definitions': ['error', 'type'],
    },
  },
  // Type-aware rules — slow (~25s on full repo), enabled only by
  // `bun run lint:strict` (which sets LINT_STRICT=1) and the pre-commit hook.
  // Inner-loop `bun run lint` does NOT run them.
  ...(process.env['LINT_STRICT']
    ? [
        {
          files: ['src/**/*.ts'],
          languageOptions: {
            parserOptions: {
              projectService: true,
              tsconfigRootDir: import.meta.dirname,
            },
          },
          rules: {
            // Lint-time equivalents of Sonar S4325 (no `!`/`as` non-narrowing assertions)
            // and S6671 (Promise.reject must be an Error).
            '@typescript-eslint/no-unnecessary-type-assertion': 'error',
            '@typescript-eslint/prefer-promise-reject-errors': 'error',
          },
        },
      ]
    : []),
  {
    plugins: { prettier },
    rules: {
      'prettier/prettier': [
        1,
        {
          endOfLine: 'lf',
          printWidth: 180,
          semi: true,
          singleQuote: true,
          tabWidth: 2,
          trailingComma: 'es5',
        },
      ],
    },
  },
  {
    plugins: { unicorn: unicornPlugin },
    rules: {
      'unicorn/empty-brace-spaces': 'off',
      'unicorn/no-null': 'off',
    },
  },
  {
    rules: {
      // false-positive-heavy rules in this codebase's idioms; disabled at project level.
      // Never inline-ignore — change severity here or refactor the code.
      'security/detect-object-injection': 'off',
      'security/detect-unsafe-regex': 'off',
      // detect-non-literal-fs-filename flags `chmodSync(mkdtempSync(...))` in FS-adapter
      // tests. Production code uses Bun.file (not flagged by this rule), so disabling
      // globally loses nothing on the real attack surface.
      'security/detect-non-literal-fs-filename': 'off',
    },
  },
  sonarjsPlugin.configs.recommended,
  {
    // SonarJS rule overrides — always-on, justified per rule. See LESSONS.md.
    rules: {
      'sonarjs/no-unused-vars': 'off',          // duplicates @typescript-eslint/no-unused-vars
      'sonarjs/no-empty-test-file': 'off',      // false positives on `describe` test layout
      'sonarjs/cognitive-complexity': 'off',    // we already cap function size; this is noise
    },
  },
  // Stryker scratch dirs, the reports/ output dir, and other non-source paths
  // must not be linted. Stryker copies the source tree into .stryker-tmp/ during a run.
  {
    ignores: ['.stryker-tmp/**', 'reports/**', 'docs/**', 'scripts/**', '.claude/**', '.agents/**'],
  },
];
```

Notes on the config:

- **One config file, two modes.** The inner-loop `bun run lint` runs only the fast non-type-aware rules (~2 s cached / ~7 s cold). `bun run lint:strict` sets `LINT_STRICT=1` and the conditional block adds `parserOptions.projectService: true` plus the type-aware `@typescript-eslint` rules (~25 s on a full repo). Pre-commit gate 4 runs the strict version. There is no separate `eslint.strict.config.js` — keeping one config eliminates drift.
- **`sonarjsPlugin.configs.recommended`** catches SonarLint findings at lint time so they no longer escape the IDE. See `references/workflow.md` for the common ones (S4325, S6594, S4123, S6551, S6671). Three rules are turned off as always-on noise: `sonarjs/no-unused-vars` (duplicate), `sonarjs/no-empty-test-file` (false-positive on `describe` blocks), `sonarjs/cognitive-complexity` (we already cap function size).
- **`no-console` is `error`**. Always use the logger port (see below), never `console.*`.
- **`security/detect-object-injection`, `detect-unsafe-regex`, and `detect-non-literal-fs-filename`** are disabled at the project level because they only false-positive on this codebase's idioms (branded-type `Record<K, V>` lookups, bounded regexes, `chmodSync(mkdtempSync(...))` in tests). Comments in the config explain why each is off. Never inline-ignore them per-line.
- **`no-restricted-imports`** blocks `mock` from `bun:test` (the entire namespace) — see hard rule 13.

See `references/workflow.md` for the zero-warning rule and the no-inline-ignore discipline that make this config load-bearing.

## `.vscode/settings.json`

```json
{
  "editor.formatOnSave": true,
  "editor.codeActionsOnSave": {
    "source.organizeImports": "explicit",
    "source.addMissingImports": "explicit",
    "source.fixAll.eslint": "explicit"
  },
  "search.exclude": { "**/node_modules": true, "**/.vscode": true },
  "search.useGlobalIgnoreFiles": true,
  "search.useParentIgnoreFiles": true,
  "git.autofetch": true,
  "editor.trimAutoWhitespace": true,
  "files.encoding": "utf8",
  "files.trimFinalNewlines": true,
  "files.trimTrailingWhitespace": true,
  "editor.quickSuggestions": { "strings": true },
  "editor.detectIndentation": false,
  "editor.tabSize": 2,
  "eslint.enable": true,
  "eslint.format.enable": true,
  "editor.defaultFormatter": "dbaeumer.vscode-eslint",
  "editor.formatOnType": true,
  "typescript.format.insertSpaceAfterOpeningAndBeforeClosingEmptyBraces": false,
  "[typescript]": { "editor.defaultFormatter": "vscode.typescript-language-features" },
  "[javascript]": { "editor.defaultFormatter": "vscode.typescript-language-features" },
  "[jsonc]": { "editor.defaultFormatter": "vscode.json-language-features" },
  "[json]": { "editor.defaultFormatter": "vscode.json-language-features" }
}
```

## `.vscode/extensions.json`

```json
{
  "recommendations": [
    "dbaeumer.vscode-eslint",
    "eamodio.gitlens",
    "usernamehw.errorlens",
    "lacroixdavid1.vscode-format-context-menu",
    "kisstkondoros.vscode-codemetrics",
    "snyk-security.snyk-vulnerability-scanner",
    "sonarsource.sonarlint-vscode"
  ]
}
```

## `.gitignore`

Base: GitHub's Node.gitignore, used unmodified. Covers logs, caches, diagnostic reports, coverage, `node_modules`, `.env*`, build output, IDE/OS files, Yarn v2.

If you use Firebase Admin, add `*-service-account*.json` to keep credentials out of git.

## Source architecture

The canonical shape for any non-trivial backend — a pipeline, batch job, or CLI with real integrations — is the Clean Architecture layout: `src/{domain,use-cases,infra,presenter,composition,test-helpers}` + `src/main.ts`. See `references/architecture.md` for the full layout, the dependency matrix, and the "adding a new external service" recipe.

For throwaway scripts, one-off CLIs, or prototypes with a single integration, a flat `src/main.ts` plus a thin `src/utils/` is fine. Graduate to the Clean Architecture layout once you have a second external service, real tests, or ~500 lines of production code.

## What not to create

- No top-level `lib/`, `types/`, or path aliases.
- Types live next to the code that uses them. Prefer `type` over `interface` (enforced).
- No `interface`, no `class`, no `function` declarations, no custom error classes.

## Naming

- Folders and filenames: `kebab-case` (e.g. `find-url-not-secure/`, `auth-token.ts`).
- Functions, variables: `camelCase` (e.g. `getAuthTokens`, `findFilesToScan`).
- Types: `PascalCase` (e.g. `SearchResult`, `TokenResponse`).

## Imports

- ESM only (`import ... from ...`, `export const ...`).
- `.ts` extensions in import specifiers are allowed and idiomatic.
- Bun-specific APIs (`Bun.file()`, `import.meta.dir`) are fine.

## Testing

No tests today. If you add them:

- Filename convention: `*.test.ts` next to the source.
- Runner: `bun test`.

## Secrets & config hygiene

No credentials in source. Load every env var through the `envVar` branded-type factory, centralised in a `config/env.ts` per feature. `.env*` is git-ignored. For Firebase Admin, add `*-service-account*.json` to `.gitignore` and load the path via env var, never commit the JSON.

See `references/security.md` for the full pattern: `envVar` factory, redacted Winston logger, never-sprinkle-`process.env` rule, and the list of what must never be committed.

## Logger (port + adapter + fake, not a module singleton)

The logger is a side-effectful dependency, so it gets port/adapter separation like every other IO dependency. Three files:

```ts
// src/use-cases/ports/logger.ts — type only, zero dependencies
export type LogMeta = Readonly<Record<string, unknown>>;

export type Logger = {
  readonly info: (event: string, meta?: LogMeta) => void;
  readonly warn: (event: string, meta?: LogMeta) => void;
  readonly error: (event: string, meta?: LogMeta) => void;
};
```

```ts
// src/infra/logger.ts — Winston-backed adapter, real for production
import { createLogger, format, transports } from 'winston';
import type { Logger } from '../use-cases/ports/logger.ts';

const REDACTED_KEYS = new Set(['password', 'token', 'authorization', 'apikey', 'secret']);

const redactFormat = format((info) => {
  for (const key of Object.keys(info)) {
    if (REDACTED_KEYS.has(key.toLowerCase())) info[key] = '[REDACTED]';
  }
  return info;
});

export const createWinstonLogger = (): Logger => {
  const winston = createLogger({
    level: process.env.LOG_LEVEL ?? 'info',
    format: format.combine(redactFormat(), format.json()),
    transports: [new transports.Console()],
  });
  return {
    info: (event, meta) => winston.info(event, meta),
    warn: (event, meta) => winston.warn(event, meta),
    error: (event, meta) => winston.error(event, meta),
  };
};
```

```ts
// src/test-helpers/logger-fake.ts — in-memory fake for tests, logs become assertable
import type { Logger, LogMeta } from '../use-cases/ports/logger.ts';

export type LoggerFake = Logger & {
  readonly calls: ReadonlyArray<{ readonly level: 'info' | 'warn' | 'error'; readonly event: string; readonly meta?: LogMeta }>;
};

export const createLoggerFake = (): LoggerFake => {
  const calls: { level: 'info' | 'warn' | 'error'; event: string; meta?: LogMeta }[] = [];
  return {
    calls,
    info: (event, meta) => { calls.push({ level: 'info', event, meta }); },
    warn: (event, meta) => { calls.push({ level: 'warn', event, meta }); },
    error: (event, meta) => { calls.push({ level: 'error', event, meta }); },
  };
};
```

Every use-case declares `readonly logger: Logger` in its `Deps` and calls `deps.logger.info(...)`. Composition wires `createWinstonLogger()` in `src/composition/build-deps.ts`. Tests inject `createLoggerFake()` and assert on the `calls` array — logs become assertable without a mocking library.

Why this and not a module-level singleton: a singleton makes the logger impossible to swap in tests without monkey-patching, and impossible to redact/reformat per-environment without mutating global state. A port is one extra type declaration and pays off the first time you want to assert that a warning fired, or run a test suite in silent mode.

Invariant: `grep -rn "from '.*infra" src/domain src/use-cases` must return nothing. The domain and use-cases know only about the `Logger` **type**; the Winston import lives in `infra/` only.

## Error handling

No `try/catch` anywhere outside `src/infra/**`, pure-domain fallbacks for native-synchronous throwers (`JSON.parse`, `URL` constructor), and exactly one top-level handler in `src/main.ts`. Every IO port returns `Promise<Result<T, PortError>>`. Use-cases pattern-match on `.ok` and aggregate port errors into `StepError`. See `references/result-type.md` for the full treatment, the discriminated-union error design, the fan-out batch semantics, and the `retryOnErr` + `captureRejection` helpers.

The shared `formatError(err: unknown): string` helper lives in `src/domain/utilities/format-error.ts`. Use it in every `catch (e)` block in `src/infra/**` — never `String(e)`, which returns `"[object Object]"` for non-Error throws (SonarJS S6551).

`process.exit(1)` is allowed only in `src/main.ts` after the top-level catch. Never inside a use-case, adapter, or domain module.

## Bootstrap checklist (fresh Bun repo)

1. `mkdir <new-repo> && cd <new-repo> && bun init -y`.
2. Replace `package.json` with the skeleton above (devDependencies include `eslint-plugin-sonarjs`; scripts include `lint`, `lint:strict`, `typecheck`, `coverage`, `mutate`, `mutate:changed`, `mutate:staged`, `start`). **No `"latest"` or `"*"` anywhere** — the skeleton's `^X.Y.Z` ranges are samples; bump them in step 7 below.
3. Create `tsconfig.json` with the block above (includes `"types": ["bun"]`).
4. Create `eslint.config.js` with the flat config above (includes `sonarjs.configs.recommended` and type-aware `@typescript-eslint` rules behind `LINT_STRICT=1`).
5. Create `.vscode/settings.json` and `.vscode/extensions.json`.
6. Drop in a Node `.gitignore`, plus `*-service-account*.json` if Firebase is in play.
7. `bun install` to resolve the skeleton's ranges; then `bun update` to bump every dep to its current latest matching version. Commit `bun.lock` and the updated `package.json` together. From this point, every new dep is added via `bun add <pkg>` (runtime) or `bun add -d <pkg>` (dev) — never hand-edit `package.json`.
8. Scaffold the Clean Architecture layout (see `references/architecture.md`): `mkdir -p src/{domain,use-cases/ports,infra,presenter,composition,test-helpers}`.
9. Create `src/domain/result.ts` with the `Result<T, E>` type and helpers from `references/result-type.md`.
10. Create `src/use-cases/ports/logger.ts` and `src/infra/logger.ts` with the port + Winston adapter above; create `src/test-helpers/logger-fake.ts`.
11. Copy the canonical helpers from the skill's `assets/`:
    - `cp <skill-path>/assets/format-error.ts src/domain/utilities/format-error.ts`
    - `cp <skill-path>/assets/capture-rejection.ts src/test-helpers/capture-rejection.ts`
    - `cp <skill-path>/assets/fetch-mock.ts src/test-helpers/fetch-mock.ts`
12. Set up the per-tier coverage gate:
    - `cp <skill-path>/assets/check-coverage.ts scripts/check-coverage.ts`
    - `cp <skill-path>/assets/coverage-preload.ts scripts/coverage-preload.ts` (edit the `TODO: add every adapter here` list as you grow)
    - In `bunfig.toml`: `[test]` section with `coverage = true`, `coverageSkipTestFiles = true`, `coverageReporter = ["text"]`. **Do not** add `coverageThreshold` (the per-tier script owns enforcement) and **do not** add `preload` (the coverage preload is loaded only by `scripts/check-coverage.ts` via `--preload` so plain `bun test` runs stay fast).
13. Set up mutation testing:
    - `cp <skill-path>/assets/stryker.conf.json stryker.conf.json`
    - `cp <skill-path>/assets/mutate-staged.sh scripts/mutate-staged.sh`
    - `cp <skill-path>/assets/mutate-changed.sh scripts/mutate-changed.sh`
    - `chmod +x scripts/*.sh`
    - Add to `.gitignore`: `.stryker-tmp/` and `reports/` (Stryker scratch + output dirs).
14. Install the pre-commit hook (eight gates):
    - `cp <skill-path>/assets/check-commit-size.sh scripts/check-commit-size.sh`
    - `cp <skill-path>/assets/check-package-json.sh scripts/check-package-json.sh`
    - `chmod +x scripts/check-commit-size.sh scripts/check-package-json.sh`
    - `mkdir -p .githooks && cp <skill-path>/assets/pre-commit .githooks/pre-commit && chmod +x .githooks/pre-commit`
    - `git config core.hooksPath .githooks`
    - Optional: `brew install gitleaks` (macOS) or grab a binary from `github.com/gitleaks/gitleaks/releases`. The hook degrades gracefully if missing.
    - See `references/workflow.md` for the eight-gate breakdown and the no-bypass rule.
15. Verify: `bun run lint`, `bun run typecheck`, `bun run coverage`, and `bun run mutate` all clean on a minimal `src/main.ts`. Run `bash scripts/check-package-json.sh` once to confirm no `"latest"` slipped in.
16. Commit with Conventional Commits; from here, follow the Clean Architecture rules for every new feature.
