# Bun TypeScript Script Variant

Applies to repos that are plain Bun + TypeScript (no Next.js, no React, no Tailwind). Used for CLIs, integration scripts, Firebase Admin jobs, CSV/PDF processing, batch jobs.

Identifiable by `"module": "src/main.ts"` in `package.json` and the Clean Architecture layout `src/{domain,use-cases,infra,presenter,composition,test-helpers}` (see `references/architecture.md`).

## Runtime

- **Runtime**: Bun (`bun init`, Bun v1.4 or newer: the TypeScript aliases below need 1.4, and `bun test --randomize` prints its seed and `--seed` replays it from 1.3, which rule 36 relies on).
- **Package manager**: Bun. `bun.lock` is committed.
- **Module system**: ESM. `"type": "module"` and `moduleDetection: "force"`.
- **Entry point**: `src/main.ts` (`"module": "src/main.ts"` in package.json).
- **Install**: `bun install`.
- **Run**: `bun run src/main.ts`.
- **TypeScript**: two devDependency aliases. `typescript` is `@typescript/typescript6`, the 6.0 API that typescript-eslint and eslint-plugin-sonarjs load. `@typescript/native` is TypeScript 7 and gives the `tsc` binary. TypeScript 7.0 has no JavaScript API, so ESLint cannot load it; the weekly canary re-probes it. The aliases need Bun 1.4: Bun 1.3.14 resolves the wrapper's own `@typescript/old` dependency (`npm:typescript@^6`) to the wrapper itself, so `require('typescript')` returns an empty object and ESLint crashes, and a `bun.lock` that Bun 1.3 wrote keeps that entry under Bun 1.4 until its `@typescript/old` line is deleted.

Never call `node`, `ts-node`, `vite`, `npm`, `pnpm`, or `yarn`. TypeScript's compiler runs as the type checker only (`tsc --noEmit`, the `typecheck` script), never to build or run.

**Server archetype.** The default shape is a CLI/batch job that runs and `process.exit`s, but `src/main.ts` may instead call `Bun.serve` to serve HTTP. The inbound server is then an `infra/` adapter (the mirror of an outbound one), `main.ts` stays the single entry with its one top-level catch, and the Dockerfile gains `EXPOSE <port>`. See `references/architecture.md` § Inbound HTTP (server archetype).

## `package.json`

Minimal skeleton:

```json
{
  "name": "<project>",
  "module": "src/main.ts",
  "type": "module",
  "scripts": {
    "start": "bun run src/main.ts",
    "lint": "eslint --cache --max-warnings=0",
    "lint:strict": "LINT_STRICT=1 eslint --max-warnings=0",
    "lint:staged": "bash scripts/lint-staged.sh",
    "test": "bun test --randomize",
    "typecheck": "tsc --noEmit",
    "coverage": "bun run scripts/check-coverage.ts",
    "mutate": "stryker run && bun run scripts/check-mutation-score.ts",
    "mutate:changed": "bash scripts/mutate-changed.sh",
    "mutate:staged": "bash scripts/mutate-staged.sh"
  },
  "devDependencies": {
    "@eslint/js": "^10.0.1",
    "@stryker-mutator/core": "^10.0.0",
    "@types/bun": "^1.4.2",
    "@typescript/native": "npm:typescript@^7.0.2",
    "eslint": "^10.11.0",
    "eslint-plugin-prettier": "^5.5.6",
    "eslint-plugin-security": "^4.1.0",
    "eslint-plugin-sonarjs": "^4.2.1",
    "eslint-plugin-unicorn": "^76.0.0",
    "globals": "^17.12.0",
    "prettier": "^3.9.9",
    "typescript": "npm:@typescript/typescript6@^6.0.2",
    "typescript-eslint": "^8.70.1"
  }
}
```

The version ranges above are sample pins as of writing, with one that is load-bearing: the two TypeScript aliases. Keep them until typescript-eslint and eslint-plugin-sonarjs support TypeScript 7.1 (TypeScript 7.0 has no JavaScript API: typescript-eslint 8.71.2 refuses it and sonarjs 4.2.2 crashes in ts-api-utils; found 2026-10-18). Never run `bun add typescript@latest`: it replaces the alias. **Never use `"latest"` or `"*"`** (atelier hard rule 19, enforced by `scripts/check-package-json.sh` in pre-commit gate 2). To get the actual current latest of every dep on a fresh repo, run `bun install` first (resolving the ranges above), then `bun update` (which rewrites the `^X.Y.Z` ranges to the latest matching version) and commit the lockfile change. After that, every new package goes in via `bun add <pkg>` (runtime) or `bun add -d <pkg>` (dev): Bun pins it to `^X.Y.Z` automatically. Hand-editing `package.json` to add a dep is a smell.

Common runtime deps in this class of repo (install on demand with `bun add <name>`):
`@google/genai`, `canvas`, `chardet`, `csv-writer`, `firebase-admin`, `iconv-lite`, `jsonwebtoken` (+ `@types/jsonwebtoken`), `papaparse`, `pdf-extract-image`, `pdf-to-png-converter`, `pdfjs-dist`, `winston`, `xlsx`.

For HTTP, reach for the native `fetch` (built into Bun) before adding an HTTP client: the lazy ladder's rung 3 (SKILL.md #2) and the whole `installFetchMock` test seam assume adapters call `globalThis.fetch` directly. Add `axios`/`got`/etc. only when you need something `fetch` genuinely lacks, and say what.

## `tsconfig.json`

```jsonc
{
  "compilerOptions": {
    "lib": ["ESNext", "DOM"],
    "target": "ESNext",
    "module": "ESNext",
    "moduleDetection": "force",
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

// `mock` from `bun:test` is process-global once installed and leaks into every
// other test file the runner loads, and `spyOn`, `jest` and `vi` are the same
// mocking machinery under other names. Use dependency injection (createXFromApi or
// installFetchMock) instead. See references/testing-infra.md. A const, because
// every layer zone below has to repeat it (see the replace-not-merge note).
const MOCK_BAN = {
  name: 'bun:test',
  importNames: ['mock', 'spyOn', 'jest', 'vi'],
  message:
    '`mock`, `spyOn`, `jest` and `vi` from bun:test are forbidden: mocks leak across test files. Use dependency injection: refactor the production code to accept the SDK as a parameter, then pass a fake at construction (hard rule 13).',
};

// The style rules as lint (hard rules 1, 7, 10, 18, and 13's call-recording assertions).
// ESLint REPLACES a rule's options
// when a second block matches the same file, it never merges them, so every scoped
// `no-restricted-syntax` block below spreads this list before its own selector.
const STYLE_BANS = [
  { selector: 'ClassDeclaration', message: 'No class keyword (hard rule 1; rule 10 for error classes): a module of arrow functions and typed records (references/design-patterns.md).' },
  { selector: 'ClassExpression', message: 'No class keyword (hard rule 1): a module of arrow functions and typed records (references/design-patterns.md).' },
  { selector: 'ImportSpecifier[importKind="type"]', message: 'Type-only imports on their own line: `import type { Foo } from ...` (hard rule 7).' },
  {
    selector: 'VariableDeclarator[id.name!=/^create[A-Z]/] > ArrowFunctionExpression > ArrowFunctionExpression.body',
    message: 'No curried arrow chains: one arrow with all its parameters, wrapped at the call site; the DI factory `createX = (deps) => (input) => ...` is the one exemption (hard rule 18, references/clean-code.md).',
  },
  {
    selector: 'MemberExpression[property.name=/^(toHaveBeen(Last|Nth)?Called|toBeCalled|toHave(Last|Nth)?Returned)/]',
    message: 'No call-recording assertions: a hand-written fake records what it received and the test asserts on that record (hard rule 13, references/testing.md).',
  },
];
// Hard rule 17: a use-case pattern-matches the Result a port returns; a catch there means the port lied.
const TRY_BAN = {
  selector: 'TryStatement',
  message: 'try/catch is quarantined to src/infra/**, the pure-domain fallback and src/main.ts; a use-case pattern-matches the Result (hard rule 17, references/result-type.md).',
};
// Hard rule 20: file IO is Bun.file / Bun.write; node:fs only in tests, src/test-helpers/**
// and the one commented directory helper in src/infra/** (File IO, below).
const FS_BAN = {
  selector: 'ImportDeclaration[source.value=/^(node:)?fs(\\/promises)?$/]',
  message: 'File IO goes through Bun.file and Bun.write; node:fs only in tests, src/test-helpers/** and the one commented directory helper in src/infra/** (hard rule 20, references/bun-typescript.md, File IO).',
};

// A `<factory>Unsafe` helper casts a brand without validating (references/testing.md,
// Branded types and `expect(...).toBe(raw)`); only tests and the fakes may import one.
// It rides inside the zones because flat config replaces a rule's options: a separate
// no-restricted-imports block would wipe the mock ban and the zone for its files.
const INFRA_NO_USE_CASE = {
  group: ['**/use-cases/*', '!**/use-cases/ports'],
  message: 'src/infra implements the use-case ports and never imports a use-case: dependencies point inward (hard rule 37, references/architecture.md, the dependency table).',
};
const UNSAFE_BAN = {
  group: ['**'],
  importNamePattern: 'Unsafe$',
  message: '*Unsafe helpers skip validation and are test-only: production code builds the value through its factory (references/testing.md, Branded types).',
};

// The dependency rule as lint (hard rule 37; references/architecture.md, the dependency
// table). Dependencies point inward, so each layer names the layers it may never import,
// and a layer left out of a list is one it is allowed to reach. Production files only: a
// test reaches for the fakes in src/test-helpers/ by design, and the fakes may build
// their values with the *Unsafe helpers, so that zone alone leaves UNSAFE_BAN out.
// `extraPatterns` carries a ban a name list cannot say, like infra's below.
const layerZone = (layer, forbidden, files = [`src/${layer}/**/*.ts`], extraPatterns = []) => ({
  files,
  ignores: ['**/*.test.ts'],
  rules: {
    'no-restricted-imports': [
      'error',
      {
        paths: [MOCK_BAN],
        patterns: [
          {
            group: forbidden.flatMap((name) => [`**/${name}`, `**/${name}/**`]),
            message: `src/${layer} must not import ${forbidden.join(', ')}: dependencies point inward (hard rule 37, references/architecture.md, the dependency table).`,
          },
          ...extraPatterns,
          ...(layer === 'test-helpers' ? [] : [UNSAFE_BAN]),
        ],
      },
    ],
  },
});

/** @type {import('eslint').Linter.Config[]} */
export default [
  pluginJs.configs.recommended,
  ...tsPlugin.configs.recommended,
  securityPlugin.configs.recommended,
  {
    // Hard rule 15: no inline ignore, ever. Directive comments (eslint-disable*, eslint-enable,
    // globals, exported) are inert AND each one is reported, so `--max-warnings=0` fails on the
    // comment and the violation it hid surfaces beside it. A finding is a refactor or a
    // project-level severity change with a reason, never a suppression.
    linterOptions: { noInlineConfig: true },
  },
  {
    files: ['**/*.ts'],
    languageOptions: { globals: globals.node },
    rules: {
      'func-style': ['error', 'expression'],
      // Rule 35: cyclomatic complexity at most 10 per function. The size caps
      // (10 lines, one indentation level) miss a one-line chain of `&&`/`??`/
      // ternaries and a wide `switch`; this counts the branches.
      complexity: ['error', 10],
      'no-console': ['error'],
      'prefer-template': 'error',
      quotes: ['error', 'single', { avoidEscape: true }],
      'no-restricted-imports': ['error', { paths: [MOCK_BAN] }],
      // Hard rules 1, 7, 10, 18 and 13's call assertions (STYLE_BANS above); the scoped blocks below add 17 and 20.
      'no-restricted-syntax': ['error', ...STYLE_BANS],
      // Hard rule 15, the other tools' escape hatches: every @ts- form (the recommended preset
      // allows a described @ts-expect-error) and the markers other tools read, anywhere in a comment.
      '@typescript-eslint/ban-ts-comment': ['error', { 'ts-expect-error': true, 'ts-ignore': true, 'ts-nocheck': true, 'ts-check': false }],
      'no-warning-comments': ['error', { terms: ['prettier-ignore', 'stryker disable', 'nosonar', 'sonar-ignore', 'snyk-ignore', 'deepcode ignore', 'biome-ignore', 'oxlint-disable', 'c8 ignore', 'v8 ignore', 'istanbul ignore', 'gitleaks:allow'], location: 'anywhere' }],
      '@typescript-eslint/explicit-function-return-type': ['error', { allowExpressions: true, allowTypedFunctionExpressions: true }],
      '@typescript-eslint/consistent-type-definitions': ['error', 'type'],
    },
  },
  {
    // Gate scripts (scripts/check-coverage.ts, scripts/regenerate-coverage-preload.ts)
    // are terminal tools, not production code: their whole job is printing to the
    // console that invoked them. The Logger port (rule 4) governs src/**; injecting
    // Winston into a pre-commit gate would be ceremony without observability value.
    // Project-level severity change with a comment, never an inline ignore (rule 15).
    files: ['scripts/**/*.ts'],
    rules: { 'no-console': 'off' },
  },
  {
    // Hard rule 17 for the one layer where the count is zero; the domain fallback, the
    // adapter catch and the single catch in main.ts stay with review.
    files: ['src/use-cases/**/*.ts'],
    ignores: ['**/*.test.ts'],
    rules: { 'no-restricted-syntax': ['error', ...STYLE_BANS, TRY_BAN, FS_BAN] },
  },
  {
    // Hard rule 20 over the rest of src/**; the carve-outs are paths, never inline ignores.
    files: ['src/**/*.ts'],
    ignores: ['**/*.test.ts', 'src/test-helpers/**', 'src/infra/**', 'src/use-cases/**'],
    rules: { 'no-restricted-syntax': ['error', ...STYLE_BANS, FS_BAN] },
  },
  layerZone('domain', ['use-cases', 'infra', 'presenter', 'composition', 'test-helpers']),
  layerZone('use-cases', ['infra', 'presenter', 'composition', 'test-helpers']),
  layerZone('presenter', ['use-cases', 'infra', 'composition', 'test-helpers']),
  // An adapter implements the use-case ports and never imports a use-case (the table
  // allows infra the domain and the ports only). Gitignore semantics cannot re-include a
  // file under an excluded directory, so the pattern excludes the entries of use-cases/
  // and re-includes ports/ itself; a use-case at any depth stays banned.
  layerZone('infra', ['presenter', 'composition', 'test-helpers'], undefined, [INFRA_NO_USE_CASE]),
  layerZone('composition', ['test-helpers']),
  // The fakes may reach the ports they stand for and the entry points a test harness
  // drives; an adapter is the one thing they may never wrap, or the fake stops being a fake.
  layerZone('test-helpers', ['infra']),
  // The entry point sees the composition root, the presenter and infra; it is still
  // production code, so the fakes stay out of it.
  layerZone('main.ts', ['test-helpers'], ['src/main.ts']),
  // Type-aware rules: slow (~25s on full repo), enabled only by
  // `bun run lint:strict` (which sets LINT_STRICT=1), CI's lint step.
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
  unicornPlugin.configs.recommended,
  {
    // unicorn's recommended set, less what contradicts the standard or prettier: each rule off for
    // its reason, never inline (hard rule 15). Probed 2026-09-27 against unicorn 61 and 76; a rule
    // renamed between them is named twice, since 'off' on a rule the installed version lacks is a no-op.
    rules: {
      // Formatting is prettier's; these three fight its output (eslint-config-prettier's list).
      'unicorn/empty-brace-spaces': 'off',
      'unicorn/no-nested-ternary': 'off',
      'unicorn/number-literal-case': 'off',
      // A port returns `T | null` for an absent row (references/architecture.md).
      'unicorn/no-null': 'off',
      // The standard's own names fail it: the DI factory's `deps`, the Result pair's `err()`, React's
      // `XProps`. Do not abbreviate (references/clean-code.md) stays with review.
      'unicorn/prevent-abbreviations': 'off',
      'unicorn/name-replacements': 'off',
      // Flags every guard clause followed by a return, clean-code.md's own GOOD example: the
      // standard says no `else`, guard clauses instead.
      'unicorn/prefer-ternary': 'off',
      // The standard folds with reduce (first-class collections, references/clean-code.md); a for-of
      // accumulator would trade the fold for a mutable `let`.
      'unicorn/no-array-reduce': 'off',
      // Domain predicates read as the domain says them (`moneyEquals`, references/object-design.md)
      // and the Result discriminant is `ok` (hard rule 16).
      'unicorn/consistent-boolean-name': 'off',
      // Number('') is 0 where parseFloat reads NaN, and a gate parsing a tool's text table (the
      // coverage gate) needs the parseFloat reading.
      'unicorn/prefer-number-coercion': 'off',
      // unicorn 61 wants `Number.NaN` and 76 wants `NaN`; `Number.NaN` satisfies both.
      'unicorn/prefer-global-number-constants': 'off',
      // `ok(undefined)` is how a `Result<void, E>` succeeds (hard rule 16).
      'unicorn/no-useless-undefined': ['error', { checkArguments: false }],
    },
  },
  {
    // The entry point exists for its side effect: it runs the composition root and sets
    // `process.exitCode` (Error handling, below). unicorn 79's no-top-level-side-effects flags exactly
    // that assignment (found 2026-10-24); every other module stays under the rule.
    files: ['src/main.ts'],
    rules: {
      'unicorn/no-top-level-side-effects': 'off',
    },
  },
  {
    // The test seams swap a global and restore it (installFetchMock, references/testing-infra.md);
    // production code never assigns one.
    files: ['src/test-helpers/**/*.ts', '**/*.test.ts'],
    rules: {
      'unicorn/no-global-object-property-assignment': 'off',
      'unicorn/no-unnecessary-global-this': 'off',
    },
  },
  {
    rules: {
      // false-positive-heavy rules in this codebase's idioms; disabled at project level.
      // Never inline-ignore: change severity here or refactor the code.
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
    // SonarJS rule overrides: always-on, each with its reason on its own line.
    rules: {
      'sonarjs/no-unused-vars': 'off',          // duplicates @typescript-eslint/no-unused-vars
      'sonarjs/no-empty-test-file': 'off',      // false positives on `describe` test layout
      'sonarjs/cognitive-complexity': 'off',    // one metric: the cyclomatic cap (rule 35) plus the size caps cover it
      // 2026-08-29, sonarjs 4.2.0: flags the PRIMITIVE side of every branded type
      // (`string & { __brand }` reported at the `string`, same for `number`), which is
      // hard rule 12. Object-object intersections are unaffected, so the rule is only
      // wrong about the nominal-typing idiom. Re-probed weekly by the skill repository's canary workflow (upstream, not a shipped asset).
      'sonarjs/no-useless-intersection': 'off',
      // 2026-08-29, sonarjs 4.2.0: its flow analysis ignores declared types AND explicit
      // narrowing, reporting `(v: string) => v.trim()` and even
      // `v === undefined ? 0 : v.trim()`. No correct code satisfies it, and `strict: true`
      // already owns this bug class at typecheck. Re-probed weekly by the skill repository's canary workflow (upstream, not a shipped asset).
      'sonarjs/null-dereference': 'off',
      // 2026-09-05, sonarjs 4.2.0: reports every function that returns through the
      // ok()/err() helpers of references/result-type.md, because `Result<T, never>` and
      // `Result<never, E>` are two return types to it. That is hard rule 16, so no correct
      // code satisfies it; `strict: true` and rule 6's explicit return types already own the
      // bug class. Re-probed weekly by the skill repository's canary workflow (upstream, not a shipped asset).
      'sonarjs/function-return-type': 'off',
    },
  },
  // Non-source paths must not be linted: Stryker copies the tree into .stryker-tmp/
  // during a run, reports/ is output, and the config file itself would trip no-undef
  // on `process` (it runs under Node semantics, not the **/*.ts globals block).
  // scripts/ IS linted: the gate scripts stay under the full rule set, with only
  // no-console turned off for them above.
  {
    ignores: ['eslint.config.js', '.stryker-tmp/**', 'reports/**', 'docs/**', '.claude/**', '.agents/**'],
  },
];
```

Notes on the config:

- **One config file, two modes.** Both scripts carry `--max-warnings=0`, so warnings fail either run (the zero-warning rule, hard rule 15); the modes differ only in depth. The inner-loop `bun run lint` runs the fast non-type-aware rules (~2 s cached / ~7 s cold); `bun run lint:strict` sets `LINT_STRICT=1` and the conditional block adds `parserOptions.projectService: true` plus the type-aware `@typescript-eslint` rules (~25 s on a full repo). CI runs the strict version (`assets/ci.yml`); the pre-commit hook's gate 6 lints only the staged files, fast and non-type-aware (`scripts/lint-staged.sh`, zero warnings too), and gate 7 is the typecheck. There is no separate `eslint.strict.config.js`; keeping one config eliminates drift.
- **`linterOptions.noInlineConfig`, `ban-ts-comment` and `no-warning-comments`** are hard rule 15 as lint: a directive comment is inert and reported, every `@ts-` comment is an error, and a marker another tool reads (`prettier-ignore`, `stryker disable`, `nosonar`, `sonar-ignore`, `snyk-ignore`, `deepcode ignore`, `biome-ignore`, `oxlint-disable`, the `c8`, `v8` and `istanbul` coverage ignores) is an error wherever it sits in a comment. Because directives are inert, nothing inside a file can switch any of this off.
- **`sonarjsPlugin.configs.recommended`** catches SonarLint findings at lint time so they no longer escape the IDE. See `references/workflow.md` for the common ones (S4325, S6594, S4123, S6551, S6671). Six rules are turned off, each justified in a comment beside it: `sonarjs/no-unused-vars` (duplicate), `sonarjs/no-empty-test-file` (false-positive on `describe` blocks), `sonarjs/cognitive-complexity` (one metric is enough: the cyclomatic cap of rule 35, `complexity: ['error', 10]` in the base block, plus the size caps cover it), and three that fire only in the type-aware lane and contradict the standard itself, `sonarjs/no-useless-intersection` (reports every branded type, i.e. hard rule 12), `sonarjs/null-dereference` (reports non-nullable and explicitly narrowed values, a class `strict: true` already owns), and `sonarjs/function-return-type` (reports every function that returns through the `ok()`/`err()` helpers, i.e. hard rule 16, since `Result<T, never>` and `Result<never, E>` are two types to it). The last three are dated against sonarjs 4.2.0 (2026-08-29 for the first two, 2026-09-05 for the third) and re-probed weekly by the skill repository's `.github/workflows/canary.yml` (upstream, not an asset a consumer copies), which turns them back on and reports if upstream has fixed them. Holding sonarjs at an older version is not an option: 4.1.0 does not load under ESLint 10 at all.
- **`unicornPlugin.configs.recommended`** is on since 2026-09-27 (the plugin had been registered with no rule on). What stays off contradicts the standard or prettier, each with its reason beside it: three rules that fight prettier's output, `no-null` (a port returns `T | null` for an absent row), the abbreviation rule under both its names (`prevent-abbreviations` in unicorn 61, `name-replacements` in 76; it flags the standard's own `deps`, `err()` and `XProps`), `prefer-ternary` (it flags every guard clause followed by a return, clean-code.md's GOOD example), `no-array-reduce` (the standard folds with reduce), `consistent-boolean-name` (domain predicates and the `ok` discriminant), `prefer-number-coercion` (`Number('')` is 0 where the coverage gate needs parseFloat's NaN), `prefer-global-number-constants` (unicorn 61 wants `Number.NaN` and 76 `NaN`), and `no-useless-undefined` keeps `ok(undefined)` legal. A scoped block lets the test seams (`src/test-helpers/**`, the tests) swap a global and restore it, as `installFetchMock` does; production code never assigns one. Everything else in the set holds on the references' examples and the shipped assets, which were brought in line the same day (`catch (error)`, no separator in a four-digit number, `for...of` over `forEach`, a callback wrapped rather than passed by reference).
- **`no-console` is `error`** under `src/**`. Always use the logger port (see below), never `console.*`. The one carve-out is `scripts/**`: the gate scripts shipped in `assets/` are terminal tools whose output *is* their interface; the config turns the rule off there at the project level rather than sprinkling inline ignores (rule 15).
- **`security/detect-object-injection`, `detect-unsafe-regex`, and `detect-non-literal-fs-filename`** are disabled at the project level because they only false-positive on this codebase's idioms (branded-type `Record<K, V>` lookups, bounded regexes, `chmodSync(mkdtempSync(...))` in tests). Comments in the config explain why each is off. Never inline-ignore them per-line.
- **`no-restricted-imports`** blocks `mock` from `bun:test` (the entire namespace), see hard rule 13, and carries the layer zones of hard rule 37: one block per layer under `src/`, each listing the layers it may never import, the mock ban repeated in each because ESLint replaces a matching rule's options rather than merging them.
- **`no-restricted-syntax`** is where the style rules stop being prose: `class` (hard rules 1 and 10), an inline `type` specifier (7) and a curried arrow chain with the `create[A-Z]` factory exempt (18) in every `.ts` file; a `try` under `src/use-cases/**` (17); an `fs` import under `src/**` outside `*.test.ts`, `src/test-helpers/**` and `src/infra/**` (20). Each scoped block spreads `STYLE_BANS` first, for the same replace-not-merge reason. A tightened selector needs its red fixture in the smoke test, like every gate.

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

The canonical shape for any non-trivial backend (a pipeline, batch job, or CLI with real integrations) is the Clean Architecture layout: `src/{domain,use-cases,infra,presenter,composition,test-helpers}` + `src/main.ts`. See `references/architecture.md` for the full layout, the dependency matrix, and the "adding a new external service" recipe.

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

Tests are mandatory, TDD is hard rule 11, and the whole gate pipeline (tests, coverage tiers, mutation) assumes they exist:

- Filename convention: `*.test.ts` next to the source.
- Runner: `bun test --randomize` (rule 36: the suite runs in a random order on every run and prints its seed on red; `--seed=<n>` replays it).
- See `references/testing.md` for the loop and the fakes-not-mocks discipline.

## Secrets & config hygiene

No credentials in source. Load every env var through the `envVar` branded-type factory (or its coerced siblings `envNumber` / `envEnum`) centralised in a `config/env.ts` per feature, including the logger's level (`createWinstonLogger(config.logLevel)`), so nothing reads `process.env` directly. `.env*` is git-ignored. For Firebase Admin, add `*-service-account*.json` to `.gitignore` and load the path via env var, never commit the JSON.

See `references/security.md` for the full pattern: the `envVar` / `envNumber` / `envEnum` factories (and the Zod-schema scale-up for large config), the redacted Winston logger, the never-sprinkle-`process.env` rule, and the list of what must never be committed.

## Logger (port + adapter + fake, not a module singleton)

The logger is a side-effectful dependency, so it gets port/adapter separation like every other IO dependency. Three files:

```ts
// src/use-cases/ports/logger.ts: type only, zero dependencies
export type LogMeta = Readonly<Record<string, unknown>>;

export type Logger = {
  readonly info: (event: string, meta?: LogMeta) => void;
  readonly warn: (event: string, meta?: LogMeta) => void;
  readonly error: (event: string, meta?: LogMeta) => void;
};
```

```ts
// src/infra/logger.ts: Winston-backed adapter, real for production
import { createLogger, format, transports } from 'winston';
import type { Logger } from '../use-cases/ports/logger.ts';

const REDACTED_KEYS = new Set(['password', 'token', 'authorization', 'apikey', 'secret', 'email', 'phone']); // secrets plus natural identifiers (rule 27); extend with the domain's own

const redactFormat = format((info) => {
  for (const key of Object.keys(info)) {
    if (REDACTED_KEYS.has(key.toLowerCase())) info[key] = '[REDACTED]';
  }
  return info;
});

export const createWinstonLogger = (level: string): Logger => {
  const winston = createLogger({
    level,
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
// src/test-helpers/logger-fake.ts: in-memory fake for tests, logs become assertable
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

Every use-case declares `readonly logger: Logger` in its `Deps` and calls `deps.logger.info(...)`. Composition wires `createWinstonLogger(config.logLevel)` in `src/composition/build-deps.ts` (the level comes from the typed-env config, never `process.env` directly). Tests inject `createLoggerFake()` and assert on the `calls` array: logs become assertable without a mocking library.

Why this and not a module-level singleton: a singleton makes the logger impossible to swap in tests without monkey-patching, and impossible to redact/reformat per-environment without mutating global state. A port is one extra type declaration and pays off the first time you want to assert that a warning fired, or run a test suite in silent mode.

Invariant: `grep -rn "from '.*infra" src/domain src/use-cases` must return nothing. The domain and use-cases know only about the `Logger` **type**; the Winston import lives in `infra/` only.

Rule 4 in full: no `console.*` anywhere in application code, enforced by ESLint's `no-console` in both variant configs. The rule is scoped to application code: `scripts/**` gate scripts are terminal tools whose output is their interface, so the config turns it off there at the project level (never inline). The one sanctioned module-level logger is the Next.js client/static exception, `src/lib/utils/logger.ts`, because the React client boundary and the static export leave no composition root to inject through (`references/nextjs-monorepo.md`, Winston logger); a Next.js server app has a composition root and injects the port like this variant. Everywhere else a module-level logger stays banned.

## Error handling

No `try/catch` anywhere outside `src/infra/**`, pure-domain fallbacks for native-synchronous throwers (`JSON.parse`, `URL` constructor), and exactly one top-level handler in `src/main.ts`. Every IO port returns `Promise<Result<T, PortError>>`. Use-cases pattern-match on `.ok` and aggregate port errors into `StepError`. See `references/result-type.md` for the full treatment, the discriminated-union error design, the fan-out batch semantics, and the `retryOnErr` + `captureRejection` helpers.

The shared `formatError(err: unknown): string` helper lives in `src/domain/utilities/format-error.ts`. Use it in every `catch (error)` block in `src/infra/**`, never `String(error)`, which returns `"[object Object]"` for non-Error throws (SonarJS S6551).

A failed run ends in `src/main.ts` only, after the top-level catch, by setting `process.exitCode = 1`: never `process.exit(1)` (unicorn's `no-process-exit`, on in the recommended set, rejects it outside a hashbang CLI file, and a hard exit can drop log lines still being written), and never an exit code set inside a use-case, adapter, or domain module.

## File IO (rule 20)

All **file** IO in `src/**` production code goes through the Bun file API: read with `Bun.file(path).text()` / `.json()` / `.arrayBuffer()` / `.bytes()` / `.exists()`; write with `Bun.write(path, contents)`, which creates parent directories itself, no `mkdir -p` ceremony; delete with `Bun.file(path).delete()` (Bun 1.1 and later). `node:fs` is forbidden for file operations under `src/**`, and the lint says so: the `FS_BAN` selector in `eslint.config.js` rejects an `fs` import under `src/**` outside `*.test.ts`, `src/test-helpers/**` and `src/infra/**` (the directory helper of point 2 below lives there).

**Directories are the exception.** Bun has no native primitive for `mkdir`, `rmdir`, or directory-existence as such (`Bun.file(dir).exists()` returns `false` for a directory: that is "not a file", not "directory missing"). Two acceptable answers, in order of preference:

1. **Let the library handle it.** Most SDKs that need a directory create it themselves (Playwright auto-creates `userDataDir`, better-sqlite3 creates the parent on file open). Pass the path; let the library do `mkdir`.
2. **`node:fs` at the boundary, with a comment.** When no library is taking the call (a CLI scaffolds an output dir; a fixture cleanup removes a tree), import `mkdirSync` / `rmSync` from `node:fs` directly, isolated to a single helper in `src/infra/**`, with a one-line comment naming the gap. Permitted because Bun has no replacement; never a workaround for laziness.

`node:fs` IS unconditionally allowed in `*.test.ts` and `src/test-helpers/**`, for real temp-dir setup (`mkdtempSync`, `writeFileSync`, `rmSync`) and for forcing error branches in FS adapters (`chmodSync` on a real file or directory): `Bun.file` has no `mkdtemp` equivalent and cannot force a directory-write throw. `node:path` (`join`, `dirname`, `resolve`, `basename`) is allowed anywhere; it is path manipulation, not IO.

Why: keeping file IO on `Bun.file` is faster, has zero import ceremony, fits the try/catch-in-`infra/**` quarantine cleanly, and lets the project turn `security/detect-non-literal-fs-filename` off at the lint level without losing real coverage (the rule does not watch `Bun.file`). See `references/result-type.md`, `references/testing-infra.md` (filesystem patterns), and `references/workflow.md` (lint-rule rationale).

## Bootstrap checklist (fresh Bun repo)

1. `mkdir <new-repo> && cd <new-repo> && bun init -y`.
2. Replace `package.json` with the skeleton above (devDependencies include `eslint-plugin-sonarjs`; scripts include `test`, `lint`, `lint:strict`, `lint:staged`, `typecheck`, `coverage`, `mutate`, `mutate:changed`, `mutate:staged`, `start`. Keep `test` as the skeleton has it, `bun test --randomize` (rule 36) and nothing more; never add `--pass-with-no-tests`, which turns a suite that has vanished into a green run, the gate-that-cannot-fail canon 15.10 rejects). **No `"latest"` or `"*"` anywhere**; the skeleton's `^X.Y.Z` ranges are samples; bump them in step 7 below.
3. Create `tsconfig.json` with the block above (includes `"types": ["bun"]`).
4. Create `eslint.config.js` with the flat config above (includes `sonarjs.configs.recommended` and type-aware `@typescript-eslint` rules behind `LINT_STRICT=1`).
5. Create `.vscode/settings.json` and `.vscode/extensions.json`.
6. Drop in a Node `.gitignore`, plus `*-service-account*.json` if Firebase is in play, and `cp <skill-path>/assets/gitattributes .gitattributes` (text checks out LF on every machine: Git for Windows turns a checkout CRLF without it, and the formatter's `endOfLine: 'lf'` then fails every file).
7. `bun install` to resolve the skeleton's ranges; then `bun update` to bump every dep to its current latest matching version. Commit `bun.lock` and the updated `package.json` together. From this point, every new dep is added via `bun add <pkg>` (runtime) or `bun add -d <pkg>` (dev), never hand-edit `package.json`.
8. Scaffold the Clean Architecture layout (see `references/architecture.md`): `mkdir -p src/{domain,use-cases/ports,infra,presenter,composition,test-helpers}`.
9. Create `src/domain/result.ts` with the `Result<T, E>` type and helpers from `references/result-type.md`.
10. Create `src/use-cases/ports/logger.ts` and `src/infra/logger.ts` with the port + Winston adapter above; create `src/test-helpers/logger-fake.ts`.
11. Copy the canonical helpers from the skill's `assets/`:
    - `cp <skill-path>/assets/format-error.ts src/domain/utilities/format-error.ts`
    - `cp <skill-path>/assets/format-error.test.ts src/domain/utilities/format-error.test.ts` (format-error is in the mutation scope; its shipped test keeps it above the 90% gate)
    - `cp <skill-path>/assets/capture-rejection.ts src/test-helpers/capture-rejection.ts`
    - `cp <skill-path>/assets/fetch-mock.ts src/test-helpers/fetch-mock.ts`
12. Set up the per-tier coverage gate:
    - `cp <skill-path>/assets/check-coverage.ts scripts/check-coverage.ts`
    - `cp <skill-path>/assets/regenerate-coverage-preload.ts scripts/regenerate-coverage-preload.ts`
    - `bun run scripts/regenerate-coverage-preload.ts`: generates `scripts/coverage-preload.ts` from the current tree; re-run it (or wire `--check` as the pre-commit pre-flight) whenever an infra/composition/presenter file is added. Never hand-edit the generated file.
    - In `bunfig.toml`: `[test]` section with `coverage = true`, `coverageSkipTestFiles = true`, `coverageReporter = ["text"]`. **Do not** add `coverageThreshold` (the per-tier script owns enforcement) and **do not** add `preload` (the coverage preload is loaded only by `scripts/check-coverage.ts` via `--preload` so plain `bun test` runs stay fast).
13. Set up mutation testing:
    - `cp <skill-path>/assets/stryker.conf.json stryker.conf.json`
    - `cp <skill-path>/assets/mutate-staged.sh scripts/mutate-staged.sh`
    - `cp <skill-path>/assets/mutate-changed.sh scripts/mutate-changed.sh`
    - `cp <skill-path>/assets/check-mutation-score.ts scripts/check-mutation-score.ts` (holds every file to the break threshold; Stryker's own break judges the whole run)
    - `chmod +x scripts/*.sh`
    - Add to `.gitignore`: `.stryker-tmp/` and `reports/` (Stryker scratch + output dirs).
14. Install the git hooks (fast-gate pre-commit + commit-msg) and the CI workflow:
    - `cp <skill-path>/assets/check-commit-size.sh scripts/check-commit-size.sh`
    - `cp <skill-path>/assets/check-package-json.sh scripts/check-package-json.sh`
    - `chmod +x scripts/check-commit-size.sh scripts/check-package-json.sh`
    - `mkdir -p .githooks`
    - `cp <skill-path>/assets/lint-staged.sh scripts/lint-staged.sh` (hook gate 6 runs it via the `lint:staged` script)
    - `cp <skill-path>/assets/check-commit-messages.sh scripts/check-commit-messages.sh` (CI re-runs the message check over the pushed range, so `--no-verify` cannot slip one past)
    - `cp <skill-path>/assets/check-commit-range.sh scripts/check-commit-range.sh` (the same for commit SIZE: the hook sees one staged diff, CI walks every commit in the range)
    - `cp <skill-path>/assets/land-on-green.sh scripts/land-on-green.sh` (the CI workflow's `land` job: a green pull request lands itself by rebase, `references/workflow.md` § Landing on green)
    - `cp <skill-path>/assets/check-branches.sh scripts/check-branches.sh`, `cp <skill-path>/assets/check-host.py scripts/check-host.py`, `cp <skill-path>/assets/host-ruleset.json scripts/host-ruleset.json` and `cp <skill-path>/assets/branches.yml .github/workflows/branches.yml` (the daily watchdog, hard rule 38: a branch whose work already landed or one more than a day off `main`, and a host whose default branch is not `main`, whose `main` is unprotected, or whose merges are not rebase-only with head-branch deletion); after the first push the owner runs `python3 scripts/check-host.py --apply` once, with an admin token (`references/workflow.md`, Branch lifecycle)
    - `cp <skill-path>/assets/check-identity.sh scripts/check-identity.sh` (hook gate 4, rule 26: no person, employer, or client named in file contents; CI runs it over the whole tree with `--all`; a private repository also commits `.github/visibility` reading `private`, and the gate then warns instead of blocking)
    - `cp <skill-path>/assets/check-reply.py scripts/check-reply.py` and `mkdir -p .claude && cp <skill-path>/assets/claude-settings.json .claude/settings.json` (the reply gate: a Claude Code Stop hook, not a git hook, that has the agent restate a reply breaking the Interaction section's mechanical rules; merge its `Stop` entry instead when `.claude/settings.json` exists; `references/workflow.md`, Reply gate)
    - `cp <skill-path>/assets/check-disciplines.sh <skill-path>/assets/check-pii-channels.sh <skill-path>/assets/check-io-deadlines.sh <skill-path>/assets/check-data-lifecycle.sh scripts/` (hook gate 5, rules 27, 29, 30: the three discipline tripwires that are core gates, run by the wrapper; CI runs it with `--all`). Where tenants or owners exist, copy `check-isolation-tests.sh` too and call it beside the wrapper in the hook and CI (rule 28, `references/workflow.md`, Discipline tripwires)
    - `chmod +x scripts/lint-staged.sh scripts/check-commit-messages.sh`
    - `mkdir -p .github/workflows && cp <skill-path>/assets/ci.yml .github/workflows/ci.yml` (the authoritative gate set: strict lint, tests, coverage, mutation on the changed files, secret scan on a frozen lockfile)
    - `cp <skill-path>/assets/check-docs.sh scripts/check-docs.sh` (ci.yml runs the README's `## Verify` block before mutation, entry points only, so a documented command that stopped working fails the push; `references/governance.md`, README stays runnable)
    - `cp <skill-path>/assets/mutation.yml .github/workflows/mutation.yml` (the daily full mutation sweep, `workflow_dispatch` on demand; ci.yml mutates the changed files only, so this is the only place the whole tree is measured)
    - `cp <skill-path>/assets/audit.yml .github/workflows/audit.yml` (the CVE watchdog: daily schedule plus dependency-scoped PR runs; deliberately not a gate on unrelated commits)
    - `cp <skill-path>/assets/check-skill-pin.sh scripts/check-skill-pin.sh` (always: the audit workflow runs it on every run. In a repo that vendors or pins the skill it compares the whole vendored tree against the repository named in its `SKILL_PIN_UPSTREAM` env and fails when any file falls behind; with nothing vendored it passes)
    - `cp <skill-path>/assets/pre-commit .githooks/pre-commit`
    - `cp <skill-path>/assets/commit-msg .githooks/commit-msg` (Conventional Commits validator, hard rule 23: dependency-free, no `package.json` change)
    - `chmod +x .githooks/pre-commit .githooks/commit-msg`
    - `git config core.hooksPath .githooks` (picks up both hooks)
    - Optional: `brew install gitleaks` (macOS) or grab a binary from `github.com/gitleaks/gitleaks/releases`. The hook degrades gracefully if missing.
    - See `references/workflow.md` for the gate breakdown (fast hook plus CI set), the commit-message format, and the no-bypass rule.
15. Verify: `bun run lint`, `bun run typecheck`, `bun run coverage`, and `bun run mutate` all clean on a minimal `src/main.ts`. Run `bash scripts/check-package-json.sh` once to confirm no `"latest"` slipped in, and confirm the `commit-msg` hook rejects a junk message by running it on a message file, the way git does (it reads `$1`, not stdin): `printf 'nope\n' > .msg && bash .githooks/commit-msg .msg` must exit non-zero; delete `.msg` after.
16. Commit with Conventional Commits (`type(scope): subject`), once the user confirms (rule 25); the `commit-msg` hook enforces the format. From here, follow the Clean Architecture rules for every new feature.

## Containerization (optional)

The atelier takes no position on deployment: `atelier-greenfield` scopes Docker out of repo-birth, and the canonical archetypes (CLIs, batch jobs, Firebase Admin jobs) ship as a `bun run`, not an image. This section exists only so that *if* you containerize, the image conforms instead of drifting. It is documentation, not a gate.

A minimal, production-ready multi-stage build:

```dockerfile
# syntax=docker/dockerfile:1
FROM oven/bun:1 AS base
WORKDIR /usr/src/app

# Production deps only, in a layer cached on the lockfile.
FROM base AS install
COPY package.json bun.lock ./
RUN bun install --frozen-lockfile --production

# Final image: prod deps + source, run as the non-root `bun` user.
FROM base AS release
COPY --from=install /usr/src/app/node_modules node_modules
COPY package.json ./
COPY src/ src/
USER bun
ENTRYPOINT ["bun", "run", "src/main.ts"]
```

Four things keep it conforming, and they are exactly where a copied-from-a-blog Dockerfile drifts:

- **Entry is `src/main.ts`**, never `src/index.ts`: the atelier's named entry (rule 5, `"module": "src/main.ts"`).
- **Copy `bun.lock`, not `bun.lockb`**: Bun's lockfile is text now; the binary `bun.lockb` is legacy.
- **No `EXPOSE`** for the CLI/batch archetype: it runs and `process.exit`s; there is no port to bind. Add `EXPOSE <port>` only for an actual server whose `src/main.ts` calls `Bun.serve`.
- **No `bun run lint` or tests inside the build.** Quality is already owned by the seven fast pre-commit gates and CI; linting in the image duplicates the gate and couples building with checking. If you want a build-time backstop anyway, run `bun run lint:strict` (the full type-aware gate) rather than bare `bun run lint` (which runs only the fast non-type-aware rules; both already fail on warnings).

Add a `.dockerignore` so the build context stays small and the image never ships local cruft:

```
node_modules
.git
coverage
.stryker-tmp
reports
```

Build and run a CLI image (argv in, process exits, no port mapping):

```bash
docker build -t my-app .
docker run --rm my-app <args>
```
