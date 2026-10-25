# Next.js Monorepo Variant

Applies to repos with the Bun-workspace + Next.js 16 layout. Identifiable by `packages/*`, `next.config.ts`, and `app/(en)/` route groups.

This reference describes two shapes that share the same toolchain: the **static content site** (the default, `output: 'export'`, build-time data, i18n route groups) which everything below assumes, and the **server app** (route handlers, runtime state) covered in its own sub-variant section. Pick the static shape unless the app must answer requests or hold state at runtime; the two are mutually exclusive (static export cannot run request-time route handlers).

## Workspace layout

```
<repo>/
├── package.json              # root workspace + commit hooks only
├── commitlint.config.mjs
├── bun.lock
├── .gitignore
├── .vscode/
└── packages/
    └── 01-<name>/            # one Next.js app per package
        ├── package.json
        ├── tsconfig.json
        ├── eslint.config.mjs
        ├── postcss.config.mjs
        ├── next.config.ts
        ├── app/
        │   ├── (en)/ (es)/ (fr)/ (de)/ (pt)/ (zh)/ (ja)/
        │   ├── layout.tsx
        │   └── globals.css   # Tailwind v4 entrypoint
        ├── data/
        │   ├── guides/       # MDX
        │   └── translations/ # JSON
        ├── public/
        └── src/
            ├── components/
            │   ├── atoms/
            │   ├── molecules/
            │   └── organisms/
            ├── config/
            ├── lib/
            │   ├── guides/
            │   ├── hooks/
            │   ├── i18n/
            │   ├── layout/
            │   ├── seo/
            │   └── utils/
            ├── page/
            └── types/
```

- Root holds only workspace plumbing + commit hooks.
- All app dependencies live in the package's `package.json`.
- Run commands: `bun install`, `bun run --filter <package-name> <script>`.

## Root `package.json`

```json
{
  "name": "workspace-root",
  "type": "module",
  "private": true,
  "workspaces": ["packages/*"],
  "devDependencies": {
    "@commitlint/cli": "^20.2.0",
    "@commitlint/config-conventional": "^20.2.0",
    "simple-git-hooks": "^2.13.1"
  },
  "scripts": {
    "prepare": "simple-git-hooks"
  },
  "simple-git-hooks": {
    "pre-commit": "bash scripts/check-package-json.sh && bash scripts/check-identity.sh && bash scripts/check-disciplines.sh && bun run --filter <package-name> test && bun run --filter <package-name> lint",
    "commit-msg": "bunx --yes commitlint --edit $1"
  }
}
```

Activate hooks after install: `bun run prepare`.

**This variant's hook mechanism is `simple-git-hooks`** (gate 2 first, `scripts/check-package-json.sh` copied from the skill's `assets/`: no `"latest"`, no foreign lockfile, no `scripts` entry calling `node`, `npm`, `npx`, `pnpm`, `yarn` or `vite`, rules 5 and 19; then the identity gate and the discipline wrapper, rules 26, 27, 29, 30; then test + lint per package, commitlint on the message). The `.githooks/pre-commit` fast-gate hook from `references/workflow.md` belongs to the Bun-script variant, never install both: `core.hooksPath` and `simple-git-hooks` overwrite each other. The commit-size and gitleaks gates are portable here if wanted; the coverage and mutation gates are not (see SKILL.md, "What applies where").

## Package `package.json`

```json
{
  "name": "<package-name>",
  "version": "0.1.0",
  "private": true,
  "type": "module",
  "scripts": {
    "dev": "bun next dev",
    "build": "rimraf out && bun next build",
    "start": "bunx serve ./out",
    "test": "bun test --randomize",
    "typecheck": "tsc --noEmit",
    "lint": "eslint --max-warnings=0"
  },
  "dependencies": {
    "next": "16.4.2",
    "next-mdx-remote": "^6.0.0",
    "react": "19.2.3",
    "react-dom": "19.2.3",
    "winston": "^3.19.0"
  },
  "devDependencies": {
    "@eslint/js": "^9.39.2",
    "@tailwindcss/postcss": "^4.1.18",
    "@types/bun": "^1.2.0",
    "@types/mdx": "^2.0.13",
    "@types/node": "^20.19.27",
    "@types/react": "^19.2.7",
    "@types/react-dom": "^19.2.3",
    "@typescript/native": "npm:typescript@^7.0.2",
    "baseline-browser-mapping": "^2.9.11",
    "eslint": "^9.39.2",
    "eslint-config-next": "16.4.2",
    "eslint-plugin-jsx-a11y": "^6.10.2",
    "eslint-plugin-prettier": "^5.5.4",
    "eslint-plugin-react": "^7.37.5",
    "eslint-plugin-react-hooks": "^7.0.1",
    "eslint-plugin-security": "^3.0.1",
    "eslint-plugin-tailwindcss": "^4.0.2",
    "eslint-plugin-unicorn": "^61.0.2",
    "globals": "^17.0.0",
    "rimraf": "^6.1.2",
    "tailwindcss": "^4.1.18",
    "typescript": "npm:@typescript/typescript6@^6.0.2",
    "typescript-eslint": "^8.51.0"
  },
  "trustedDependencies": ["sharp", "unrs-resolver"],
  "browserslist": ["> 0.5%", "last 2 versions", "not dead", "not IE 11", "not op_mini all"]
}
```

No `format` or `lint:fix` script: save-in-editor triggers ESLint autofix. The `test` and `typecheck` scripts are mandatory: `test` is the TDD gate (see Testing below) and `typecheck` (`tsc --noEmit`) is the standalone type gate the workflow loop calls, `next build` typechecks too, but you want the fast check without a full build. Notes on the skeleton: `@types/bun` ships the `bun:test` module types, without it, `import { describe, it, expect } from 'bun:test'` in the mandated test files does not type-resolve (the same caveat the Bun-script variant documents); set the tsconfig `types` key to `["bun"]`: from TypeScript 6.0, `tsc` does not load `@types/bun` without it (TS2307 on `bun:test`, in `tsc` and in `next build`). `@types/react` and `@types/react-dom` still resolve through the `react` imports; add `"mdx"` when the app imports `.mdx` files. Restart the VS Code TS server after adding `@types/bun`. `next build` type-checks with `tsc6` (TypeScript 6.0), because Next.js runs the binary of the `typescript` package; the `typecheck` script runs TypeScript 7, and CI runs both. The two aliases need Bun 1.4 or newer (the Bun-script variant's TypeScript note says why). `@types/winston` must NOT be added (winston 3 ships its own types; the v2 stub conflicts), and `trustedDependencies` is Bun's lifecycle-script allowlist, there is no `ignoreScripts` package.json field.

## `tsconfig.json`

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "lib": ["dom", "dom.iterable", "esnext"],
    "allowJs": true,
    "skipLibCheck": true,
    "strict": true,
    "noFallthroughCasesInSwitch": true,
    "noUncheckedIndexedAccess": true,
    "noImplicitOverride": true,
    "noUnusedLocals": false,
    "noUnusedParameters": false,
    "noPropertyAccessFromIndexSignature": false,
    "noEmit": true,
    "types": ["bun"],
    "esModuleInterop": true,
    "module": "esnext",
    "moduleResolution": "bundler",
    "allowImportingTsExtensions": true,
    "resolveJsonModule": true,
    "isolatedModules": true,
    "jsx": "preserve",
    "incremental": true,
    "plugins": [{ "name": "next" }],
    "paths": { "@/*": ["./*"] }
  },
  "include": [
    "next-env.d.ts",
    "**/*.ts",
    "**/*.tsx",
    ".next/types/**/*.ts",
    ".next/dev/types/**/*.ts",
    "**/*.mts"
  ],
  "exclude": ["node_modules"]
}
```

Full strictness: `strict`, `noUncheckedIndexedAccess`, `noImplicitOverride`. `moduleResolution: "bundler"` + `isolatedModules` required by Next.js / Turbopack.

`"allowImportingTsExtensions": true` is standard here, not a vendoring exception: this variant and the Bun-script variant both import with explicit `.ts`/`.tsx` extensions, so the same import style works across every package in the monorepo (and the server-app example below uses it). Turbopack resolves the extensionful specifier at build time; without the flag, `tsc --noEmit` errors `TS5097`.

`"jsx"` is Next-managed: Next runs its own JSX transform and rewrites this key on the first `dev`/`build` regardless of what you set (observed on 16.1.1: `preserve` becomes `react-jsx`, "next.js uses the React automatic runtime"). The build succeeds either way, so treat the value as owned by Next and do not fight the managed diff.

## `eslint.config.mjs`

Flat config, ESM, filename ends in `.mjs` (not `.js`).

```js
import pluginJs from '@eslint/js';
import nextVitals from 'eslint-config-next/core-web-vitals';
import nextTs from 'eslint-config-next/typescript';
import prettier from 'eslint-plugin-prettier';
import react from 'eslint-plugin-react';
import securityPlugin from 'eslint-plugin-security';
import tailwind from 'eslint-plugin-tailwindcss';
import unicornPlugin from 'eslint-plugin-unicorn';
import { defineConfig, globalIgnores } from 'eslint/config';
import globals from 'globals';
import tsPlugin from 'typescript-eslint';

// The style rules as lint (hard rules 1, 7, 10, 18), the same list as the Bun variant's
// eslint.config.js. ESLint REPLACES a rule's options when a second block matches the same
// file, it never merges them, so the design-system and app-shell blocks below spread this
// list before their own selectors.
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
// Hard rule 17 for the server archetype: a use-case pattern-matches the Result a port returns.
const TRY_BAN = {
  selector: 'TryStatement',
  message: 'try/catch is quarantined to the inbound adapter (the route handler), src/infra/** and the pure-domain fallback; a use-case pattern-matches the Result (hard rule 17, references/result-type.md).',
};
// Hard rule 13, one object for every block that sets no-restricted-imports: ESLint REPLACES a
// rule's options when a second block matches the same file, so each scoped block and each
// layer zone below carries the mock ban itself.
const MOCK_BAN = {
  name: 'bun:test',
  importNames: ['mock', 'spyOn', 'jest', 'vi'],
  message: '`mock`, `spyOn`, `jest` and `vi` from bun:test are forbidden, mocks leak across test files. Use a hand-written fake (hard rule 13).',
};
// Hard rule 21: the design system imports react and its own lower layers only. The atoms and
// molecules zones spread this list again for the same replace-not-merge reason.
const DESIGN_SYSTEM_BANS = [
  { group: ['next', 'next/*'], message: 'Design-system components import react only: inject links/images as ComponentType props (hard rule 21).' },
  { group: ['**/lib/**', '**/config/**', '**/page/**'], message: 'Design-system components must not import application code (hard rule 21).' },
];
// Hard rule 37: one no-restricted-imports zone per layer, listing the layers it may never
// import, tests excepted. Two shapes in this variant: the design system's own layers point
// upward (references/atomic-design.md, Imports point strictly upward) and the server
// archetype's src/{domain,use-cases,infra,presenter,composition} mirror the Bun config's
// zones (references/architecture.md, the dependency table).
const INWARD = 'dependencies point inward (hard rule 37, references/architecture.md, the dependency table)';
// Hard rule 20 in the server sub-variant: file IO stays at the edges. The domain and the
// use-cases never import fs; an infra adapter does, behind a port. Elsewhere in a Next package
// Node is the runtime and its fs stays legal; the static layout has no server layers at all.
const FS_AT_THE_EDGES = {
  group: ['fs', 'node:fs', 'fs/promises', 'node:fs/promises'],
  message: 'File IO stays at the edges: node:fs only in src/infra/**, tests and src/test-helpers/**; the domain and the use-cases read through a port (hard rule 20, references/bun-typescript.md, File IO).',
};
const UPWARD = 'imports point upward inside the design system (hard rule 37, references/atomic-design.md, Imports point strictly upward)';
// A `<factory>Unsafe` helper casts a brand without validating (references/testing.md,
// Branded types and `expect(...).toBe(raw)`); only tests and the fakes may import one. It
// rides inside the zones: a separate no-restricted-imports block would replace them.
// An adapter implements the use-case ports and never imports a use-case; the pattern
// excludes the entries of use-cases/ and re-includes ports/ (gitignore semantics).
const INFRA_NO_USE_CASE = {
  group: ['**/use-cases/*', '!**/use-cases/ports'],
  message: 'src/infra implements the use-case ports and never imports a use-case: dependencies point inward (hard rule 37, references/architecture.md, the dependency table).',
};
const UNSAFE_BAN = {
  group: ['**'],
  importNamePattern: 'Unsafe$',
  message: '*Unsafe helpers skip validation and are test-only: production code builds the value through its factory (references/testing.md, Branded types).',
};
const layerZone = (layer, forbidden, extraPatterns = [], why = INWARD) => ({
  files: [`src/${layer}/**/*.ts`, `src/${layer}/**/*.tsx`],
  ignores: ['**/*.test.ts', '**/*.test.tsx'],
  rules: {
    'no-restricted-imports': [
      'error',
      {
        paths: [MOCK_BAN],
        patterns: [
          ...extraPatterns,
          {
            group: forbidden.flatMap((name) => [`**/${name}`, `**/${name}/**`]),
            message: `src/${layer} must not import ${forbidden.join(', ')}: ${why}.`,
          },
          ...(layer === 'test-helpers' ? [] : [UNSAFE_BAN]),
        ],
      },
    ],
  },
});

const eslintConfig = defineConfig([
  securityPlugin.configs.recommended,
  {
    // Hard rule 15: no inline ignore, ever. Directive comments (eslint-disable*, eslint-enable,
    // globals, exported) are inert AND each one is reported, so `--max-warnings=0` fails on the
    // comment and the violation it hid surfaces beside it. A finding is a refactor or a
    // project-level severity change with a reason, never a suppression.
    linterOptions: { noInlineConfig: true },
  },
  {
    files: ['**/*.ts', '**/*.tsx'],
    rules: {
      // false-positive-heavy security rules on this codebase's idioms; disabled
      // at project level. Never inline-ignore: change severity here or refactor.
      // detect-object-injection fires on the prescribed Winston redaction loop
      // (`info[key] = '[REDACTED]'`); detect-unsafe-regex on bounded regexes.
      'security/detect-object-injection': 'off',
      'security/detect-unsafe-regex': 'off',
      'security/detect-non-literal-fs-filename': 'off',
    },
  },
  {
    languageOptions: {
      globals: { ...globals.node, ...globals.browser },
    },
  },
  {
    rules: {
      'func-style': ['error', 'expression'],
      // Rule 35: cyclomatic complexity at most 10 per function (the size caps
      // miss a one-line chain of `&&`/`??`/ternaries and a wide `switch`).
      complexity: ['error', 10],
      'no-console': 'error',
      // Hard rules 1, 7, 10, 18 (STYLE_BANS above); the rule 21 and rule 22 blocks spread it again.
      'no-restricted-syntax': ['error', ...STYLE_BANS],
      'prefer-template': 'error',
      quotes: ['error', 'single', { avoidEscape: true }],
      // Mock ban (hard rule 13). Lives in this unscoped block (which precedes the
      // design-system block) AND is re-declared inside the design-system block and every
      // layer zone: ESLint flat config REPLACES (never merges) two `no-restricted-imports`
      // objects that match the same file, so each scope must carry its full set. The
      // same holds for `no-restricted-syntax`, hence `...STYLE_BANS` in every scoped block.
      'no-restricted-imports': ['error', { paths: [MOCK_BAN] }],
    },
  },
  {
    rules: {
      '@typescript-eslint/explicit-function-return-type': ['error', { allowExpressions: true, allowTypedFunctionExpressions: true }],
      '@typescript-eslint/consistent-type-definitions': ['error', 'type'],
      '@typescript-eslint/consistent-type-imports': [
        'error',
        { prefer: 'type-imports', fixStyle: 'separate-type-imports' },
      ],
      // Hard rule 15, the other tools' escape hatches: every @ts- form (the recommended preset
      // allows a described @ts-expect-error) and the markers other tools read, anywhere in a comment.
      '@typescript-eslint/ban-ts-comment': ['error', { 'ts-expect-error': true, 'ts-ignore': true, 'ts-nocheck': true, 'ts-check': false }],
      'no-warning-comments': ['error', { terms: ['prettier-ignore', 'stryker disable', 'nosonar', 'sonar-ignore', 'snyk-ignore', 'deepcode ignore', 'biome-ignore', 'oxlint-disable', 'c8 ignore', 'v8 ignore', 'istanbul ignore', 'gitleaks:allow'], location: 'anywhere' }],
    },
  },
  {
    // Hard rule 17 for the one layer where the count is zero (server archetype); inert in
    // the static layout, which has no src/use-cases.
    files: ['src/use-cases/**/*.ts'],
    ignores: ['**/*.test.ts'],
    rules: { 'no-restricted-syntax': ['error', ...STYLE_BANS, TRY_BAN] },
  },
  unicornPlugin.configs.recommended,
  {
    // The Bun variant's unicorn block: the recommended set, less what contradicts the standard or
    // prettier, each rule off for its reason, never inline (hard rule 15). Probed 2026-09-27 against
    // unicorn 61 and 76; a rule renamed between them is named twice, since 'off' on a rule the
    // installed version lacks is a no-op.
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
    // The test seams swap a global and restore it (installFetchMock, references/testing-infra.md);
    // production code never assigns one.
    files: ['src/test-helpers/**/*.ts', 'src/test-helpers/**/*.tsx', '**/*.test.ts', '**/*.test.tsx'],
    rules: {
      'unicorn/no-global-object-property-assignment': 'off',
      'unicorn/no-unnecessary-global-this': 'off',
    },
  },
  {
    files: ['**/*.tsx'],
    plugins: { react },
    languageOptions: { parserOptions: { ecmaFeatures: { jsx: true } } },
    settings: { react: { version: 'detect' } },
    rules: { 'react/react-in-jsx-scope': 'off' },
  },
  {
    // Hard rules 21–22: the design system imports react only, holds no state, owns all styling
    files: ['src/components/**/*.ts', 'src/components/**/*.tsx'],
    rules: {
      'no-restricted-imports': [
        'error',
        {
          // Carries BOTH the mock ban (hard rule 13) and the design-system bans:
          // flat config replaces, not merges, so re-declaring the mock ban here is
          // mandatory: the general unscoped block's copy is overwritten for these files.
          paths: [MOCK_BAN],
          patterns: [...DESIGN_SYSTEM_BANS],
        },
      ],
      'no-restricted-syntax': [
        'error',
        ...STYLE_BANS,
        // Bare `useX()`, React 19's `use()`, and the member forms (`React.useState()`): the
        // bare selector alone missed the last two until 2026-09-27.
        { selector: 'CallExpression[callee.name=/^use([A-Z]|$)/]', message: 'No hooks inside the design system: hoist state to the page shell via src/lib/hooks (hard rule 21).' },
        { selector: 'CallExpression[callee.property.name=/^use([A-Z]|$)/]', message: 'No hooks inside the design system: hoist state to the page shell via src/lib/hooks (hard rule 21).' },
        { selector: 'Program > ExpressionStatement[directive="use client"]', message: "The 'use client' boundary belongs to page shells, not design-system components (hard rule 21)." },
      ],
      // Accessible by default (canon 17.6): the design system is where a11y is won or
      // lost, so the doctrine's structural rules are error-level here. The jsx-a11y plugin
      // is registered by next/core-web-vitals (spread below), which enables its recommended
      // subset app-wide but NOT the interaction rules that catch the flagship clickable-div;
      // these reference that same plugin and make them explicit and unmissable. Contrast is
      // not lintable (needs layout): it lives in the design tokens (rule 22) and review.
      'jsx-a11y/no-static-element-interactions': 'error',
      'jsx-a11y/click-events-have-key-events': 'error',
      'jsx-a11y/interactive-supports-focus': 'error',
      'jsx-a11y/control-has-associated-label': 'error',
      'jsx-a11y/alt-text': 'error',
      'jsx-a11y/anchor-is-valid': 'error',
    },
  },
  {
    // Hard rule 22 (the mirror of rule 21): styling is sealed inside the design system.
    // Routes, page shells, lib, and config never carry a class string, visual variation
    // is a typed variant prop on a design-system component, never free-form className/
    // style outside src/components/**. Rule 21 bans the imports; this bans the styling leak.
    files: ['app/**/*.tsx', 'src/page/**/*.tsx', 'src/lib/**/*.tsx', 'src/config/**/*.tsx'],
    rules: {
      'no-restricted-syntax': [
        'error',
        ...STYLE_BANS,
        { selector: "JSXAttribute[name.name='className']", message: 'No className outside the design system. Tailwind is sealed under src/components/** (hard rule 22). Move the styling into a design-system component with a typed variant.' },
        { selector: "JSXAttribute[name.name='class']", message: 'No class attribute outside the design system: styling is sealed under src/components/** (hard rule 22).' },
        { selector: "JSXAttribute[name.name='style']", message: 'No inline style outside the design system: styling is sealed under src/components/** (hard rule 22).' },
      ],
    },
  },
  // Hard rule 37, the design system's own layers. After the rule-21 block on purpose: for
  // these files a zone is the no-restricted-imports that counts, so it carries rules 13 and
  // 21 as well (the smoke test's atom fixtures for both rules prove the copy survived).
  layerZone('components/atoms', ['molecules', 'organisms'], DESIGN_SYSTEM_BANS, UPWARD),
  layerZone('components/molecules', ['organisms'], DESIGN_SYSTEM_BANS, UPWARD),
  // Hard rule 37, the server archetype: the Bun config's zones, .tsx included, plus the UI
  // layers no server layer may reach. Inert in the static layout, which has none of these
  // directories.
  layerZone('domain', ['use-cases', 'infra', 'presenter', 'composition', 'test-helpers', 'lib', 'page', 'components'], [FS_AT_THE_EDGES]),
  layerZone('use-cases', ['infra', 'presenter', 'composition', 'test-helpers', 'lib', 'page', 'components'], [FS_AT_THE_EDGES]),
  layerZone('presenter', ['use-cases', 'infra', 'composition', 'test-helpers']),
  layerZone('infra', ['presenter', 'composition', 'test-helpers', 'page', 'components'], [INFRA_NO_USE_CASE]),
  layerZone('composition', ['test-helpers']),
  layerZone('test-helpers', ['infra']),
  pluginJs.configs.recommended,
  ...tsPlugin.configs.recommended,
  tailwind.configs.recommended,
  {
    settings: {
      tailwindcss: {
        cssConfigPath: './app/globals.css',
      },
    },
  },
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
  ...nextVitals,
  ...nextTs,
  globalIgnores(['eslint.config.mjs', '.next/**', 'out/**', 'build/**', 'next-env.d.ts', 'node_modules/**']),
]);

export default eslintConfig;
```

**Layer zones (hard rule 37).** `layerZone` blocks carry the dependency direction in both shapes this variant has: an atom never imports a molecule or an organism and a molecule never an organism (`references/atomic-design.md`, Imports point strictly upward; rule 21 seals the design system from app code and frameworks), and the server sub-variant's `src/{domain,use-cases,infra,presenter,composition,test-helpers}` get the Bun config's zones with `.tsx` included and the UI layers (`lib`, `page`, `components`) added to what a server layer may never reach. A zone replaces the design-system block's `no-restricted-imports` for its files, which is why every zone spreads `MOCK_BAN` and the atoms and molecules zones spread `DESIGN_SYSTEM_BANS`. The `domain` and `use-cases` zones also carry `FS_AT_THE_EDGES` (hard rule 20 in this variant): `fs` and `fs/promises` in either form are banned there and nowhere else, since Node is the runtime everywhere else in a Next package and the static layout has no server layers. The Next smoke test proves an atom importing a molecule, a molecule importing an organism and a domain file importing infra red, each with the rule number in the message.

Note: `no-console: 'error'` is the enforcement (hard rule 4); `next.config.ts` → `compiler.removeConsole` is defence-in-depth, not a substitute: a stripped `console.*` is a violation that silently vanished, which is why the lint rule exists. Log through the Winston module (below).

**Tailwind plugin (v4 API).** `eslint-plugin-tailwindcss` v4 exposes a single flat-config **object** at `tailwind.configs.recommended`: drop it in as one array element, do **not** spread `...tailwind.configs['flat/recommended']` (that key is a v3 artefact, is `undefined` in v4, and spreading `undefined` throws at config load). It reads its CSS entrypoint from the mandatory `settings.tailwindcss.cssConfigPath`, which accepts a **relative** path, so the old `config:` key and the `fileURLToPath`/`dirname` absolute-path dance are both gone. The `^4.0.2` range floats to the newest 4.x stable (a caret floats minors too): a caret anchored on a `-beta` tag (the previous `^4.0.0-beta.0`) still resolves to in-range **stable** releases, so it was already installing 4.0.x stable, which is exactly why the v3-era API above had to be corrected.

## `postcss.config.mjs`

```js
const postcssConfig = {
  plugins: { '@tailwindcss/postcss': {} },
};

export default postcssConfig;
```

No standalone `tailwind.config.{js,ts}`. Tailwind v4 config lives inside `app/globals.css` (CSS-first config).

Name the object before exporting: do not `export default { … }` anonymously. `eslint-config-next` enables `import/no-anonymous-default-export` (severity `warn`), and `postcss.config.mjs` is linted (it is not in `globalIgnores`), so under the `--max-warnings=0` gate the anonymous form **fails** the lint.

## `next.config.ts`

The variant marker and the home of four load-bearing behaviours: static export, console stripping, unoptimised images, and page extensions.

```ts
import type { NextConfig } from 'next';

const nextConfig: NextConfig = {
  output: 'export',
  pageExtensions: ['js', 'jsx', 'ts', 'tsx'],
  compiler: {
    removeConsole: process.env.NODE_ENV === 'production',
    reactRemoveProperties: process.env.NODE_ENV === 'production',
  },
  productionBrowserSourceMaps: false,
  images: {
    unoptimized: true, // required for static export
  },
};

export default nextConfig;
```

If this package sits inside a repo that has its own lockfile in a parent directory, `next build` may warn that it inferred the wrong workspace root. Silence it by pinning `turbopack.root` in the config: `turbopack: { root: import.meta.dirname }` (must be an **absolute** path; `import.meta.dirname` works in an ESM `next.config.ts` on Node ≥ 20.11 and points at the app dir). The official docs example points one level up (`path.join(import.meta.dirname, '..')`) for a true monorepo root, pick app-dir vs parent based on where the real workspace root / stray lockfile lives.

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
  "[javascript]": { "editor.defaultFormatter": "dbaeumer.vscode-eslint" },
  "[jsonc]": { "editor.defaultFormatter": "dbaeumer.vscode-eslint" },
  "[typescriptreact]": { "editor.defaultFormatter": "vscode.typescript-language-features" },
  "files.associations": { "*.css": "tailwindcss" },
  "[xml]": { "editor.defaultFormatter": "redhat.vscode-xml" },
  "[json]": { "editor.defaultFormatter": "vscode.json-language-features" }
}
```

## `.vscode/extensions.json`

```json
{
  "recommendations": [
    "dbaeumer.vscode-eslint",
    "bradlc.vscode-tailwindcss",
    "usernamehw.errorlens",
    "lacroixdavid1.vscode-format-context-menu",
    "kisstkondoros.vscode-codemetrics",
    "snyk-security.snyk-vulnerability-scanner",
    "sonarsource.sonarlint-vscode"
  ]
}
```

## `.gitignore`

```
# dependencies (bun install)
node_modules

# output
out
dist
*.tgz

# code coverage
coverage
*.lcov

# logs
logs
*.log
report.[0-9]*.[0-9]*.[0-9]*.[0-9]*.json

# dotenv
.env
.env.development.local
.env.test.local
.env.production.local
.env.local

# caches
.eslintcache
.cache
*.tsbuildinfo

# IDEs
.idea

# macOS
.DS_Store

# Snyk Security Extension - AI Rules (auto-generated)
.cursor/rules/snyk_rules.mdc
```

## `commitlint.config.mjs`

ESM like everything else (rule 9); commitlint 20 loads an `.mjs` config as is.

```js
export default {
  extends: ['@commitlint/config-conventional'],
  rules: {
    'body-max-line-length': [2, 'always', 200],
  },
};
```

## CI (`assets/ci-next.yml`)

The hook is the fast first line; CI is the gate set that `--no-verify` cannot skip, so make it the
required status check. The shipped workflow runs, on push to `main` and on pull requests, on a frozen
lockfile: the commit messages over the pushed range through commitlint (the hook's own grammar, one
grammar per variant, canon 1.3), the commit-size gate per commit (`scripts/check-commit-range.sh`),
gate 2 over every manifest (`scripts/check-package-json.sh`, rules 5 and 19), gitleaks over the full
history, the identity gate and the discipline wrapper over the whole tree (`--all`), then
`bun run --filter '*' test`, `lint`, `typecheck` and `build`, the README's Verify block
(`scripts/check-docs.sh`), and the bundle budget on
each `packages/*/out` (`scripts/check-bundle-size.sh`, `BUDGET_KB` in the workflow's `env`, canon 17.7).
No coverage or mutation step: this variant has neither gate (SKILL.md, What applies where).

The dependency CVE watchdog is the Bun variant's `assets/audit.yml`, copied as is: `bun audit` at the
workspace root covers every workspace, and it runs daily plus on any pull request that touches a
manifest or the lockfile (`references/workflow.md`, Dependency CVE scanning). Its second step,
`check-skill-pin.sh`, passes when the repo vendors no copy of the standard, so it is copied either way.

```bash
mkdir -p .github/workflows scripts
cp <skill>/assets/gitattributes          .gitattributes
cp <skill>/assets/ci-next.yml            .github/workflows/ci.yml
cp <skill>/assets/audit.yml              .github/workflows/audit.yml
cp <skill>/assets/check-skill-pin.sh     scripts/check-skill-pin.sh
cp <skill>/assets/check-commit-range.sh  scripts/check-commit-range.sh
cp <skill>/assets/land-on-green.sh      scripts/land-on-green.sh     # the CI land job: a green pull request lands itself (workflow.md, Landing on green)
cp <skill>/assets/check-branches.sh     scripts/check-branches.sh
cp <skill>/assets/check-host.py          scripts/check-host.py
cp <skill>/assets/host-ruleset.json      scripts/host-ruleset.json
cp <skill>/assets/branches.yml          .github/workflows/branches.yml   # rule 38's daily watchdog; after the first push the owner runs python3 scripts/check-host.py --apply once
cp <skill>/assets/check-package-json.sh  scripts/check-package-json.sh
cp <skill>/assets/check-bundle-size.sh   scripts/check-bundle-size.sh
cp <skill>/assets/check-docs.sh          scripts/check-docs.sh
cp <skill>/assets/check-identity.sh      scripts/check-identity.sh   # rule 26; a private repository also commits .github/visibility reading private, and the gate then warns
cp <skill>/assets/check-disciplines.sh   scripts/check-disciplines.sh
cp <skill>/assets/check-pii-channels.sh  scripts/check-pii-channels.sh
cp <skill>/assets/check-io-deadlines.sh  scripts/check-io-deadlines.sh
cp <skill>/assets/check-data-lifecycle.sh scripts/check-data-lifecycle.sh
cp <skill>/assets/check-reply.py         scripts/check-reply.py
mkdir -p .claude && cp <skill>/assets/claude-settings.json .claude/settings.json   # the reply gate; merge its Stop entry if the file exists
chmod +x scripts/*.sh
```

## Atomic Design (enforced: full doctrine in `references/atomic-design.md`)

Directory model under `packages/<app>/src/components/`:

- `atoms/`: no internal composition, only HTML elements and icons. Example: `button`, `badge`, `icons`.
- `molecules/`: may import atoms only. Example: `article-card`, `breadcrumbs`, `nav-header`, `language-switcher`.
- `organisms/`: may import atoms and molecules only. Example: `hero`, `faq`, `pricing`, `nav-bar`, `footer`.
- `src/page/`: page shells consumed by `app/(lang)/page.tsx`. May import any of the above and `src/lib/*`.

**Imports are strictly upward, and the design system is logic-free** (SKILL.md hard rule 21): components under `src/components/**` are stateless `const` arrow functions, no hooks, no fetching, no translation lookups, no `next/*` imports, no `'use client'`. State lives in `src/lib/hooks/` and is wired by page shells; links and images are injected as `ComponentType` props from `src/lib/layout/wrappers.tsx`.

**Styling is sealed inside it** (SKILL.md hard rule 22): Tailwind utilities exist only under `src/components/**`, design tokens in `app/globals.css`. Routes, page shells, `src/lib/**`, and `src/config/**` never carry a class string; component APIs expose typed variants, not `className`. The app does not know Tailwind exists.

Read `references/atomic-design.md` before any component work: layer table, component anatomy, interactivity ladder, injection pattern, styling seal, red flags.

## Testing (what TDD means in this variant)

Hard rule 11 still holds (no production logic without a failing test) and rules 21–22 are what make it tractable here: every line of logic lives in `src/lib/**` or `src/config/**`, so that is where the tests live.

- Runner: `bun test --randomize`, files `*.test.ts` next to source, exactly as in the Bun variant. The package script is `"test": "bun test --randomize"` (rule 36) and the root pre-commit runs it.
- **TDD-mandatory:** `src/lib/**` (i18n path helpers, guides/MDX utils, SEO builders, tag utils) and `src/config/**` factories. Red-Green-Refactor, domain-language test names.
- **Hooks stay thin.** A hook like `useNavState` is four lines of `useState` wiring: keep it that way. The moment a hook grows real logic (derivation, branching), extract that logic into a pure function in `src/lib/**` and TDD the function; the hook remains a trivial adapter.
- **Design-system components are not unit-tested.** Rule 21 makes them deterministic prop→JSX maps: no state, no IO, no business decisions, nothing worth owning a test. They are verified by the design-system ESLint block (above), review against `references/atomic-design.md`, and the build. Do not add React Testing Library ceremony to prove that props render.
- Page shells are wiring; when one accumulates a mapping (e.g. a `toPlanCard` transform), extract the mapping to `src/lib/**` and test it there.
- The mock ban (hard rule 13) applies: hand-written fakes, never `mock` from `bun:test`. The `no-restricted-imports` ban is already wired into the skeleton `eslint.config.mjs` above, and deliberately appears in **two** blocks: the general unscoped rules block and the design-system block. ESLint flat config **replaces, not merges** two `no-restricted-imports` declarations that match the same file: the later object wins outright. So the design-system block must carry its own `patterns` (the `next/*` and app-code bans) **and** the mock-ban `paths` in one object; a separate trailing ban object would silently drop the design-system bans for `src/components/**`.

Coverage tiers and Stryker mutation are Bun-variant gates; they do not run here (SKILL.md, "What applies where").

## Static-export data loading

The app builds with `output: 'export'` in `next.config.ts`. This means:

- All data must be available at build time.
- Server components in `app/(lang)/.../page.tsx` read from `data/` (MDX, JSON) at build time.
- They pass plain JSON-serialisable props down to client components (`src/page/<...>-page.tsx`, organisms, molecules).
- No runtime data fetching with `useEffect` or `fetch` in client components.
- MDX rendering goes through `next-mdx-remote` with custom components from `src/lib/guides/mdx-components.tsx`.
- Images served unoptimised (`images.unoptimized: true`). Required for static export.

## Next.js server app (sub-variant)

Everything above this point assumes the default shape: a **static content/marketing site** (`output: 'export'`, build-time data, no request-time code). When the app instead holds state or answers requests at runtime, route handlers like `POST /api/dossier`, an in-memory or DB-backed store, anything that reads a `Request`, you are building a **server app**, and the deltas below apply. The two shapes are **mutually exclusive**: `output: 'export'` emits static assets only and physically cannot run a request-time route handler (a build-time `GET` with no `Request` access is emitted as a static file; a `POST`, or any handler that reads the request, is not). An in-memory API needs a real server, so a server app drops static export.

This is the Next.js mirror of the Bun-script **Inbound HTTP (server archetype)**: read `references/architecture.md` § Inbound HTTP and the "Mapping errors to an HTTP status" paragraph of `references/result-type.md` for the shared rules. The route handler is just another **`infra/` inbound adapter**; the domain and use-cases stay free of `next/*`.

### Deltas to the skeleton

- **`next.config.ts`:** drop `output: 'export'` and `images.unoptimized` (the latter exists only to satisfy static export). Console stripping and page extensions stay. Keep the `turbopack.root` pin if the package is nested.
- **`package.json` scripts:** replace the static pair
  ```jsonc
  "build": "rimraf out && bun next build",   // static
  "start": "bunx serve ./out",               // static
  ```
  with the server pair, and drop the `rimraf` devDependency (the static `start` runs `bunx serve`, which fetches its latest on every run: pin it, `bunx serve@<version>`, if the static preview matters):
  ```jsonc
  "build": "bun next build",
  "start": "bun next start",
  ```
- **Vendoring Bun-script domain code** (a `Result` type, branded-id constructors, use-cases) is the normal way to share logic; the tsconfig above already allows its explicit `.ts` import extensions.
- **Logger:** server code uses the injected `Logger` **port** + Winston adapter + recording fake from the Bun variant (`references/bun-typescript.md` § Logger), **not** the client singleton. The rule-4 singleton exception is scoped to client components / static code only: a server app has a composition root, so inject the port (this is hard rule 4, not an exception to it).
- **Client-side data fetching goes through a gateway.** When page shells fetch at runtime (from the app's own route handlers or an external API), components never call `fetch` directly: a gateway port in `src/lib/` with a real client and a canned fake, returning `Result` and mapping the wire DTO into the frontend's own model at that one point (`references/architecture.md` § API shape, the frontend gateway). The static shape needs none of this: build-time loaders play that role.

### Route handler = inbound adapter

Validate the branded id at the URL boundary (rule 12), delegate to a composed use-case, and map the `Result` to an HTTP `Response`. Dynamic route `params` are **async in Next 16** (`Promise<…>`): `await` them. `export const dynamic = 'force-dynamic'` opts a stateful route out of static optimization.

The route handler **is** the inbound adapter (the Next equivalent of the Bun archetype's `src/infra/http/server.ts`), so it is the one sanctioned spot for a request-level `try/catch` if you need a safety net, but prefer total, `Result`-returning helpers so you usually don't.

```ts
// app/api/dossier/[id]/route.ts: thin inbound adapter, no business logic
import type { NextRequest } from 'next/server';
import { deps } from '@/src/composition/build-deps';
import { parseDossierId } from '@/src/domain/dossier-id';   // branded-id smart constructor
import { createGetDossier } from '@/src/use-cases/get-dossier';
import { toResponse } from '@/src/infra/http/to-response';

export const dynamic = 'force-dynamic';

export const GET = async (
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }   // Next 16: params is a Promise
): Promise<Response> => {
  const { id } = await params;
  const parsed = parseDossierId(id);
  // Precise client errors are decided HERE, at the branded checkpoint, where the
  // error is still narrow: a 400 before any IO runs (rule 12).
  if (!parsed.ok) return Response.json({ error: parsed.error.message }, { status: 400 });
  return toResponse(await createGetDossier(deps)(parsed.value));
};
```

`force-dynamic` caveat: it is the right knob in the **default** model. Under Cache Components (`cacheComponents: true`) the `dynamic` segment export is removed: delete it there. Any handler that reads the `Request` (and every non-`GET`) is already dynamic since Next 15, so `force-dynamic` is an explicit guarantee, not the only mechanism.

### Result → HTTP lives in `infra/`, not `presenter/`

This is the exact mapper from `references/architecture.md` § Inbound HTTP: read it for the full rationale. The `Result → Response` mapper must read `StepError` (which lives in `use-cases/ports/`), and the `presenter → domain/ only` dependency rule forbids a presenter from importing it, so the mapper is an **`infra/` inbound adapter**, not a presenter. A use-case failure **defaults to `500`**: the use-case flatten already stringified the port `kind` into `cause: string`, so there is no typed discriminant left to `switch` on (per `references/result-type.md` § Mapping errors to an HTTP status). The narrow client errors (`400`) were already decided upstream at the branded checkpoint, above. The mapper itself (`src/infra/http/to-response.ts`, pure and total, `ok` to 200, a use-case failure to 500 with `step`, `cause` and `message`) is printed once, in `references/architecture.md` § Inbound HTTP; copy it from there.

### Server-side body parsing without breaking the try/catch quarantine

`POST`/`PUT` handlers read the body. `JSON.parse` is a native synchronous thrower, and the branded-input smart constructor (the `400` checkpoint) needs it, so wrap it in a pure-domain helper that returns a `Result` (the sanctioned fallback pattern for native throwers from `references/result-type.md`). The branded constructor consumes it, and the route handler still carries no bare `try`.

```ts
// src/domain/safe-json-parse.ts, pure, no throw escapes
import type { Result } from './result.ts';
import { err, ok } from './result.ts';

export const safeJsonParse = (raw: string): Result<unknown, 'invalid-json'> => {
  try {
    return ok(JSON.parse(raw) as unknown);
  } catch {
    return err('invalid-json');
  }
};
```

### The in-session store is a composition-root singleton

Process-level server state (an in-memory store) is created **once** in the composition root (the one sanctioned place for wiring) behind a port, so it stays the persistence seam. Swapping the in-memory adapter for a DB-backed one later touches only this file; use-cases never change.

```ts
// src/composition/build-deps.ts: the one process-level wiring point
import { config } from '../config/env.ts';                 // typed-env, never process.env directly
import { createDossierStoreMemory } from '../infra/dossier-store-memory.ts';
import { createWinstonLogger } from '../infra/logger.ts';

// Built once at module load; route handlers import `deps` and never new-up infra.
export const deps = {
  store: createDossierStoreMemory(),                       // the DossierStore port's adapter
  logger: createWinstonLogger(config.logLevel),
} as const;
```

The store adapter is plain infra: `createDossierStoreMemory` returns `{ get, put }` whose arrows are object-property expressions, which is why the Next eslint config relaxes `explicit-function-return-type` with `allowTypedFunctionExpressions` (and why the Winston adapter's redaction loop needs `detect-object-injection` off): both are server-side idioms the static-only config never had to admit. Tests inject a fake store and the `createLoggerFake()` recorder and assert on outcomes, no mocking library, same discipline as the Bun variant.

## Internationalisation

Languages are Next.js route groups: `app/(en)`, `app/(es)`, `app/(fr)`, `app/(de)`, `app/(pt)`, `app/(zh)`, `app/(ja)`.

Translations are JSON files in `data/translations/` loaded by `src/lib/i18n/`. Each language has its own `page.tsx` / `layout.tsx`; shared shells live in `src/page/` and `src/lib/layout/`.

Every user-facing string lives in the catalog, keyed by meaning: error copy included, so a failure names its cause and next step in the user's language over a stable machine-readable code, and no prose is hardcoded in a component (`references/product.md`). Accessibility rides the same rails: semantic components, keyboard operability, and contrast-safe token pairs are design-system duties (`references/atomic-design.md`, Accessible by default).

## Secrets & config

No credentials in source. Centralise env reads in `src/config/env.ts` (the same location the server-app composition root imports from) rather than sprinkling `process.env` across modules (SKILL.md, Security). Besides `NODE_ENV`, the only env vars the app consumes are `LOG_LEVEL` and `LOG_FILE`: both read inside the sanctioned logger singleton below, where `NODE_ENV` is also read to pick prod-vs-dev formatting. Treat those reads as part of the singleton, not as a precedent for new code.

`.env*` is git-ignored.

## Winston logger

Location: `src/lib/utils/logger.ts`.

```ts
import winston from 'winston';

const { combine, timestamp, json, colorize, errors, printf } = winston.format;

const devFormat = printf(({ level, message, timestamp: ts, stack, ...meta }) => {
  const metaStr = Object.keys(meta).length > 0 ? ` ${JSON.stringify(meta)}` : '';
  return `${ts} [${level}]: ${stack ?? message}${metaStr}`;
});

const outputFormat = (): winston.Logform.Format => (process.env.NODE_ENV === 'production' ? json() : combine(colorize(), devFormat));

const logger = winston.createLogger({
  level: process.env.LOG_LEVEL ?? (process.env.NODE_ENV === 'production' ? 'info' : 'debug'),
  format: combine(errors({ stack: true }), timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }), outputFormat()),
  defaultMeta: { service: '<service-name>' },
  transports: [new winston.transports.Console({ format: outputFormat() })],
  exitOnError: false,
});

if (process.env.NODE_ENV === 'production' && process.env.LOG_FILE) {
  logger.add(new winston.transports.File({ filename: process.env.LOG_FILE, level: 'error' }));
}

export default logger;
```

Import as default: `import logger from '@/src/lib/utils/logger';`.

This module-level singleton is the **sanctioned rule-4 exception** for this variant (SKILL.md hard rule 4): static export plus the React client boundary leave no composition root through which to inject a `Logger` port into client components, so the variant trades injection for one well-known module. The exception is scoped to exactly that boundary, **client components and build-time/static code only**. It is *not* a licence to log through a singleton from server code: a Next.js **server app** (route handlers, use-cases, infra adapters; the server-app sub-variant above) has a real composition root, so it uses the injected `Logger` port + Winston adapter + recording fake exactly as the Bun-script variant does. One singleton at the client boundary; the port everywhere server-side. No other module-level service objects either way.

## Bootstrap checklist (new package in the monorepo)

1. From repo root: `mkdir -p packages/<NN>-<name> && cd packages/<NN>-<name>`.
2. `bun init -y`, then replace `package.json` with the skeleton above (rename `name`).
3. Create `tsconfig.json`, `eslint.config.mjs`, `postcss.config.mjs`, and `next.config.ts` with the blocks above.
4. Create `.vscode/settings.json` and `.vscode/extensions.json` at the repo root if not present.
5. From repo root: copy `assets/gitattributes` to `.gitattributes` (LF on every machine), then the CI workflows and their scripts as the CI section above shows (`assets/ci-next.yml` to `.github/workflows/ci.yml`, `assets/audit.yml` to `.github/workflows/audit.yml`; `check-skill-pin.sh`, `check-commit-range.sh`, `check-package-json.sh`, `check-docs.sh`, `check-identity.sh`, `check-disciplines.sh` with `check-pii-channels.sh`, `check-io-deadlines.sh` and `check-data-lifecycle.sh`, `check-bundle-size.sh` to `scripts/`, `chmod +x`; gate 2 is what the hook calls first), then `bun install`, then `bun run prepare` to install git hooks.
6. Create `src/lib/utils/logger.ts`.
7. Set up `app/globals.css` for Tailwind v4.
8. Lay out `src/components/{atoms,molecules,organisms}/`, `src/page/`, `src/lib/`, `src/config/`, `src/types/`.
9. Verify: `bun run --filter <package-name> test`, `bun run --filter <package-name> lint`, and `bun run --filter <package-name> build` all exit clean.
10. Commit with Conventional Commits: once the user confirms (rule 25).
