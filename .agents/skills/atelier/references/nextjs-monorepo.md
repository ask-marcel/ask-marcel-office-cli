# Next.js Monorepo Variant

Applies to repos with the Bun-workspace + Next.js 16 layout. Identifiable by `packages/*`, `next.config.ts`, and `app/(en)/` route groups.

This reference describes two shapes that share the same toolchain: the **static content site** (the default — `output: 'export'`, build-time data, i18n route groups) which everything below assumes, and the **server app** (route handlers, runtime state) covered in its own sub-variant section. Pick the static shape unless the app must answer requests or hold state at runtime; the two are mutually exclusive (static export cannot run request-time route handlers).

## Workspace layout

```
<repo>/
├── package.json              # root workspace + commit hooks only
├── commitlint.config.cjs
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
    "pre-commit": "bun run --filter <package-name> test && bun run --filter <package-name> lint",
    "commit-msg": "bunx --yes commitlint --edit $1"
  }
}
```

Activate hooks after install: `bun run prepare`.

**This variant's hook mechanism is `simple-git-hooks`** (test + lint per package, commitlint on the message). The `.githooks/pre-commit` fast-gate hook from `references/workflow.md` belongs to the Bun-script variant, never install both: `core.hooksPath` and `simple-git-hooks` overwrite each other. The commit-size, package.json, and gitleaks gates are portable here if wanted; the coverage and mutation gates are not (see SKILL.md, "What applies where").

## Package `package.json`

```json
{
  "name": "<package-name>",
  "version": "0.1.0",
  "private": true,
  "scripts": {
    "dev": "bun next dev",
    "build": "rimraf out && bun next build",
    "start": "bunx serve ./out",
    "test": "bun test",
    "typecheck": "tsc --noEmit",
    "lint": "eslint --max-warnings=0"
  },
  "dependencies": {
    "next": "16.1.1",
    "next-mdx-remote": "^5.0.0",
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
    "baseline-browser-mapping": "^2.9.11",
    "eslint": "^9.39.2",
    "eslint-config-next": "16.1.1",
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
    "typescript": "^5.9.3",
    "typescript-eslint": "^8.51.0"
  },
  "trustedDependencies": ["sharp", "unrs-resolver"],
  "browserslist": ["> 0.5%", "last 2 versions", "not dead", "not IE 11", "not op_mini all"]
}
```

No `format` or `lint:fix` script — save-in-editor triggers ESLint autofix. The `test` and `typecheck` scripts are mandatory: `test` is the TDD gate (see Testing below) and `typecheck` (`tsc --noEmit`) is the standalone type gate the workflow loop calls — `next build` typechecks too, but you want the fast check without a full build. Notes on the skeleton: `@types/bun` ships the `bun:test` module types — without it, `import { describe, it, expect } from 'bun:test'` in the mandated test files does not type-resolve (the same caveat the Bun-script variant documents); leave the tsconfig `types` key **unset** so automatic `@types/*` inclusion still picks up `@types/react` etc. (an explicit `["bun"]` would suppress them), and restart the VS Code TS server after adding it. `@types/winston` must NOT be added (winston 3 ships its own types; the v2 stub conflicts), and `trustedDependencies` is Bun's lifecycle-script allowlist — there is no `ignoreScripts` package.json field.

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

`"jsx"` is Next-managed: Next runs its own JSX transform and rewrites this key on the first `dev`/`build` regardless of what you set (observed on 16.1.1: `preserve` becomes `react-jsx`, "next.js uses the React automatic runtime"). The build succeeds either way, so treat the value as owned by Next and do not fight the managed diff. (The Bun-script variant sets `react-jsx` explicitly because plain `tsc`, not Next, compiles it there.)

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

const eslintConfig = defineConfig([
  securityPlugin.configs.recommended,
  {
    files: ['**/*.ts', '**/*.tsx'],
    rules: {
      // false-positive-heavy security rules on this codebase's idioms; disabled
      // at project level. Never inline-ignore — change severity here or refactor.
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
      'no-console': 'error',
      'no-restricted-syntax': ['off', 'ForOfStatement'],
      'prefer-template': 'error',
      quotes: ['error', 'single', { avoidEscape: true }],
      // Mock ban (hard rule 13). Lives in this unscoped block (which precedes the
      // design-system block) AND is re-declared inside the design-system block:
      // ESLint flat config REPLACES — never merges — two `no-restricted-imports`
      // objects that match the same file, so each scope must carry its full set.
      'no-restricted-imports': [
        'error',
        {
          paths: [
            {
              name: 'bun:test',
              importNames: ['mock'],
              message: '`mock` from bun:test is forbidden — it leaks across test files. Use a hand-written fake (hard rule 13).',
            },
          ],
        },
      ],
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
          // mandatory — the general unscoped block's copy is overwritten for these files.
          paths: [
            {
              name: 'bun:test',
              importNames: ['mock'],
              message: '`mock` from bun:test is forbidden — it leaks across test files. Use a hand-written fake (hard rule 13).',
            },
          ],
          patterns: [
            { group: ['next', 'next/*'], message: 'Design-system components import react only — inject links/images as ComponentType props (hard rule 21).' },
            { group: ['**/lib/**', '**/config/**', '**/page/**'], message: 'Design-system components must not import application code (hard rule 21).' },
          ],
        },
      ],
      'no-restricted-syntax': [
        'error',
        { selector: 'CallExpression[callee.name=/^use[A-Z]/]', message: 'No hooks inside the design system — hoist state to the page shell via src/lib/hooks (hard rule 21).' },
        { selector: 'Program > ExpressionStatement[directive="use client"]', message: "The 'use client' boundary belongs to page shells, not design-system components (hard rule 21)." },
      ],
      // Accessible by default (rule 17.6): the design system is where a11y is won or
      // lost, so the doctrine's structural rules are error-level here. The jsx-a11y plugin
      // is registered by next/core-web-vitals (spread below), which enables its recommended
      // subset app-wide but NOT the interaction rules that catch the flagship clickable-div;
      // these reference that same plugin and make them explicit and unmissable. Contrast is
      // not lintable (needs layout) — it lives in the design tokens (rule 22) and review.
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
    // Routes, page shells, lib, and config never carry a class string — visual variation
    // is a typed variant prop on a design-system component, never free-form className/
    // style outside src/components/**. Rule 21 bans the imports; this bans the styling leak.
    files: ['app/**/*.tsx', 'src/page/**/*.tsx', 'src/lib/**/*.tsx', 'src/config/**/*.tsx'],
    rules: {
      'no-restricted-syntax': [
        'error',
        { selector: "JSXAttribute[name.name='className']", message: 'No className outside the design system — Tailwind is sealed under src/components/** (hard rule 22). Move the styling into a design-system component with a typed variant.' },
        { selector: "JSXAttribute[name.name='class']", message: 'No class attribute outside the design system — styling is sealed under src/components/** (hard rule 22).' },
        { selector: "JSXAttribute[name.name='style']", message: 'No inline style outside the design system — styling is sealed under src/components/** (hard rule 22).' },
      ],
    },
  },
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

Note: `no-console: 'error'` is the enforcement (hard rule 4); `next.config.ts` → `compiler.removeConsole` is defence-in-depth, not a substitute — a stripped `console.*` is a violation that silently vanished, which is why the lint rule exists. Log through the Winston module (below).

**Tailwind plugin (v4 API).** `eslint-plugin-tailwindcss` v4 exposes a single flat-config **object** at `tailwind.configs.recommended` — drop it in as one array element, do **not** spread `...tailwind.configs['flat/recommended']` (that key is a v3 artefact, is `undefined` in v4, and spreading `undefined` throws at config load). It reads its CSS entrypoint from the mandatory `settings.tailwindcss.cssConfigPath`, which accepts a **relative** path — so the old `config:` key and the `fileURLToPath`/`dirname` absolute-path dance are both gone. The `^4.0.2` pin floats to the current 4.0.x stable: a caret anchored on a `-beta` tag (the previous `^4.0.0-beta.0`) still resolves to in-range **stable** releases, so it was already installing 4.0.x stable — which is exactly why the v3-era API above had to be corrected.

## `postcss.config.mjs`

```js
const postcssConfig = {
  plugins: { '@tailwindcss/postcss': {} },
};

export default postcssConfig;
```

No standalone `tailwind.config.{js,ts}`. Tailwind v4 config lives inside `app/globals.css` (CSS-first config).

Name the object before exporting — do not `export default { … }` anonymously. `eslint-config-next` enables `import/no-anonymous-default-export` (severity `warn`), and `postcss.config.mjs` is linted (it is not in `globalIgnores`), so under the `--max-warnings=0` gate the anonymous form **fails** the lint.

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

If this package sits inside a repo that has its own lockfile in a parent directory, `next build` may warn that it inferred the wrong workspace root. Silence it by pinning `turbopack.root` in the config — `turbopack: { root: import.meta.dirname }` (must be an **absolute** path; `import.meta.dirname` works in an ESM `next.config.ts` on Node ≥ 20.11 and points at the app dir). The official docs example points one level up (`path.join(import.meta.dirname, '..')`) for a true monorepo root — pick app-dir vs parent based on where the real workspace root / stray lockfile lives.

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

## `commitlint.config.cjs`

```js
module.exports = {
  extends: ['@commitlint/config-conventional'],
  rules: {
    'body-max-line-length': [2, 'always', 200],
  },
};
```

## Atomic Design (enforced — full doctrine in `references/atomic-design.md`)

Directory model under `packages/<app>/src/components/`:

- `atoms/`: no internal composition, only HTML elements and icons. Example: `button`, `badge`, `icons`.
- `molecules/`: may import atoms only. Example: `article-card`, `breadcrumbs`, `nav-header`, `language-switcher`.
- `organisms/`: may import atoms and molecules only. Example: `hero`, `faq`, `pricing`, `nav-bar`, `footer`.
- `src/page/`: page shells consumed by `app/(lang)/page.tsx`. May import any of the above and `src/lib/*`.

**Imports are strictly upward, and the design system is logic-free** (SKILL.md hard rule 21): components under `src/components/**` are stateless `const` arrow functions — no hooks, no fetching, no translation lookups, no `next/*` imports, no `'use client'`. State lives in `src/lib/hooks/` and is wired by page shells; links and images are injected as `ComponentType` props from `src/lib/layout/wrappers.tsx`.

**Styling is sealed inside it** (SKILL.md hard rule 22): Tailwind utilities exist only under `src/components/**`, design tokens in `app/globals.css`. Routes, page shells, `src/lib/**`, and `src/config/**` never carry a class string; component APIs expose typed variants, not `className`. The app does not know Tailwind exists.

Read `references/atomic-design.md` before any component work — layer table, component anatomy, interactivity ladder, injection pattern, styling seal, red flags.

## Testing (what TDD means in this variant)

Hard rule 11 still holds — no production logic without a failing test — and rules 21–22 are what make it tractable here: every line of logic lives in `src/lib/**` or `src/config/**`, so that is where the tests live.

- Runner: `bun test`, files `*.test.ts` next to source, exactly as in the Bun variant. The package script is `"test": "bun test"` and the root pre-commit runs it.
- **TDD-mandatory:** `src/lib/**` (i18n path helpers, guides/MDX utils, SEO builders, tag utils) and `src/config/**` factories. Red-Green-Refactor, domain-language test names.
- **Hooks stay thin.** A hook like `useNavState` is four lines of `useState` wiring — keep it that way. The moment a hook grows real logic (derivation, branching), extract that logic into a pure function in `src/lib/**` and TDD the function; the hook remains a trivial adapter.
- **Design-system components are not unit-tested.** Rule 21 makes them deterministic prop→JSX maps: no state, no IO, no business decisions — nothing worth owning a test. They are verified by the design-system ESLint block (above), review against `references/atomic-design.md`, and the build. Do not add React Testing Library ceremony to prove that props render.
- Page shells are wiring; when one accumulates a mapping (e.g. a `toPlanCard` transform), extract the mapping to `src/lib/**` and test it there.
- The mock ban (hard rule 13) applies: hand-written fakes, never `mock` from `bun:test`. The `no-restricted-imports` ban is already wired into the skeleton `eslint.config.mjs` above, and deliberately appears in **two** blocks — the general unscoped rules block and the design-system block. ESLint flat config **replaces, not merges** two `no-restricted-imports` declarations that match the same file: the later object wins outright. So the design-system block must carry its own `patterns` (the `next/*` and app-code bans) **and** the mock-ban `paths` in one object; a separate trailing ban object would silently drop the design-system bans for `src/components/**`.

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

Everything above this point assumes the default shape: a **static content/marketing site** (`output: 'export'`, build-time data, no request-time code). When the app instead holds state or answers requests at runtime — route handlers like `POST /api/dossier`, an in-memory or DB-backed store, anything that reads a `Request` — you are building a **server app**, and the deltas below apply. The two shapes are **mutually exclusive**: `output: 'export'` emits static assets only and physically cannot run a request-time route handler (a build-time `GET` with no `Request` access is emitted as a static file; a `POST`, or any handler that reads the request, is not). An in-memory API needs a real server, so a server app drops static export.

This is the Next.js mirror of the Bun-script **Inbound HTTP (server archetype)** — read `references/architecture.md` § Inbound HTTP and `references/result-type.md` § Inbound HTTP for the shared rules. The route handler is just another **`infra/` inbound adapter**; the domain and use-cases stay free of `next/*`.

### Deltas to the skeleton

- **`next.config.ts`:** drop `output: 'export'` and `images.unoptimized` (the latter exists only to satisfy static export). Console stripping and page extensions stay. Keep the `turbopack.root` pin if the package is nested.
- **`package.json` scripts:** replace the static pair
  ```jsonc
  "build": "rimraf out && bun next build",   // static
  "start": "bunx serve ./out",               // static
  ```
  with the server pair, and drop the `rimraf` / `serve` devDeps:
  ```jsonc
  "build": "bun next build",
  "start": "bun next start",
  ```
- **Vendoring Bun-script domain code** (a `Result` type, branded-id constructors, use-cases) is the normal way to share logic — add `"allowImportingTsExtensions": true` to `tsconfig.json` if that code imports with explicit `.ts` extensions (see the tsconfig note above).
- **Logger:** server code uses the injected `Logger` **port** + Winston adapter + recording fake from the Bun variant (`references/bun-typescript.md` § Logger), **not** the client singleton. The rule-4 singleton exception is scoped to client components / static code only — a server app has a composition root, so inject the port (this is hard rule 4, not an exception to it).
- **Client-side data fetching goes through a gateway.** When page shells fetch at runtime (from the app's own route handlers or an external API), components never call `fetch` directly: a gateway port in `src/lib/` with a real client and a canned fake, returning `Result` and mapping the wire DTO into the frontend's own model at that one point (`references/architecture.md` § API shape, the frontend gateway). The static shape needs none of this: build-time loaders play that role.

### Route handler = inbound adapter

Validate the branded id at the URL boundary (rule 12), delegate to a composed use-case, and map the `Result` to an HTTP `Response`. Dynamic route `params` are **async in Next 16** (`Promise<…>`) — `await` them. `export const dynamic = 'force-dynamic'` opts a stateful route out of static optimization.

The route handler **is** the inbound adapter (the Next equivalent of the Bun archetype's `src/infra/http/server.ts`), so it is the one sanctioned spot for a request-level `try/catch` if you need a safety net — but prefer total, `Result`-returning helpers so you usually don't.

```ts
// app/api/dossier/[id]/route.ts — thin inbound adapter, no business logic
import type { NextRequest } from 'next/server';
import { deps } from '@/src/composition/build-deps';
import { parseDossierId } from '@/src/domain/dossier-id';   // branded-id smart constructor
import { getDossier } from '@/src/use-cases/get-dossier';
import { toResponse } from '@/src/infra/http/to-response';

export const dynamic = 'force-dynamic';

export const GET = async (
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }   // Next 16: params is a Promise
): Promise<Response> => {
  const { id } = await params;
  const parsed = parseDossierId(id);
  // Precise client errors are decided HERE, at the branded checkpoint, where the
  // error is still narrow — a 400 before any IO runs (rule 12).
  if (!parsed.ok) return Response.json({ error: parsed.error.message }, { status: 400 });
  return toResponse(await getDossier(deps)(parsed.value));
};
```

`force-dynamic` caveat: it is the right knob in the **default** model. Under Cache Components (`cacheComponents: true`) the `dynamic` segment export is removed — delete it there. Any handler that reads the `Request` (and every non-`GET`) is already dynamic since Next 15, so `force-dynamic` is an explicit guarantee, not the only mechanism.

### Result → HTTP lives in `infra/`, not `presenter/`

This is the exact mapper from `references/architecture.md` § Inbound HTTP — read it for the full rationale. The `Result → Response` mapper must read `StepError` (which lives in `use-cases/ports/`), and the `presenter → domain/ only` dependency rule forbids a presenter from importing it — so the mapper is an **`infra/` inbound adapter**, not a presenter. A use-case failure **defaults to `500`**: the use-case flatten already stringified the port `kind` into `cause: string`, so there is no typed discriminant left to `switch` on (per `references/result-type.md` § Mapping errors to an HTTP status). The narrow client errors (`400`) were already decided upstream at the branded checkpoint, above.

```ts
// src/infra/http/to-response.ts — pure, total: a use-case Result → an HTTP Response
import type { Result } from '../../domain/result.ts';
import type { StepError, Summary } from '../../use-cases/ports/step-error.ts';

export const toResponse = <T extends Summary>(result: Result<T, StepError>): Response => {
  if (result.ok) return Response.json(result.value, { status: 200 });
  const { step, cause, message } = result.error;
  // `cause` is a plain string after the flatten — no typed discriminant to switch on,
  // so a use-case failure is a 500 by default. 400s are handled at the branded checkpoint.
  return Response.json({ step, error: cause, message }, { status: 500 });
};
```

### Server-side body parsing without breaking the try/catch quarantine

`POST`/`PUT` handlers read the body. `JSON.parse` is a native synchronous thrower, and the branded-input smart constructor (the `400` checkpoint) needs it — so wrap it in a pure-domain helper that returns a `Result` (the sanctioned fallback pattern for native throwers from `references/result-type.md`). The branded constructor consumes it, and the route handler still carries no bare `try`.

```ts
// src/domain/safe-json-parse.ts — pure, no throw escapes
import { err, ok, type Result } from './result.ts';

export const safeJsonParse = (raw: string): Result<unknown, 'invalid-json'> => {
  try {
    return ok(JSON.parse(raw) as unknown);
  } catch {
    return err('invalid-json');
  }
};
```

### The in-session store is a composition-root singleton

Process-level server state (an in-memory store) is created **once** in the composition root — the one sanctioned place for wiring — behind a port, so it stays the persistence seam. Swapping the in-memory adapter for a DB-backed one later touches only this file; use-cases never change.

```ts
// src/composition/build-deps.ts — the one process-level wiring point
import { config } from '../config/env.ts';                 // typed-env, never process.env directly
import { createDossierStoreMemory } from '../infra/dossier-store-memory.ts';
import { createWinstonLogger } from '../infra/logger.ts';

// Built once at module load; route handlers import `deps` and never new-up infra.
export const deps = {
  store: createDossierStoreMemory(),                       // the DossierStore port's adapter
  logger: createWinstonLogger(config.logLevel),
} as const;
```

The store adapter is plain infra — `createDossierStoreMemory` returns `{ get, put }` whose arrows are object-property expressions, which is why the Next eslint config relaxes `explicit-function-return-type` with `allowTypedFunctionExpressions` (and why the Winston adapter's redaction loop needs `detect-object-injection` off): both are server-side idioms the static-only config never had to admit. Tests inject a fake store and the `createLoggerFake()` recorder and assert on outcomes — no mocking library, same discipline as the Bun variant.

## Internationalisation

Languages are Next.js route groups: `app/(en)`, `app/(es)`, `app/(fr)`, `app/(de)`, `app/(pt)`, `app/(zh)`, `app/(ja)`.

Translations are JSON files in `data/translations/` loaded by `src/lib/i18n/`. Each language has its own `page.tsx` / `layout.tsx`; shared shells live in `src/page/` and `src/lib/layout/`.

Every user-facing string lives in the catalog, keyed by meaning: error copy included, so a failure names its cause and next step in the user's language over a stable machine-readable code, and no prose is hardcoded in a component (`references/product.md`). Accessibility rides the same rails: semantic components, keyboard operability, and contrast-safe token pairs are design-system duties (`references/atomic-design.md`, Accessible by default).

## Secrets & config

No credentials in source. Centralise env reads in `src/config/env.ts` (the same location the server-app composition root imports from) rather than sprinkling `process.env` across modules (SKILL.md, Security). Besides `NODE_ENV`, the only env vars the app consumes are `LOG_LEVEL` and `LOG_FILE` — both read inside the sanctioned logger singleton below, where `NODE_ENV` is also read to pick prod-vs-dev formatting. Treat those reads as part of the singleton, not as a precedent for new code.

`.env*` is git-ignored.

## Winston logger

Location: `src/lib/utils/logger.ts`.

```ts
import winston from 'winston';

const { combine, timestamp, json, colorize, errors, printf } = winston.format;

const devFormat = printf(({ level, message, timestamp: ts, stack, ...meta }) => {
  const metaStr = Object.keys(meta).length ? ` ${JSON.stringify(meta)}` : '';
  return `${ts} [${level}]: ${stack ?? message}${metaStr}`;
});

const logger = winston.createLogger({
  level: process.env.LOG_LEVEL ?? (process.env.NODE_ENV === 'production' ? 'info' : 'debug'),
  format: combine(
    errors({ stack: true }),
    timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
    process.env.NODE_ENV === 'production' ? json() : combine(colorize(), devFormat)
  ),
  defaultMeta: { service: '<service-name>' },
  transports: [
    new winston.transports.Console({
      format: process.env.NODE_ENV === 'production' ? json() : combine(colorize(), devFormat),
    }),
  ],
  exitOnError: false,
});

if (process.env.NODE_ENV === 'production' && process.env.LOG_FILE) {
  logger.add(new winston.transports.File({ filename: process.env.LOG_FILE, level: 'error' }));
}

export default logger;
```

Import as default: `import logger from '@/src/lib/utils/logger';`.

This module-level singleton is the **sanctioned rule-4 exception** for this variant (SKILL.md hard rule 4): static export plus the React client boundary leave no composition root through which to inject a `Logger` port into client components, so the variant trades injection for one well-known module. The exception is scoped to exactly that boundary — **client components and build-time/static code only**. It is *not* a licence to log through a singleton from server code: a Next.js **server app** (route handlers, use-cases, infra adapters — see the server-app sub-variant below) has a real composition root, so it uses the injected `Logger` port + Winston adapter + recording fake exactly as the Bun-script variant does. One singleton at the client boundary; the port everywhere server-side. No other module-level service objects either way.

## Bootstrap checklist (new package in the monorepo)

1. From repo root: `mkdir -p packages/<NN>-<name> && cd packages/<NN>-<name>`.
2. `bun init -y`, then replace `package.json` with the skeleton above (rename `name`).
3. Create `tsconfig.json`, `eslint.config.mjs`, `postcss.config.mjs`, and `next.config.ts` with the blocks above.
4. Create `.vscode/settings.json` and `.vscode/extensions.json` at the repo root if not present.
5. From repo root: `bun install`, then `bun run prepare` to install git hooks.
6. Create `src/lib/utils/logger.ts`.
7. Set up `app/globals.css` for Tailwind v4.
8. Lay out `src/components/{atoms,molecules,organisms}/`, `src/page/`, `src/lib/`, `src/config/`, `src/types/`.
9. Verify: `bun run --filter <package-name> test`, `bun run --filter <package-name> lint`, and `bun run --filter <package-name> build` all exit clean.
10. Commit with Conventional Commits — once the user confirms (rule 25).
