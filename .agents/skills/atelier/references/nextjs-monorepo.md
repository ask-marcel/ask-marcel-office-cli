# Next.js Monorepo Variant

Applies to repos with the Bun-workspace + Next.js 16 layout. Identifiable by `packages/*`, `next.config.ts`, and `app/(en)/` route groups.

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
    "pre-commit": "bun run --filter <package-name> lint",
    "commit-msg": "bunx --yes commitlint --edit $1"
  }
}
```

Activate hooks after install: `bun run prepare`.

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
    "lint": "eslint"
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
    "@types/mdx": "^2.0.13",
    "@types/node": "^20.19.27",
    "@types/react": "^19.2.7",
    "@types/react-dom": "^19.2.3",
    "@types/winston": "^2.4.4",
    "baseline-browser-mapping": "^2.9.11",
    "eslint": "^9.39.2",
    "eslint-config-next": "16.1.1",
    "eslint-plugin-jsx-a11y": "^6.10.2",
    "eslint-plugin-prettier": "^5.5.4",
    "eslint-plugin-react": "^7.37.5",
    "eslint-plugin-react-hooks": "^7.0.1",
    "eslint-plugin-security": "^3.0.1",
    "eslint-plugin-tailwindcss": "beta",
    "eslint-plugin-unicorn": "^61.0.2",
    "globals": "^17.0.0",
    "rimraf": "^6.1.2",
    "tailwindcss": "^4.1.18",
    "typescript": "^5.9.3",
    "typescript-eslint": "^8.51.0"
  },
  "ignoreScripts": ["sharp", "unrs-resolver"],
  "trustedDependencies": ["sharp", "unrs-resolver"],
  "browserslist": ["> 0.5%", "last 2 versions", "not dead", "not IE 11", "not op_mini all"]
}
```

No `format`, `test`, or `lint:fix` script. Save-in-editor triggers ESLint autofix.

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
    "resolveJsonModule": true,
    "isolatedModules": true,
    "jsx": "react-jsx",
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
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import tsPlugin from 'typescript-eslint';

const eslintConfig = defineConfig([
  securityPlugin.configs.recommended,
  {
    files: ['**/*.ts', '**/*.tsx'],
    rules: { 'security/detect-non-literal-fs-filename': 'off' },
  },
  {
    languageOptions: {
      globals: { ...globals.node, ...globals.browser },
    },
  },
  {
    rules: {
      'func-style': ['error', 'expression'],
      'no-restricted-syntax': ['off', 'ForOfStatement'],
      'prefer-template': 'error',
      quotes: ['error', 'single', { avoidEscape: true }],
    },
  },
  {
    rules: {
      '@typescript-eslint/explicit-function-return-type': 'error',
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
  pluginJs.configs.recommended,
  ...tsPlugin.configs.recommended,
  ...tailwind.configs['flat/recommended'],
  {
    settings: {
      tailwindcss: {
        config: `${dirname(fileURLToPath(import.meta.url))}/app/globals.css`,
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
  globalIgnores(['.next/**', 'out/**', 'build/**', 'next-env.d.ts', 'node_modules/**']),
]);

export default eslintConfig;
```

Note: `no-console` is not present here. Console calls are stripped from production builds via `next.config.ts` → `compiler.removeConsole`. Still use the Winston logger in code: calls to `console.*` silently disappear in prod.

## `postcss.config.mjs`

```js
export default {
  plugins: { '@tailwindcss/postcss': {} },
};
```

No standalone `tailwind.config.{js,ts}`. Tailwind v4 config lives inside `app/globals.css` (CSS-first config).

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

## Atomic Design (enforced)

Directory model under `packages/<app>/src/components/`:

- `atoms/`: no internal composition, only HTML elements and icons. Example: `button`, `badge`, `icons`.
- `molecules/`: may import atoms only. Example: `article-card`, `breadcrumbs`, `nav-header`, `language-switcher`.
- `organisms/`: may import atoms and molecules only. Example: `hero`, `faq`, `pricing`, `nav-bar`, `footer`.
- `src/page/`: page shells consumed by `app/(lang)/page.tsx`. May import any of the above and `src/lib/*`.

**Imports are strictly upward.** An atom never imports a molecule. A molecule never imports an organism. A page shell may import anything.

Every component gets a `displayName`:

```tsx
export const Button: FC<ButtonProps> = ({ children, variant = 'primary' }) => {
  return <button className={...}>{children}</button>;
};
Button.displayName = 'Button';
```

## Static-export data loading

The app builds with `output: 'export'` in `next.config.ts`. This means:

- All data must be available at build time.
- Server components in `app/(lang)/.../page.tsx` read from `data/` (MDX, JSON) at build time.
- They pass plain JSON-serialisable props down to client components (`src/page/<...>-page.tsx`, organisms, molecules).
- No runtime data fetching with `useEffect` or `fetch` in client components.
- MDX rendering goes through `next-mdx-remote` with custom components from `src/lib/guides/mdx-components.tsx`.
- Images served unoptimised (`images.unoptimized: true`). Required for static export.

## Internationalisation

Languages are Next.js route groups: `app/(en)`, `app/(es)`, `app/(fr)`, `app/(de)`, `app/(pt)`, `app/(zh)`, `app/(ja)`.

Translations are JSON files in `data/translations/` loaded by `src/lib/i18n/`. Each language has its own `page.tsx` / `layout.tsx`; shared shells live in `src/page/` and `src/lib/layout/`.

## Secrets & config

No credentials in source. Read from `process.env` inline. The only env vars the app currently consumes besides `NODE_ENV` are `LOG_LEVEL` and `LOG_FILE`. If you add many env vars later, centralise them in `src/lib/config/env.ts`.

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

## Bootstrap checklist (new package in the monorepo)

1. From repo root: `mkdir -p packages/<NN>-<name> && cd packages/<NN>-<name>`.
2. `bun init -y`, then replace `package.json` with the skeleton above (rename `name`).
3. Create `tsconfig.json`, `eslint.config.mjs`, `postcss.config.mjs` with the blocks above.
4. Create `.vscode/settings.json` and `.vscode/extensions.json` at the repo root if not present.
5. From repo root: `bun install`, then `bun run prepare` to install git hooks.
6. Create `src/lib/utils/logger.ts`.
7. Set up `app/globals.css` for Tailwind v4.
8. Lay out `src/components/{atoms,molecules,organisms}/`, `src/page/`, `src/lib/`, `src/config/`, `src/types/`.
9. Verify: `bun run --filter <package-name> lint` exits clean and `bun run --filter <package-name> build` succeeds.
10. Commit with Conventional Commits.
