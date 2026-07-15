# Atomic Design — the logic-free design system

Applies to every repo with React UI (Next.js monorepo variant). Read this before creating or modifying anything under `src/components/`, `src/page/`, or `src/lib/{hooks,layout}/`.

The design system is a standalone catalogue of presentational components. It renders props. It decides nothing, fetches nothing, stores nothing, and imports nothing from the application around it. All intelligence — state, data loading, routing, i18n, analytics — lives outside the design system and arrives through props. And the wall runs both ways: just as no application knowledge enters the design system, no styling knowledge leaves it — Tailwind is invisible outside `src/components/**`. The Next.js side does not know the project uses Tailwind. This is SKILL.md hard rules 21 and 22, and it is non-negotiable.

Why the hard line:

- **Portability.** Components that depend only on `react` render anywhere: any router, Storybook, a marketing microsite, a test with no providers.
- **Refactor freedom.** The app can swap its state management, i18n library, or data source without touching a single component file — and redesign a component without touching behaviour.
- **Restyle freedom.** Styling lives only in the design system, so a full rebrand touches `src/components/**` and the design tokens — never a page, hook, route, or config file. Swapping the styling engine itself leaves the app byte-identical.
- **Reviewability.** A diff under `src/components/**` is a pure visual diff. If one PR changes a component *and* application behaviour, the boundary has leaked.

## The five levels, mapped to directories

Brad Frost's hierarchy (atoms → molecules → organisms → templates → pages) maps onto the repo like this:

| Level | Directory | Responsibility | May import |
|:---|:---|:---|:---|
| Atoms | `src/components/atoms/` | Smallest primitives: button, badge, icons. No internal composition — only HTML elements. | `react` only |
| Molecules | `src/components/molecules/` | Small groupings: article-card, breadcrumbs, section-header, language-switcher. | atoms, `react` |
| Organisms | `src/components/organisms/` | Full page sections: hero, faq, pricing, nav-bar, footer. | atoms, molecules, `react` |
| Templates | `src/lib/layout/` | Layout shells and framework wrappers shared across pages. | anything |
| Pages | `src/page/` | Page shells consumed by `app/(lang)/page.tsx`. Own all state and wiring. | anything |

**Imports point strictly upward.** An atom never imports a molecule. A molecule never imports an organism. Nothing inside `src/components/**` imports from `src/lib/**`, `src/config/**`, `src/page/**`, `app/**`, or any framework module (`next/link`, `next/image`, `next/navigation`). The only allowed imports inside the design system are `react` (types and JSX runtime) and lower design-system layers.

One component per kebab-case folder, component in `index.tsx`, PascalCase named export, exported props type, `displayName` set:

```
src/components/
├── atoms/
│   ├── button/index.tsx          # exports Button, ButtonProps
│   ├── badge/index.tsx
│   └── icons/arrow-right-icon.tsx
├── molecules/
│   └── article-card/index.tsx    # exports ArticleCard, ArticleCardProps
└── organisms/
    └── pricing/index.tsx         # exports Pricing, PricingProps
```

## The no-logic rule

Every component in `src/components/**` is a stateless `const` arrow function. Render output derives from props and nothing else.

Banned inside the design system:

- `useState`, `useReducer`, `useEffect`, `useContext`, `useRef` for behaviour — any hook that creates state or side effects.
- Data fetching, `async`, promises, timers.
- Translation lookups. Components receive final display strings (`title`, `label`, `description`) as props; the page shell resolves translations upstream.
- Business decisions. A component may map a typed prop to a class string (`variant → classes`); it may not decide *which* variant applies — that decision arrives as a prop.
- Imperative DOM access, `window`/`document`, global side effects.
- `dangerouslySetInnerHTML` on raw strings. If a prop is HTML, it crosses a `SanitizedHtml` checkpoint upstream (see `references/security.md`) — the prop type says so.

The test: could this component render in Storybook with nothing but hardcoded props? If anything else is needed — a provider, a router, an env var, a fetch — logic has leaked in.

### The interactivity ladder

"No logic" does not mean "no interactivity". Reach for these in order:

1. **Native HTML first.** Disclosure widgets are `<details>`/`<summary>`; styling reacts with CSS (`group-open:rotate-180`). Hover and focus states are CSS. Zero JavaScript, zero props, accessible by default.

```tsx
// organisms/faq — an accordion with no state anywhere
<details className="rounded-md border border-primary-200 p-6">
  <summary className="flex cursor-pointer list-none items-center justify-between">
    <span>{item.question}</span>
    <ChevronDownIcon className="transition-transform group-open:rotate-180" />
  </summary>
  <p className="mt-4">{item.answer}</p>
</details>
```

2. **Hoisted state via props.** When JS state is genuinely needed (mobile menu, dropdown), the component receives the state and its transitions as props — `isOpen: boolean` plus `onToggle: () => void` — and stays pure:

```tsx
export type LanguageSwitcherProps = {
  currentLang: string;
  languages: LanguageItem[];
  label?: string;
  isOpen: boolean;          // state lives upstream
  onToggle: () => void;     // transition lives upstream
};
```

3. **The state itself lives in `src/lib/hooks/`,** consumed by the page shell — never by a component:

```ts
// src/lib/hooks/use-nav-state.ts
export const useNavState = (): NavState => {
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [languageSwitcherIsOpen, setLanguageSwitcherIsOpen] = useState(false);
  return {
    mobileMenuOpen,
    languageSwitcherIsOpen,
    handleMobileMenuToggle: (): void => setMobileMenuOpen(!mobileMenuOpen),
    handleLanguageSwitcherToggle: (): void => setLanguageSwitcherIsOpen(!languageSwitcherIsOpen),
  };
};
```

Visibility toggles with classes (`hidden`, `block`, conditional class strings), not by mounting/unmounting whole subtrees where a class would do.

## Framework pieces are injected, never imported

Routing and image optimisation belong to the framework; the design system must not know which framework. Links and images arrive as component props typed against plain HTML attributes:

```tsx
// organism props — knows "a link goes here", not "Next.js exists"
export type NavBarProps = {
  brandImage: ComponentType<ImgHTMLAttributes<HTMLImageElement>>;
  brandLink: ComponentType<AnchorHTMLAttributes<HTMLAnchorElement>>;
  navLinkProps: { name: string; href: string }[];
  mobileMenuOpen: boolean;
  onMobileMenuToggle: () => void;
};
```

The adapters live in `src/lib/layout/wrappers.tsx` — the one place that imports `next/link` and `next/image`:

```tsx
// src/lib/layout/wrappers.tsx
export const createLinkWrapper = (href: string): ComponentType<AnchorHTMLAttributes<HTMLAnchorElement>> => {
  const isExternal = href.startsWith('http://') || href.startsWith('https://');
  const LinkWrapper: ComponentType<AnchorHTMLAttributes<HTMLAnchorElement>> = ({ children, ...props }) =>
    isExternal ? (
      <Link href={href} target="_blank" rel="noopener noreferrer" {...props}>{children}</Link>
    ) : (
      <Link href={href} {...props}>{children}</Link>
    );
  LinkWrapper.displayName = `LinkWrapper(${href})`;
  return LinkWrapper;
};
```

The page shell builds the wrappers and hands them down. Render via composition: `<Item.Link>children</Item.Link>`. A component that imports `createLinkWrapper` directly has coupled itself to `src/lib` — that is the leak this pattern exists to prevent.

## Styling is sealed inside the design system

The mirror image of the no-logic rule: the application never styles anything. Tailwind — the utility classes, the responsive grammar, the token scale — exists only under `src/components/**`, plus the token sheet `app/globals.css` (Tailwind v4 CSS-first config, the rebrand lever). To the rest of the codebase, "how things look" is not a concept it can express.

- **No utility class outside the design system.** `app/**` routes, `src/page/**` shells, `src/lib/**`, `src/config/**` never contain a Tailwind string. A page shell stacks organisms inside a bare `<main>`; each organism owns its own section spacing (`py-16 lg:py-20`), so the page has nothing left to say about layout.
- **No free-form `className`/`style` in public component APIs.** Molecules and organisms expose typed props — `variant`, `size`, `tone`, content — never a class-string escape hatch. If a caller "needs" to pass a class, the design system is missing a variant: add the variant there instead.
- **Leaf atoms are the one exception.** Icons and similar primitives may accept `className` so design-system parents can size and position them (`<ChevronDownIcon className="h-5 w-5" />`). That is internal composition between design-system layers; it never crosses the app boundary.
- **Copy is plain text.** Strings in `src/config/` and `data/translations/` carry no embedded class names and no styled JSX.
- **If it needs styling, it is a design-system component.** An MDX component map, a styled prose block, a styled link — the styled implementation lives under `src/components/**`; `src/lib` composes and wires it.

Two tests, applied at review:

1. **The rebrand test.** A full visual redesign touches `src/components/**` and the tokens in `app/globals.css` — nothing else shows up in the diff.
2. **The swap test.** Migrating Tailwind to another styling engine leaves `app/`, `src/page/`, `src/lib/`, and `src/config/` byte-identical.

Why: utility classes scattered through pages are how design drift starts — three slightly different paddings for the same kind of section, a rogue `mt-7` nobody can explain. One owner for the visual layer keeps every spacing decision reviewable in one directory, and keeps application diffs about behaviour, never pixels.

## Component anatomy

```tsx
// src/components/atoms/button/index.tsx
import type { ButtonHTMLAttributes, FC, ReactNode } from 'react';

export type ButtonVariant = 'primary' | 'secondary' | 'premium' | 'ghost';

export type ButtonProps = {
  children: ReactNode;
  variant?: ButtonVariant;
  icon?: ReactNode;
} & ButtonHTMLAttributes<HTMLButtonElement>;

const variantStyles: Record<ButtonVariant, string> = {
  primary: 'bg-primary-950 text-white hover:bg-primary-900',
  secondary: 'bg-transparent text-primary-900 border border-primary-300',
  premium: 'bg-accent-800 text-white hover:bg-accent-900',
  ghost: 'bg-transparent text-primary-700 hover:text-primary-950',
};

export const Button: FC<ButtonProps> = ({ children, variant = 'primary', icon, className = '', ...props }) => {
  const variantClass = ((): string => {
    switch (variant) {
      case 'secondary': return variantStyles.secondary;
      case 'premium': return variantStyles.premium;
      case 'ghost': return variantStyles.ghost;
      default: return variantStyles.primary;
    }
  })();

  return (
    <button className={`inline-flex items-center justify-center gap-x-2 rounded-md ${variantClass} ${className}`} {...props}>
      {icon && <span className="shrink-0">{icon}</span>}
      <span>{children}</span>
    </button>
  );
};

Button.displayName = 'Button';
```

Style rules, all enforced at review:

- `const` arrow function typed `FC<Props>`; never a `function` declaration (hard rule 2).
- Props type exported next to the component; `type`, never `interface` (hard rule 3).
- Group related inputs into named prop objects (`authProps`, `languageSwitcherProps`, `items`) instead of long flat lists.
- Destructure props in the parameter list; defaults in the parameter list.
- Variant/size dispatch through a typed `Record` map; look up via `switch` so `eslint-plugin-security`'s object-injection rule stays quiet without an inline ignore (hard rule 15).
- Semantic elements first: `header`, `nav` with `aria-label`, `section`, `ul`/`li` — not `div` soup.
- Lists render with `.map` and stable keys (id, slug, question text) — never the array index.
- Tailwind utilities on the design-token scale; responsive shifts via `md:*` / `lg:*`.
- Accessibility is part of the contract: visible focus rings, `aria-expanded`/`aria-haspopup` on disclosure triggers, `aria-hidden="true"` on decorative SVGs.
- Named imports only — no `import * as X` wildcards.
- `displayName` on every component (and on generated wrappers).
- The `className` merge on this atom is the leaf-atom exception to the styling seal: it exists so design-system parents can size and position the primitive. The application side never passes a class through it.

## The wiring: how data reaches the design system

```
app/(en)/page.tsx                 server component, build time
  loadTranslations('en')          ── reads data/translations/en.json
  getLandingPageConfig('en')      ── src/config: copy + hrefs + image configs, typed
        │ JSON-serialisable props
        ▼
src/page/home-page.tsx            'use client' page shell
  useNavState()                   ── owns ALL state (src/lib/hooks)
  createLinkWrapper / createImageWrapper (src/lib/layout/wrappers)
        │ props only: data + callbacks + injected components
        ▼
src/components/organisms → molecules → atoms     stateless, logic-free
```

The page shell is the composition root of the UI: it is the only `'use client'` boundary, the only consumer of hooks, and the only place design-system props get assembled. `src/config/` may import design-system **prop types** (`FeaturesProps`, `TestimonialsProps`) to stay in sync with the components it feeds — types flow downward, code never does.

## Accessible by default

A user who cannot see, hear, or use a mouse is still a user, and this is law that follows the user (the EU Accessibility Act applies since mid-2025). The design system is where accessibility is won or lost, because every screen inherits what the components do (full product-side doctrine: `references/product.md`):

- **Semantic elements first.** A `<button>`, never a clickable `div`; `<details>`/`<summary>` for disclosure; `<nav>`, `<main>`, real headings in order. The interactivity ladder's "native HTML first" is accessibility by default: focus, keyboard, and roles come free.
- **Keyboard everywhere.** Every interactive component operable by keyboard alone, focus visible (`focus-visible:` styles in the component), order following the DOM. A pointer-only interaction is a broken component.
- **Contrast lives in the tokens.** Token pairs in `globals.css` (`--color-primary` with `--color-on-primary`) are chosen once to pass WCAG contrast; components consume pairs, never mix-and-match raw colours, so compliance is inherited (rule 22).
- **Labels and states are props.** Inputs get labels, icon-only buttons get an `aria-label` prop, images get meaningful `alt` (or explicit `alt=""` when decorative), busy/expanded/selected states surface as `aria-*` driven by the same props that drive the visuals.
- **Error, empty, and loading are designed states**, not afterthoughts: components expose them as typed variants so a page shell with a failed `Result` renders an explicit state instead of a blank (`references/product.md`, error copy).
- **The gate is automated.** `eslint-plugin-jsx-a11y` runs error-level on `src/components/**` in the canonical Next config (`references/nextjs-monorepo.md`): it fails the build on a clickable `<div>`, a control with no accessible name, a missing `alt`, or an invalid anchor, and the Next smoke test proves those rules fire. It is AST-based, so it cannot judge contrast (no layout) or focus order; those stay with the design tokens (rule 22) and review, and a runtime axe scan is the optional deeper pass a team adds when it wants dynamic-state and contrast coverage.

## Where does it go?

| You are about to write… | It belongs in |
|:---|:---|
| A reusable visual primitive (button, badge, icon) | `src/components/atoms/` |
| A grouping of atoms with one purpose (card, breadcrumbs) | `src/components/molecules/` |
| A full page section (hero, pricing, footer) | `src/components/organisms/` |
| A Tailwind utility class | inside a component under `src/components/**` — nowhere else |
| A design token (colour, spacing scale, font) | `app/globals.css` (`@theme`) |
| `useState` / `useEffect` / any hook | `src/lib/hooks/`, consumed by `src/page/` |
| A `next/link` or `next/image` usage | `src/lib/layout/wrappers.tsx`, injected as props |
| Display copy, labels, hrefs | `data/translations/` + `src/config/`, resolved by the route |
| Data loading (MDX, JSON) | server components in `app/`, at build time |
| Page assembly, state wiring | `src/page/<name>-page.tsx` |
| SEO/structured data | `src/lib/seo/`, rendered by the page shell |

When a component seems to "need" something not in this table — a store, a context, a fetch — the need is real but the location is wrong: satisfy it in the page shell and pass the result down.

## Red flags (design system)

- A hook call — any `use*` — inside `src/components/**`.
- `import ... from '../../lib/...'`, `'@/src/config/...'`, `'next/link'`, or `'next/image'` anywhere under `src/components/**`. Links and images are injected as `ComponentType` props.
- A component that resolves translations, reads `process.env`, or touches `window`.
- `'use client'` on a design-system component. The directive belongs to page shells; pure components inherit the boundary.
- A `<div onClick>` where a `<button>` or `<details>` does the job natively.
- An interactive component a keyboard cannot operate, an icon-only control without an accessible name, or a raw colour pairing that sidesteps the contrast-safe token pairs.
- Conditional `null` returns to hide content where a `hidden`/responsive class is the honest tool.
- Index keys in a `.map`.
- A new component folder without an exported props type, without `displayName`, or holding two components.
- A page shell importing an atom directly to rebuild what an organism already provides — compose at the right level instead.
- A Tailwind utility string in `app/**` (anywhere but `globals.css`), `src/page/**`, `src/lib/**`, or `src/config/**`. Styling lives only under `src/components/**`.
- A molecule or organism whose public props include `className` or `style`, or a page shell passing one in. That is a missing variant — add it to the component.
- A class name embedded in config or translation strings.
