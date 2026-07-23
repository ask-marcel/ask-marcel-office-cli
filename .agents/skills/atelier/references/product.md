# Product experience and validation (build the right thing, whole)

You can satisfy every engineering rule and still build something people quietly abandon. The experience the user actually has is the whole product: how fast it responds, how it treats a mistake, whether there is a human when the bot fails, whether it works without a mouse. And before any of that: whether the thing was worth building at all. This reference binds user-facing work (error copy, flows, accessibility) and the validation moments before and after building.

## The whole journey is the product

Speed, payments, delivery, error messages, and support are the experience, not adjuncts to it.

**Error copy names the cause and the next step**, in the app's voice, from the i18n catalog; the machine-readable shape sits underneath so clients branch on a code, never on prose. This is where the `Result` discipline surfaces to humans: a discriminated-union error maps to a stable `code`, and the catalog maps the code to words a person can act on.

```tsx
// BAD: dumps the transport at the person
<p>Error 413</p>

// GOOD: cause + next step, from the catalog; the API shape behind it is stable
<Callout tone="warning">{t('receipt.tooLarge')}</Callout>
// "This photo is over 10 MB. Retake it or shrink it."
// wire shape: { "error": { "code": "receipt_too_large", "maxBytes": 10485760 } }
```

Never leak a stack trace, an internal path, or an SQL message to the user (`references/security.md`, Data exposure): generic outside, detailed in the server log with the trace id (`references/observability.md`). Error, empty, and loading states are designed states of every screen, not afterthoughts: the design system exposes them as variants (rule 21; a gateway returning `Result` gives the page shell an explicit error to render instead of a blank screen).

## Earn trust rather than extract a sale

Being honest with users, even when it costs a conversion, is what makes them come back; a short-term win that erodes trust is the most expensive kind.

- Cancelling is as easy as subscribing: self-serve, symmetric, keeps access already paid for. Forcing a phone call to cancel extracts a month of resentment, not a renewal.
- No dark patterns: no pre-ticked consent, no buried exits, no guilt-tripping copy in the catalog.
- Consent and privacy defaults honour the user (`references/privacy.md`); accessibility is not gated behind a plan.

## Design for real behaviour, not the demo

Ground flows in how real people behave, which differs sharply by market and culture, rather than what looks good in a pitch. Watch what users do, not what they say they would do.

```ts
// BAD: the demo audience pays by card, so card is hardcoded for the world
const methods = ['card'];

// GOOD: defaults come from measured completion per market, re-ranked on evidence
const methods = paymentMethods(ctx.market); // config-driven (framework vs configuration, architecture.md)
analytics.track('checkout_started', { market: ctx.market, first: methods[0] }); // signal only, no PII (rule 27)
```

Market-varying data is configuration, never hardcoded unions in framework code (`references/architecture.md`, Framework vs configuration).

## Technology serves the person

Automation removes friction and gives people their time back; it does not paper over a worse experience. The human path is always visible: a support bot renders the "talk to someone" link permanently, never behind N failed bot turns. The human touch is the part a competitor cannot clone in a sprint; protect it.

## Speak the user's language

Every user-facing string lives in a catalog keyed by meaning, never hardcoded in a component, so localization is a data change. The Next.js variant already enforces this shape (`data/translations/` + `src/lib/i18n/`, `references/nextjs-monorepo.md`); the rule 21 design system takes display strings as props, which is what makes the catalog the single source.

## Accessible by default

A user who cannot see, hear, or use a mouse is still a user, and like privacy this is law that follows the user (the EU Accessibility Act applies since mid-2025).

- **Semantic elements first**: a `<button>`, not a clickable `div`; the interactivity ladder's "native HTML first" (`references/atomic-design.md`) is accessibility by default.
- **Every flow workable by keyboard**: focus visible, order sensible, no pointer-only interactions.
- **Contrast lives in the design tokens**: token pairs (`--color-primary` / `--color-on-primary`) are chosen to pass WCAG once, in `globals.css`, so components inherit compliance (rule 22).
- **Labels and states**: inputs labelled, images with meaningful `alt`, busy/expanded states via `aria-*` driven by props.
- **Automated checks in the gate**: `eslint-plugin-jsx-a11y` runs error-level on the design system (`references/atomic-design.md`, Accessible by default; `references/nextjs-monorepo.md`), failing the build on the structural violations (clickable div, missing label, missing alt) exactly as any lint rule does. Contrast and focus order are not lintable and stay with tokens plus review; a runtime axe scan is the optional deeper pass.

```tsx
// BAD: no role, no keyboard path, contrast by vibe
<div className="btn" onClick={submit} style={{ color: '#9ca3af' }}>Submit</div>

// GOOD: real button; label from the catalog; contrast from the tokens
<button type="submit" className="btn" aria-busy={saving}>{t('expense.submit')}</button>
```

## Mobile first, and a light interface

Design the **smallest screen first** with **one clear primary action per view**, and let each screen carry only the controls that view needs. Starting from the phone forces the ruthless prioritization that leaves every surface simpler; starting from a wide desktop and shrinking it never does, and the phone inherits a dense mess in a thumb's reach.

```tsx
// DON'T: desktop-first wall of columns, a toolbar of eight equal buttons, no clear next step
<div className="grid grid-cols-4 gap-8 p-12">{/* eight equal buttons */}</div>
// DO: one column by default, one primary action, the rest one tap away
<div className="grid grid-cols-1 gap-4 p-4 md:grid-cols-4">
  <button className="btn-primary">{t('expense.submit')}</button>
  <OverflowMenu items={secondaryActions} />
</div>
```

The small screen is the default case, not the exception: breakpoints scale **up** (`md:` / `lg:`) from the base layout, tap targets are finger-sized (~44px), the primary action sits in thumb reach, and the form asks the fewest fields it can (rule 2.6). A light interface is a light payload too, so make weight a number the pipeline enforces: the shipped `assets/check-bundle-size.sh` (or `size-limit` / Lighthouse CI budgets) fails the build when the built JS crosses its gzipped ceiling, measured on the shipped bundle, not a developer laptop, and a bump is a deliberate reviewed change (pillar 12, and rule 15.1 for the gate).

## Validate before you build

The most expensive software is the beautifully built kind nobody needed. Every other rule makes you good at building the thing right; this one checks it is the right thing. It comes first in time, and again at every fork. The atelier-grill-me skill owns the interactive version of this moment; these are the artifacts it should leave behind.

- **Talk to real users first.** Short problem interviews about the last real occurrence of the pain, in their words: "walk me through the last time", "where did it get slow", "what did you do to work around it", "how often". Ask about their problem, never pitch your idea; leading questions harvest polite yeses.
- **Test demand with the cheapest thing that carries it.** A landing page with an email capture, a mockup, a concierge version done by hand. The address goes to your own endpoint; the analytics event carries only the signal (`references/privacy.md`, Third parties).
- **Set a dated, honest go/no-go.** Criteria written before the evidence comes in, a decision on the record, an unmet criterion stated rather than smoothed over:

```markdown
# Go / No-Go: <feature>: decided YYYY-MM-DD
- [x] >= 10 problem interviews surfaced this pain unprompted   (met: 12)
- [ ] >= 3 committed paid pilots                               (NOT met: 1)
Decision: NO-GO. Re-decide by <date>. Owner: <name>.
```

- **Keep validating after launch.** Ship behind a flag, instrument adoption (`references/observability.md`, Watch behaviour), and keep or kill on a measured threshold. "Shipped" is an output; "used" is the outcome. A feature below the bar gets disabled, its reason written down, and its code removed: no zombie features half-live in the product. Killing a plausible idea before it becomes code is a win, not a failure.

## Review checklist (user-facing changes)

1. New failure path: does the user see cause + next step from the catalog, with a stable error code underneath, and no internals leaked?
2. Any copy hardcoded in a component instead of the catalog?
3. Keyboard-only walk of the new flow: does it work? Labels, focus, contrast from tokens? Does the axe gate cover it?
4. Any flow that traps the user: no human path, asymmetric cancel, a dark pattern in the copy?
5. New feature: what evidence says someone wants it, and what adoption threshold decides keep-or-kill after launch?
6. Market-specific behaviour: driven by config and measured completion, not the home market's habits?
