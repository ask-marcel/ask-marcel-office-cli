# Privacy (private by default)

Personal data is a liability you choose to hold, not an asset. The cheapest data to protect, and the only data that can never leak, is the data you chose not to collect. This reference binds whenever code touches personal data: names, emails, phone numbers, addresses, identifiers tied to a person, free text a user typed, and anything sensitive (health, finances, biometrics, minors). It is the doctrine behind hard rules 27 (no personal data in logs, URLs, or query strings) and 34 (production data never leaves production).

## Mental model

1. **Minimize.** Collect only what a stated purpose requires. If you cannot say the purpose in one sentence, do not collect the field.
2. **Contain.** Personal data stays out of the places it leaks by accident: logs, URLs, query strings, analytics events, third-party payloads, lower environments.
3. **Serve rights.** See, correct, export, delete, withdraw consent: routine authenticated endpoints from day one, never a manual database ticket.

## Know the law that follows the user

Privacy regimes attach to the user, not to your office: serving people in a region generally puts you under its regime (GDPR, PIPL, and their siblings) regardless of where the company or servers sit. Design to the strictest regime you serve.

```ts
// BAD: one hardcoded assumption about the home market
const needsConsent = (): boolean => false;

// GOOD: resolve the governing regimes from residency + data locations, apply the strictest
export const policyFor = (user: User): PrivacyPolicy => strictest(regimesFor(user.residency, user.dataLocations));
```

Do not hand-roll legal analysis in code reviews; the code-level duty is that the consent and rights machinery exists and is driven by config, not baked-in assumptions.

## Minimize and justify collection

```ts
// BAD: hoover up sensitive extras "in case they are useful later"
const Signup = z.object({ email: z.string().email(), ssn: z.string(), religion: z.string(), dob: z.string() });

// GOOD: only what the purpose needs; a sensitive field would require its own explicit consent flow
const Signup = z.object({
  email: z.string().email(), // creates the account
  displayName: z.string().min(1), // renders the profile
});
```

Sensitive categories (health, biometrics, finances, anything concerning minors) need explicit, specific, purpose-bound consent, never a pre-ticked box or a bundled agreement.

## Personal data never in logs, URLs, or query strings (hard rule 27)

Query strings are written to access logs, cached by proxies, stored in browser history, forwarded in referrer headers, and handed to analytics. Log lines are shipped, indexed, and retained. So:

- Personal data and any free text the user typed travel in a **POST body**. You cannot know at runtime whether a search term is sensitive, so user wording goes in the body by default. Searching by someone's name is a POST, not a GET.
- Query strings carry only structural, public values: a page cursor, a sort key, a locale. Never what someone wrote, never a token.
- **Opaque internal identifiers are loggable** (a UUID `userId` correlates without exposing). **Natural identifiers are not**: email, phone, name, token, national id.
- Redaction happens once, at the logger adapter, not at every call site. Extend the Winston `redactFormat` key set (`references/security.md`, Logging discipline) with the natural-identifier keys of the domain: `email`, `phone`, `name`, `address`, `ssn` alongside `password`, `token`, `authorization`.

```ts
// BAD: PII in the query string and a raw-body log line
await fetch(`/search?email=${user.email}`);
logger.info('request', { body: req.body });

// GOOD: wording in the body; the log carries an opaque id only
await fetch('/search', { method: 'POST', body: JSON.stringify({ email: user.email }) });
logger.info('search handled', { userId: user.id });
```

React/Next.js: never build a route with personal data in it. `navigate('/search?email=...')` lands in history, referrers, and analytics; pass it in state or POST it, and keep the address bar structural.

Java (Quarkus): same rule. `@QueryParam("email")` on a search endpoint is a violation; take a `@Valid SearchRequest` body, and log a pseudonymous reference through the redacting logger (`references/java-quarkus.md`, Logging).

## User rights are routine endpoints

Ship the five rights as ordinary authenticated, audited operations from the first schema version:

```ts
app.get('/me/data', requireAuth, (c) => exportSubject(c)); // see + export
app.patch('/me', requireAuth, zValidator('json', Correction), correctSubject); // correct
app.delete('/me', requireAuth, (c) => eraseSubject(c)); // delete: REAL erasure or anonymization
app.post('/me/consent/withdraw', requireAuth, withdrawConsent); // withdraw
```

**Delete means delete.** Subject erasure hard-deletes or anonymizes the personal fields. This is the deliberate exception to the soft-delete default of hard rule 30: business records may keep their soft-deleted rows, but the personal fields inside them are truly erased or anonymized. A scheduled retention sweep is where the two rules reconcile (`references/reliability.md`, Data lifecycle). The erasure job runs under its own narrowly granted role, not the general runtime role (`references/isolation.md`, Blast radius).

## Map and classify

Every personal field carries a recorded class and purpose, so the data map is generated, not remembered:

```ts
// domain/schemas/customer-data-map.ts: read by tooling to build the inventory
export const customerDataMap = {
  email:   { class: 'pii',       purpose: 'account',   crossesBorder: false },
  taxId:   { class: 'sensitive', purpose: 'invoicing', crossesBorder: false },
  country: { class: 'pii',       purpose: 'tax-rules', crossesBorder: true },
} as const;
```

What good looks like: you can produce a current map of every category of personal data on request, including where it flows, who it is shared with, and whether it crosses a border. A field with no owner, purpose, or class is a red flag.

## Production data never leaves production (hard rule 34)

Test and dev environments have weaker controls and wider access, which makes them the cheapest place to lose real data. Never restore a production dump into a lower environment, a laptop, or a test.

```ts
// BAD: real customers seeded into dev
const rows = await pgRestore('prod_backup.sql');

// GOOD: deterministic synthetic fixtures with the right shape and volume, zero real subjects
const rows = Array.from({ length: 500 }, (_, i) => ({
  email: `user${i}@example.test`,
  taxId: fakeTaxId(i),
  country: pickCountry(i),
}));
```

Seed builders live in `src/test-helpers/` next to the port fakes; make them deterministic (seeded index, not random) so tests are reproducible. When a bug only reproduces on production data, debug against production with read access and observability (`references/observability.md`), not by copying the data out.

## Assess before risky processing

Before an automated decision, sensitive-data processing, or a cross-border transfer, record a short impact assessment (DPIA) and gate the operation on it:

```ts
export const autoDecide = async (deps: Deps, a: Application): Promise<Result<Decision, 'dpia_missing'>> => {
  const dpia = await deps.assessments.find('auto-credit-decision');
  if (!dpia.ok || dpia.value.status !== 'approved') return err('dpia_missing');
  return ok(score(a) < 0.4 ? 'reject' : 'review');
};
```

The assessment is a committed document (like an ADR, `references/governance.md`), revisited when the purpose, a contract, a regulation, or a breach changes.

## Third parties and analytics

- Never send personal data to an analytics or telemetry vendor as event properties. The event carries the signal (`waitlist_signup`), your own endpoint carries the address.
- Never put personal data in URL parameters sent to any third party.
- Sharing data with a processor is part of the data map: record who, what, and under which agreement.

## Executable tripwire

The mechanical slice of rule 27 ships as a staged-diff gate: `assets/check-pii-channels.sh` blocks a natural identifier in a query string (written literally or built via `new URLSearchParams`), a logger message interpolation, or a Java `@QueryParam`, on the lines a commit adds (`--all` audits the whole tree). It is a tripwire, not a proof: this checklist remains the review duty; the script just refuses the common concrete leaks. Wire it as a pre-commit pre-flight or CI step wherever the repo holds personal data (`references/workflow.md`, Discipline tripwires).

## Review checklist (changes touching personal data)

1. New field collected: what is the one-sentence purpose? Is it in the data map with a class?
2. Does any personal value reach a log line, a URL, a query string, or a third-party event? (rule 27)
3. Do the rights endpoints still cover the new field (export includes it, erasure clears it)?
4. Do fixtures stay synthetic? No prod dump, no real subject in a test. (rule 34)
5. Sensitive category or automated decision: is there an approved assessment on record?
6. Retention: when does this data die, and which sweep enforces it?
