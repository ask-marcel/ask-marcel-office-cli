# Isolation (one user's data must never reach another)

The worst failure in any system that holds data for more than one person is the quiet one: customer A sees customer B's records. It rarely announces itself, and by the time you notice, the exposure is history. Treat the boundary between one tenant, customer, or user and the next as a first-class part of the architecture, not a WHERE clause you remember to add. This is the doctrine behind hard rule 28: tenant isolation is fail-closed and token-derived, and every endpoint proves it with a cross-tenant test.

Applies to any code path that serves more than one user or organisation: multi-tenant SaaS, a user-scoped API, even a CLI acting on behalf of different accounts.

## 1. Derive the owner from one trusted source

The tenant or user a request acts for comes from a **verified token claim**, never from a URL segment, header, or body field the caller controls.

```ts
// BAD: tenant from the path; any caller can name another tenant
app.get('/orgs/:orgId/invoices', (c) => c.json(listInvoices(c.req.param('orgId'))));

// GOOD: tenant from the verified JWT; the URL cannot override it
app.get('/invoices', requireAuth, (c) => {
  const orgId = c.get('claims').org_id; // signed by the IdP, not caller-supplied
  return c.json(listInvoices(orgId));
});
```

Java (Quarkus): read the tenant from `SecurityIdentity` (the validated token), never from `@HeaderParam("X-Org-Id")` or a path param the caller chose.

Brand it: `OrgId` is a branded type whose only production factory reads the verified claim (rule 12). A function signature taking `OrgId` then carries the proof that the value came through the checkpoint.

## 2. Defend in depth

Assume each layer has a bug; enforce the boundary in more than one place. Application code filters by owner, and the data store re-checks with row-level security, inside the same transaction:

```ts
// app filter AND a tx-local session var the RLS policy reads
const rows = await db.transaction(async (tx) => {
  await tx.execute(sql`SELECT set_config('app.current_org', ${orgId}, true)`); // true: tx-local
  return tx.select().from(invoices).where(eq(invoices.orgId, orgId));
});
```

```sql
-- migration: the second layer; a forgotten WHERE clause is caught here
CREATE POLICY tenant_isolation ON invoices
  USING (org_id = current_setting('app.current_org')::uuid);
```

Java: run the `set_config` native query inside the same `@Transactional` block as the read. The policy is identical; the layer that forgets is different, the layer that catches is the same.

## 3. Fail closed

Missing owner context returns **nothing**: not everything, and not a 500 that hints at internals. A bug should surface as absent data, never as someone else's data.

```ts
// BAD: undefined falls through to an unscoped query
const invoicesFor = (orgId?: string) => (orgId ? scoped(orgId) : db.select().from(invoices));

// GOOD: no owner means no data
export const invoicesFor = (orgId: OrgId | undefined): Promise<Invoice[]> => {
  if (!orgId) return Promise.resolve([]);
  return db.select().from(invoices).where(eq(invoices.orgId, orgId));
};
```

The branded `OrgId` makes the closed default structural: an unscoped read does not typecheck.

## 4. Shrink the blast radius

Give the runtime the narrowest access it can work with, so the worst case of one injection or one leaked credential is one owner's data, not the estate.

```sql
-- BAD: the app connects as a superuser that can read, drop, and bypass RLS
ALTER ROLE app_runtime WITH SUPERUSER BYPASSRLS;

-- GOOD: least privilege; crucially the runtime role cannot bypass RLS
CREATE ROLE app_runtime NOSUPERUSER NOBYPASSRLS;
GRANT SELECT, INSERT, UPDATE ON invoices, receipts TO app_runtime; -- no DELETE, no DDL
```

- Migrations run as a separate migrator role, only in CI, never at runtime.
- Subject erasure (`references/privacy.md`) runs as its own narrowly granted job; the runtime role anonymizes at most.
- The connection string the app opens is the constrained role's, read through the validated config module (`references/security.md`).

## 5. Prove isolation per endpoint (the test that ships with every route)

Isolation you have not tested is isolation you do not have. Every endpoint that serves owner-scoped data ships a test where owner A's credentials against owner B's resource return **404 not_found** (not 403: a 403 confirms the resource exists, which is itself a leak).

```ts
test('cross-tenant read is not_found', async () => {
  const res = await app.request(`/invoices/${orgBInvoice}`, authAs(orgA));
  expect(res.status).toBe(404); // absence, not refusal: existence is not disclosed
});
```

And test the seam a real attacker would walk (SKILL.md red flags; `references/testing.md`, Bypass tests): drive the request through the real edge with a forged trust header and assert it is inert.

```ts
test('a forged org header is ignored; scope derives from the token', async () => {
  const res = await app.request(`/invoices/${orgAInvoice}`, {
    headers: { authorization: bearer(orgB), 'x-org-id': orgA }, // attacker forges the header
  });
  expect(res.status).toBe(404);
});
```

Java: the same pair with REST Assured (`given().auth().oauth2(tokenOrgA).get(...).then().statusCode(404)`).

## 6. Identifiers: unguessable, and never the authorization

- A sequential id in a URL is an invitation: one loop enumerates every record, and the top id leaks your volume. Use **UUIDv7**: the random bits stop guessing, the time-ordered prefix keeps the index local (`references/reliability.md`, Keyset pagination relies on the same index locality). If creation time is itself sensitive, use v4 and accept the index cost.
- Keep internal keys internal. Where the flow allows, expose no id at all: scope routes by the verified identity (`GET /me/invoices`) so most flows never carry a raw id.
- An unguessable id is defense in depth **on top of** authorization, never a substitute. The cross-tenant 404 test above still holds with UUIDs.

```ts
export const invoices = pgTable('invoices', {
  id: uuid('id').primaryKey().$defaultFn(() => uuidv7()),
});
app.get('/me/invoices', requireAuth, (c) => c.json(listInvoices(c.get('claims').org_id)));
```

## Fakes for isolation

The in-memory fakes in `src/test-helpers/` must model the boundary, or use-case tests cannot exercise it: a fake repository stores rows keyed by owner and its readers require the `OrgId`. A fake whose `list()` ignores the owner will happily pass a use-case that leaks.

## Executable tripwire

`assets/check-isolation-tests.sh` refuses a newly staged route/resource file with no nearby test mentioning a 404 (globs configurable at the top of the script; `*public*`/`*health*` paths exempt by convention). It is deliberately the weakest of the four guards: it proves a cross-tenant test exists near the route, not that it asserts the right thing. The per-endpoint test above remains the contract; the wire just refuses the common failure of landing a route with no isolation test at all.

## Review checklist (changes in a multi-user code path)

1. Where does the owner id come from? A verified claim (or a branded `OrgId` minted from one), or something the caller controls?
2. Is there a second enforcement layer (RLS or equivalent) behind the application filter?
3. What happens when the owner context is missing? Empty result, or everything?
4. Does the runtime role's grant list match what the code path needs, nothing more?
5. Does the endpoint ship its cross-tenant 404 test and, at the edge, the forged-header test?
6. Any new id in a URL: unguessable, and not doing the job of authorization?
