// Graph returns some `@odata.nextLink` / `@odata.deltaLink` cursors with the
// `$` of its OData system query options percent-encoded as `%24` — e.g.
// `/me/people?%24top=2&%24skip=2` (the `$skip`/`$top`-paged endpoints do this;
// the `$skiptoken`-paged ones return a literal `$skiptoken`). That `%24` form
// is technically valid but reads as malformed to an LLM consumer and invites a
// double-encoding bug: an agent that "helpfully" re-encodes the URL turns
// `%24top` into `%2524top`, which Graph then treats as an unknown parameter and
// silently ignores — pagination loops forever on page 1.
//
// Decode ONLY `%24` → `$` so the surfaced cursor matches the canonical
// `$top` / `$skip` / `$skiptoken` form the CLI documents and every `next-page`
// example uses. Other percent-escapes (`%2B`, `%2F`, `%3D`, `%20`, ...) are
// left untouched: those encode literal bytes inside opaque skiptoken VALUES,
// where decoding would change what the server receives (`%2B` → `+` decodes to
// a space on the far side). `%24` is safe because `$` is a non-delimiter inside
// a query value and decodes identically whether sent raw or as `%24`.
const canonicalizeGraphCursor = (link: string): string => link.replace(/%24/gi, '$');

export { canonicalizeGraphCursor };
