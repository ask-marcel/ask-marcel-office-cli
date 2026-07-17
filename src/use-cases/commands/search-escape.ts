/**
 * Wire-safe builder for Graph `$search=` KQL clauses. graph-client
 * concatenates command paths verbatim (`https://graph.microsoft.com/v1.0${path}`)
 * with no percent-encoding pass, so a raw `&`, `#`, `%`, or `+` inside a user
 * query corrupts the query string on the wire: a live ` & ` truncated the KQL
 * mid-phrase (`$search="subject:"Contoso A2` reached Graph). Values are
 * percent-encoded here, at the only layer that knows which part of the path
 * is data.
 *
 * The sibling concern — a value inside a single-quoted OData literal
 * (`$filter=x eq '…'`, `contains(…)`, `search(q='…')`) — is
 * `odataStringLiteral` in odata-query.ts, which owns OData query construction.
 */

/**
 * Build a `$search=` KQL clause: escape embedded double quotes as `\"`
 * (Graph's documented KQL escaping, so `subject:"multi word"` and whole
 * `"phrase"` queries become phrase matches instead of a BadRequest), wrap the
 * expression in the double quotes Graph requires, then percent-encode the
 * value. Backslashes are NOT pre-escaped: raw quotes are the documented
 * input contract and backslash has no other KQL meaning here.
 */
const kqlSearchClause = (query: string): string => {
  const quoted = `"${query.replaceAll('"', '\\"')}"`;
  return `$search=${encodeURIComponent(quoted)}`;
};

export { kqlSearchClause };
