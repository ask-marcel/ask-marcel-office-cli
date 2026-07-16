/**
 * Wire-safe builders for Graph search clauses. graph-client concatenates
 * command paths verbatim (`https://graph.microsoft.com/v1.0${path}`) with no
 * percent-encoding pass, so a raw `&`, `#`, `%`, or `+` inside a user query
 * corrupts the query string on the wire: a live ` & ` truncated the KQL
 * mid-phrase (`$search="subject:"Contoso A2` reached Graph). Values are
 * percent-encoded here, at the only layer that knows which part of the path
 * is data.
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

/**
 * Value for the OData `search(q='…')` function form: single quotes double
 * per OData string-literal escaping, then percent-encode.
 */
const odataSearchQ = (query: string): string => encodeURIComponent(query.replaceAll("'", "''"));

export { kqlSearchClause, odataSearchQ };
