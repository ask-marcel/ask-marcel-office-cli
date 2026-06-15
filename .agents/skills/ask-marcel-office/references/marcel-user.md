# User commands

## get-current-user
Return the signed-in user's Microsoft Graph profile. The CLI ships a slim default `--select=id,displayName,mail,userPrincipalName,jobTitle,officeLocation,mobilePhone` covering the common identity fields. Pass `--select id,displayName,givenName,surname,preferredLanguage,...` to widen, or `--select '*
Optional: --select --expand
Example: ask-marcel get-current-user
Graph: GET /me

## get-group
Return metadata for a single Azure AD / Microsoft 365 group. Use `--select` to slim large group payloads (the full group resource includes 30+ fields).
Required: --group-id
Optional: --select --expand
Example: ask-marcel get-group --group-id 'a1b2c3d4-...' --select 'id,displayName,mail'
Graph: GET /groups/{group-id}

## get-my-manager
Return the signed-in user's manager (a single `user` resource). When no manager is set in the directory, Graph returns 404 `Request_ResourceNotFound`; this command maps that one specific 404 to `{ ok: true, data: { manager: null, note: '...' } }` so an LLM can distinguish 'no manager' from a permiss
Optional: --select --expand
Example: ask-marcel get-my-manager --select 'id,displayName,mail'
Graph: GET /me/manager

## get-my-profile-photo
Download the signed-in user's profile photo (largest available size), inlined. The CLI follows the Graph 302 → CDN redirect internally so the LLM never has to fetch an external URL.
Example: ask-marcel get-my-profile-photo
Graph: GET /me/photo/$value

## get-organization
Return the tenant's organization metadata — display name, country, verified domains, business phones, technical / security notification contacts, assigned Microsoft 365 SKUs / licensing. Graph wraps the single organization resource under `value[]` (audit v1.0.0 §D7 — even though only one tenant exis
Optional: --select --expand
Example: ask-marcel get-organization --select 'id,displayName,verifiedDomains'
Graph: GET /organization

## get-user-manager
Return a specific user's manager (a single `user` resource). When the user has no manager set in the directory, Graph returns 404 `Request_ResourceNotFound`; this command maps that one specific 404 to `{ ok: true, data: { manager: null, note: '...' } }` (same shape as `get-my-manager`) so an LLM can
Required: --user-id
Optional: --select --expand
Example: ask-marcel get-user-manager --user-id 'alice@contoso.com' --select 'id,displayName,mail'
Graph: GET /users/{user-id}/manager

## list-group-members
List members of an Azure AD / Microsoft 365 group. Returns users, groups, and other directoryObjects depending on the group's membership.
Required: --group-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-group-members --group-id 'a1b2c3d4-...'
Graph: GET /groups/{group-id}/members

## list-group-owners
List the owners of an Azure AD / Microsoft 365 group.
Required: --group-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-group-owners --group-id 'a1b2c3d4-...'
Graph: GET /groups/{group-id}/owners

## list-groups
List Microsoft 365 groups, security groups, and distribution groups in the tenant directory. Use `--top` and `next-page` to paginate over very large directories.
Optional: --top --select --filter --orderby --expand
Example: ask-marcel list-groups
Graph: GET /groups

## list-my-direct-reports
List the signed-in user's direct reports (employees who report to them in the directory). When `--orderby` is supplied the CLI auto-injects the `ConsistencyLevel: eventual` header Graph requires on directory endpoints — otherwise Graph rejects the sort with `Request_UnsupportedQuery`.
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-my-direct-reports
Graph: GET /me/directReports

## list-my-memberships
List the groups, directory roles, and administrative units the signed-in user is a member of. Each entry's `@odata.type` distinguishes #microsoft.graph.group from #microsoft.graph.directoryRole, etc.
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-my-memberships
Graph: GET /me/memberOf

## list-my-transitive-memberships
List all groups, directory roles, and administrative units the signed-in user is a member of *transitively* — including memberships inherited via nested groups. Sibling to `list-my-memberships` (`/me/memberOf`) which only returns direct memberships.
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-my-transitive-memberships
Graph: GET /me/transitiveMemberOf

## list-relevant-people
List people relevant to the signed-in user — colleagues they email and meet with most. Microsoft's relevance ranking, not the full directory. Returns `displayName`, `emailAddresses`, `jobTitle`, `companyName`, etc.
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-relevant-people
Graph: GET /me/people

## list-sensitivity-labels
List the Microsoft Information Protection sensitivity labels available to the signed-in user — the labels Outlook / Word / SharePoint surfaces in the "Sensitivity" picker (e.g. Public / Internal / Confidential / Highly Confidential). Each label has `id`, `displayName`, `priority`, `isAppliable`, `to
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-sensitivity-labels
Graph: GET /me/informationProtection/sensitivityLabels

## list-user-direct-reports
List a specific user's direct reports.
Required: --user-id
Optional: --top --skip --select --filter --orderby --expand
Example: ask-marcel list-user-direct-reports --user-id 'alice@contoso.com'
Graph: GET /users/{user-id}/directReports
