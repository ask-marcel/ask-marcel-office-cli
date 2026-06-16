import type { CommandCategory, CommandMeta, PaginationStrategy } from './command-types.ts';

export type CommandManifestEntry = {
  readonly name: string;
  readonly summary: string;
  readonly category: CommandCategory;
  readonly graphMethod: CommandMeta['graphMethod'];
  readonly graphPathTemplate: string;
  readonly graphDocsUrl: string;
  readonly options: CommandMeta['options'];
  readonly positionalArguments?: CommandMeta['positionalArguments'];
  readonly example: string;
  readonly responseShape?: string;
  readonly bodyTemplate?: string;
  readonly pagination?: true;
  readonly paginationStrategy?: CommandMeta['paginationStrategy'];
  readonly scopesRequired?: CommandMeta['scopesRequired'];
  readonly needsElevatedToken?: CommandMeta['needsElevatedToken'];
  readonly producesBytes?: CommandMeta['producesBytes'];
  readonly producesMedia?: CommandMeta['producesMedia'];
  readonly mutates?: CommandMeta['mutates'];
  readonly stability?: CommandMeta['stability'];
};

// Audit Alex-session §5 follow-up: PAGINATION_HINT used to be one
// `nextLink`-shaped string applied to every paginated command, including the
// 5 deltaLink and 2 preferMaxPageSize commands where the cursor and the
// `--top` semantics differ. `paginationHintFor` returns the right variant
// per strategy so an LLM reading `--help` no longer sees self-contradictory
// pagination advice (the audit caught this on `list-calendar-view-delta`
// and `list-team-installed-apps`).
const PAGINATION_HINT_NEXT_LINK =
  'Paginated by Microsoft Graph. The CLI hoists `@odata.nextLink` out of `data` to the **top-level `nextLink`** field of the response envelope. Pass that URL to `next-page --url <link>` and repeat until the field is absent. Do NOT look for `data["@odata.nextLink"]` — the presenter strips it from `data` so the cursor is always at envelope level.';

const PAGINATION_HINT_NEXT_LINK_NO_SKIP = `${PAGINATION_HINT_NEXT_LINK} (Graph rejects \`$skip\` on this endpoint with \`invalidRequest: $skip is not supported on this API.\` — the CLI also omits \`--skip\` from this command's option set. Use \`--top\` + the \`nextLink\` cursor only.)`;

const PAGINATION_HINT_DELTA_LINK =
  'Delta-paginated. While paging, the CLI hoists `@odata.nextLink` to the top-level `nextLink` field — use it with `next-page --url <link>`. On the FINAL page Graph emits `@odata.deltaLink` instead; the CLI hoists that to the top-level **`deltaLink`** field. Stash that delta token and pass it back as the URL on the next invocation to resume the delta from where it left off (the response will contain only resources changed since then). Both cursors sit at envelope level — `data["@odata.nextLink"]` / `data["@odata.deltaLink"]` are stripped.';

const PAGINATION_HINT_PREFER_MAX_PAGE_SIZE =
  'Paginated, with a quirk: Graph rejects `$top` as a query parameter on this endpoint (`ErrorInvalidUrlQuery`). The CLI translates `--top <N>` into a `Prefer: odata.maxpagesize=<N>` request header internally — semantically equivalent, just routed differently. The cursor itself is the standard `@odata.nextLink` → top-level `nextLink` envelope field (and `@odata.deltaLink` → top-level `deltaLink` on the final page when this endpoint is also a delta); feed it back to `next-page --url <link>`. Other OData passthroughs (`$select` / `$filter` / `$orderby` / `$skip`) are silently ignored by Graph on this endpoint, so the CLI does not advertise them either.';

export const paginationHintFor = (strategy: PaginationStrategy | undefined): string => {
  if (strategy === 'nextLinkNoSkip') return PAGINATION_HINT_NEXT_LINK_NO_SKIP;
  if (strategy === 'deltaLink') return PAGINATION_HINT_DELTA_LINK;
  if (strategy === 'preferMaxPageSize') return PAGINATION_HINT_PREFER_MAX_PAGE_SIZE;
  // 'nextLink' (default when paginationStrategy is omitted on a paginated command)
  return PAGINATION_HINT_NEXT_LINK;
};

// Back-compat alias for any external readers of `PAGINATION_HINT`; new code
// should use `paginationHintFor(strategy)`.
export const PAGINATION_HINT = PAGINATION_HINT_NEXT_LINK;

export type CommandManifest = {
  readonly package: string;
  readonly version: string;
  readonly generatedAt: string;
  readonly commands: ReadonlyArray<CommandManifestEntry>;
};

// Audit Alex-session §4 follow-up: `auth` and `contacts` were declared here
// but no command ever set `category: 'auth' | 'contacts'` — they surfaced
// as bogus options in the unknown-category error from `help-json --category
// <bad>`, misleading LLMs into burning round-trips on empty categories.
// Removed both AND from the `CommandCategory` union in command-types.ts so
// the type system enforces "no command can ever claim a dead category".
const CATEGORY_LABELS: Readonly<Record<CommandCategory, string>> = {
  drive: 'OneDrive Files',
  excel: 'Excel (workbook files)',
  sharepoint: 'SharePoint Sites',
  tasks: 'Tasks (To Do + Planner)',
  mail: 'Mail',
  notes: 'Notes (OneNote)',
  user: 'User',
  calendar: 'Calendar',
  chats: 'Chats',
  teams: 'Teams',
  meta: 'Meta / Pagination',
  lifecycle: 'Lifecycle',
};

const CATEGORY_ORDER: ReadonlyArray<CommandCategory> = ['lifecycle', 'drive', 'excel', 'sharepoint', 'tasks', 'mail', 'notes', 'user', 'calendar', 'chats', 'teams', 'meta'];

const sortByName = (a: CommandManifestEntry, b: CommandManifestEntry): number => a.name.localeCompare(b.name);

const renderAliasSuffix = (aliases: CommandManifestEntry['options'][number]['aliases']): string => {
  if (!aliases || aliases.length === 0) return '';
  const names = aliases.map((a) => `\`--${a.name}\``).join(', ');
  return ` _(aliases: ${names})_`;
};

const renderRequiredParams = (entry: CommandManifestEntry): string => {
  const positionals = (entry.positionalArguments ?? []).map((p) => `\`<${p.name}>\``);
  const flags = entry.options.map((o) => `\`--${o.name}\``);
  const all = [...positionals, ...flags];
  if (all.length === 0) return '_(none)_';
  return all.join(', ');
};

const renderCategoryTable = (category: CommandCategory, entries: ReadonlyArray<CommandManifestEntry>): string => {
  const rows = entries.map((e) => `| \`${e.name}\` | ${e.summary} | ${renderRequiredParams(e)} | \`${e.graphMethod} ${e.graphPathTemplate}\` |`).join('\n');
  return `### ${CATEGORY_LABELS[category]}\n\n| Command | Description | Required params | Graph endpoint |\n|---------|-------------|-----------------|----------------|\n${rows}`;
};

export const renderReadmeTables = (manifest: CommandManifest): string => {
  const sections: string[] = [];
  for (const category of CATEGORY_ORDER) {
    const entries = manifest.commands.filter((c) => c.category === category).toSorted(sortByName);
    if (entries.length === 0) continue;
    sections.push(renderCategoryTable(category, entries));
  }
  return sections.join('\n\n');
};

export const renderCommandMarkdown = (entry: CommandManifestEntry): string => {
  const lines = [
    `# \`${entry.name}\``,
    '',
    entry.summary,
    '',
    `- **Category:** ${CATEGORY_LABELS[entry.category]}`,
    `- **Graph endpoint:** \`${entry.graphMethod} ${entry.graphPathTemplate}\``,
    `- **Microsoft Learn:** ${entry.graphDocsUrl}`,
  ];
  if (entry.responseShape) lines.push(`- **Response:** ${entry.responseShape}`);
  if (entry.pagination) lines.push(`- **Pagination:** ${paginationHintFor(entry.paginationStrategy)}`);
  if (entry.scopesRequired && entry.scopesRequired.length > 0) {
    const tagged = entry.scopesRequired.map((s) => `\`${s}\``).join(', ');
    lines.push(`- **Scopes required:** ${tagged} — run \`ask-marcel scopes-check\` to verify before invoking.`);
  }
  if (entry.needsElevatedToken) {
    lines.push(
      '- **Needs elevated token:** This command requires the M365ChatClient token captured at login (ODSP allow-list). If the silent SSO capture failed, the command will time out — run `ask-marcel login` to retry the capture.'
    );
  }
  if (entry.stability === 'experimental') {
    lines.push(
      '- **Stability:** `experimental` — rides a Microsoft-internal substrate that is not in the public Graph API and can break without notice on a Teams web-client update. Prefer a `stable` sibling when one exists.'
    );
  }
  if (entry.positionalArguments !== undefined && entry.positionalArguments.length > 0) {
    lines.push('', '## Positional arguments', '');
    lines.push('| Argument | Required | Description |', '|----------|----------|-------------|');
    for (const a of entry.positionalArguments) lines.push(`| \`<${a.name}>\` | ${a.required ? 'yes' : 'no'} | ${a.description} |`);
  }
  if (entry.options.length > 0) {
    lines.push('', '## Options', '');
    lines.push('| Flag | Description |', '|------|-------------|');
    for (const o of entry.options) lines.push(`| \`--${o.name}\` | ${o.description}${renderAliasSuffix(o.aliases)} |`);
  }
  if (entry.bodyTemplate) lines.push('', '## Request body', '', '```json', entry.bodyTemplate, '```');
  lines.push('', '## Example', '', '```bash', entry.example, '```');
  return lines.join('\n');
};

export { CATEGORY_LABELS, CATEGORY_ORDER };
