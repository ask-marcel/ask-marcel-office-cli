import type { CommandCategory, CommandMeta } from './command-types.ts';

export type CommandManifestEntry = {
  readonly name: string;
  readonly summary: string;
  readonly category: CommandCategory;
  readonly graphMethod: CommandMeta['graphMethod'];
  readonly graphPathTemplate: string;
  readonly graphDocsUrl: string;
  readonly options: CommandMeta['options'];
  readonly example: string;
  readonly responseShape?: string;
  readonly bodyTemplate?: string;
  readonly pagination?: true;
};

export const PAGINATION_HINT =
  'Paginated by Microsoft Graph. The response may include `@odata.nextLink`; pass that URL to `next-page --url <link>` and repeat until the field is absent.';

export type CommandManifest = {
  readonly package: string;
  readonly version: string;
  readonly generatedAt: string;
  readonly commands: ReadonlyArray<CommandManifestEntry>;
};

const CATEGORY_LABELS: Readonly<Record<CommandCategory, string>> = {
  auth: 'Authentication',
  drive: 'OneDrive Files',
  excel: 'Excel (workbook files)',
  sharepoint: 'SharePoint Sites',
  tasks: 'Tasks (To Do + Planner)',
  mail: 'Mail',
  notes: 'Notes (OneNote)',
  user: 'User',
  calendar: 'Calendar',
  contacts: 'Contacts',
  chats: 'Chats',
  teams: 'Teams',
  meta: 'Meta / Pagination',
  lifecycle: 'Lifecycle',
};

const CATEGORY_ORDER: ReadonlyArray<CommandCategory> = [
  'lifecycle',
  'drive',
  'excel',
  'sharepoint',
  'tasks',
  'mail',
  'notes',
  'user',
  'calendar',
  'contacts',
  'chats',
  'teams',
  'meta',
];

const sortByName = (a: CommandManifestEntry, b: CommandManifestEntry): number => a.name.localeCompare(b.name);

const renderRequiredParams = (entry: CommandManifestEntry): string => {
  if (entry.options.length === 0) return '_(none)_';
  return entry.options.map((o) => `\`--${o.name}\``).join(', ');
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
  if (entry.pagination) lines.push(`- **Pagination:** ${PAGINATION_HINT}`);
  if (entry.options.length > 0) {
    lines.push('', '## Options', '');
    lines.push('| Flag | Description |', '|------|-------------|');
    for (const o of entry.options) lines.push(`| \`--${o.name}\` | ${o.description} |`);
  }
  if (entry.bodyTemplate) lines.push('', '## Request body', '', '```json', entry.bodyTemplate, '```');
  lines.push('', '## Example', '', '```bash', entry.example, '```');
  return lines.join('\n');
};

export { CATEGORY_LABELS, CATEGORY_ORDER };
