import { describe, expect, it } from 'bun:test';
import type { CommandManifest, CommandManifestEntry } from './docs-render.ts';
import { CATEGORY_LABELS, CATEGORY_ORDER, renderCommandMarkdown, renderReadmeTables } from './docs-render.ts';

const calendarEvent: CommandManifestEntry = {
  name: 'get-calendar-event',
  summary: 'Fetch a single calendar event by ID.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/events/{event-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-get',
  options: [{ name: 'event-id', key: 'eventId', required: true, description: 'The Graph event ID.' }],
  example: "ask-marcel-office get-calendar-event --event-id 'AAMk...'",
  responseShape: 'single event',
};

const listDrives: CommandManifestEntry = {
  name: 'list-drives',
  summary: 'List the OneDrive drives.',
  category: 'drive',
  graphMethod: 'GET',
  graphPathTemplate: '/me/drives',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/drive-list',
  options: [],
  example: 'ask-marcel-office list-drives',
};

const sampleManifest: CommandManifest = {
  package: 'ask-marcel-office-cli',
  version: '0.1.2',
  generatedAt: '2026-04-30T00:00:00Z',
  commands: [calendarEvent, listDrives],
};

describe('renderReadmeTables', () => {
  it('groups commands by category and lists each in its own table', () => {
    const md = renderReadmeTables(sampleManifest);
    expect(md).toContain('### OneDrive Files');
    expect(md).toContain('### Calendar');
    expect(md).toContain('| `list-drives` | List the OneDrive drives. | _(none)_ | `GET /me/drives` |');
    expect(md).toContain('| `get-calendar-event` | Fetch a single calendar event by ID. | `--event-id` | `GET /me/events/{event-id}` |');
  });

  it('orders OneDrive Files before Calendar (canonical category order)', () => {
    const md = renderReadmeTables(sampleManifest);
    expect(md.indexOf('### OneDrive Files')).toBeLessThan(md.indexOf('### Calendar'));
  });

  it('skips categories with no commands', () => {
    const md = renderReadmeTables(sampleManifest);
    expect(md).not.toContain('### SharePoint Sites');
  });

  it('sorts commands alphabetically within a category', () => {
    const zebra: CommandManifestEntry = { ...listDrives, name: 'list-zebra-drives' };
    const apple: CommandManifestEntry = { ...listDrives, name: 'list-apple-drives' };
    const manifest: CommandManifest = { ...sampleManifest, commands: [zebra, apple] };
    const md = renderReadmeTables(manifest);
    expect(md.indexOf('list-apple-drives')).toBeLessThan(md.indexOf('list-zebra-drives'));
  });

  it("renders positional arguments in the readme table's required-params column with `<name>` markers", () => {
    const positional: CommandManifestEntry = {
      ...listDrives,
      name: 'docs',
      category: 'lifecycle',
      positionalArguments: [{ name: 'command', required: true, description: 'Command name.' }],
    };
    const manifest: CommandManifest = { ...sampleManifest, commands: [positional] };
    const md = renderReadmeTables(manifest);
    expect(md).toContain('`<command>`');
  });
});

describe('renderCommandMarkdown', () => {
  it('renders a command with options into a Markdown brief', () => {
    const md = renderCommandMarkdown(calendarEvent);
    expect(md).toContain('# `get-calendar-event`');
    expect(md).toContain('Fetch a single calendar event by ID.');
    expect(md).toContain('**Graph endpoint:** `GET /me/events/{event-id}`');
    expect(md).toContain('**Microsoft Learn:** https://learn.microsoft.com/en-us/graph/api/event-get');
    expect(md).toContain('**Response:** single event');
    expect(md).toContain('## Options');
    expect(md).toContain('| `--event-id` | The Graph event ID. |');
    expect(md).toContain('## Example');
    expect(md).toContain("ask-marcel-office get-calendar-event --event-id 'AAMk...'");
  });

  it('omits the Options section when the command has no options', () => {
    const md = renderCommandMarkdown(listDrives);
    expect(md).not.toContain('## Options');
    expect(md).toContain('## Example');
  });

  it('omits the Response line when responseShape is not set', () => {
    const md = renderCommandMarkdown(listDrives);
    expect(md).not.toContain('**Response:**');
  });

  it('renders a Pagination line that names the next-page command when the entry is paginated', () => {
    const paginated: CommandManifestEntry = { ...listDrives, pagination: true };
    const md = renderCommandMarkdown(paginated);
    expect(md).toContain('**Pagination:**');
    expect(md).toContain('@odata.nextLink');
    expect(md).toContain('next-page --url');
  });

  it('omits the Pagination line when the entry is not paginated', () => {
    const md = renderCommandMarkdown(listDrives);
    expect(md).not.toContain('**Pagination:**');
  });

  // PAGINATION_HINT used to be one string
  // for every paginated command, including the 5 deltaLink and 2
  // preferMaxPageSize ones where the cursor + `--top` semantics differ.
  // `paginationHintFor(strategy)` returns the matching variant.
  it('renders the nextLinkNoSkip pagination hint with the explicit $skip-rejection clause when paginationStrategy is set accordingly', () => {
    const paginated: CommandManifestEntry = { ...listDrives, pagination: true, paginationStrategy: 'nextLinkNoSkip' };
    const md = renderCommandMarkdown(paginated);
    expect(md).toContain('Graph rejects `$skip` on this endpoint');
    expect(md).toContain('--top');
  });

  it('renders the deltaLink pagination hint pointing at `@odata.deltaLink` (NOT `nextLink`) for the final-page cursor', () => {
    const paginated: CommandManifestEntry = { ...listDrives, pagination: true, paginationStrategy: 'deltaLink' };
    const md = renderCommandMarkdown(paginated);
    expect(md).toContain('@odata.deltaLink');
    expect(md).toContain('deltaLink');
    expect(md).toContain('Delta-paginated');
  });

  it('renders the preferMaxPageSize pagination hint explaining the `--top` → `Prefer: odata.maxpagesize` header translation (Graph rejects $top as a query param on these endpoints)', () => {
    const paginated: CommandManifestEntry = { ...listDrives, pagination: true, paginationStrategy: 'preferMaxPageSize' };
    const md = renderCommandMarkdown(paginated);
    expect(md).toContain('Prefer: odata.maxpagesize');
    expect(md).toContain('rejects `$top` as a query parameter');
  });

  it('renders options unchanged (no suffix) so the format is unambiguous', () => {
    const md = renderCommandMarkdown(calendarEvent);
    expect(md).toContain('| `--event-id` | The Graph event ID. |');
    expect(md).not.toContain('aliases:');
  });

  it('renders a Scopes required line when scopesRequired is set', () => {
    const withScopes: CommandManifestEntry = { ...calendarEvent, scopesRequired: ['Chat.ReadBasic', 'User.Read'] };
    const md = renderCommandMarkdown(withScopes);
    expect(md).toContain('**Scopes required:**');
    expect(md).toContain('`Chat.ReadBasic`');
    expect(md).toContain('`User.Read`');
    expect(md).toContain('scopes-check');
  });

  it('renders an elevated-token warning when needsElevatedToken is true', () => {
    const elevated: CommandManifestEntry = { ...calendarEvent, needsElevatedToken: true };
    const md = renderCommandMarkdown(elevated);
    expect(md).toContain('**Needs elevated token:**');
    expect(md).toContain('M365ChatClient');
    expect(md).toContain('ask-marcel-office login');
  });

  it('renders a Positional arguments section when positionalArguments is set', () => {
    const positional: CommandManifestEntry = {
      ...listDrives,
      positionalArguments: [{ name: 'command', required: true, description: 'Name of the command to show docs for.' }],
    };
    const md = renderCommandMarkdown(positional);
    expect(md).toContain('## Positional arguments');
    expect(md).toContain('| `<command>` | yes | Name of the command to show docs for. |');
  });

  it('renders a Stability line when the entry is flagged experimental — surfaces the "may break without notice" warning structurally', () => {
    const experimental: CommandManifestEntry = { ...listDrives, stability: 'experimental' };
    const md = renderCommandMarkdown(experimental);
    expect(md).toContain('**Stability:** `experimental`');
    expect(md).toContain('Microsoft-internal substrate');
    expect(md).toContain('Prefer a `stable` sibling');
  });

  it('omits the Stability line when the entry is stable (default) — keeps stable commands quiet so the field carries signal only when it matters', () => {
    const md = renderCommandMarkdown(listDrives);
    expect(md).not.toContain('**Stability:**');
  });

  it('renders a Request body section (fenced json) when the entry has a bodyTemplate', () => {
    const withBody: CommandManifestEntry = { ...calendarEvent, bodyTemplate: '{ "subject": "<text>" }' };
    const md = renderCommandMarkdown(withBody);
    expect(md).toContain('## Request body');
    expect(md).toContain('```json');
    expect(md).toContain('{ "subject": "<text>" }');
  });

  it('omits the Request body section when the entry has no bodyTemplate', () => {
    expect(renderCommandMarkdown(listDrives)).not.toContain('## Request body');
  });

  it('omits the Scopes line for an EMPTY scopesRequired array (not just undefined) — guards the `.length > 0` check', () => {
    const md = renderCommandMarkdown({ ...calendarEvent, scopesRequired: [] });
    expect(md).not.toContain('**Scopes required:**');
  });

  it('omits the Positional arguments section for an EMPTY positionalArguments array — guards the `.length > 0` check', () => {
    const md = renderCommandMarkdown({ ...listDrives, positionalArguments: [] });
    expect(md).not.toContain('## Positional arguments');
  });

  // 2026-07-24: one name per flag, one per command. The alias-suffix and
  // commandAliases rendering these two tests covered was deleted with it.
  it('renders an option row as flag + description, with no alias annotation', () => {
    const md = renderCommandMarkdown({ ...calendarEvent, options: [{ name: 'event-id', key: 'eventId', required: true, description: 'The Graph event ID.' }] });
    expect(md).toContain('| `--event-id` | The Graph event ID. |');
    expect(md).not.toContain('aliases:');
  });
});

describe('CATEGORY_ORDER', () => {
  // CATEGORY_ORDER is the single source the MCP list-commands description derives its category list
  // from. Pinning it to CATEGORY_LABELS (a Record over every CommandCategory) guarantees a new
  // category cannot be added without appearing in the advertised list too.
  it('lists every command category, so a list derived from it can never silently omit one', () => {
    const advertised: readonly string[] = CATEGORY_ORDER;
    expect([...advertised].toSorted((a, b) => a.localeCompare(b))).toEqual(Object.keys(CATEGORY_LABELS).toSorted((a, b) => a.localeCompare(b)));
  });
});

describe('docs-render — exact output, which is the only thing that pins the blank lines', () => {
  // Every `it` above asserts with toContain, so the SPACING between sections is
  // unasserted: the empty-string literals in each `lines.push('', '## X', '')`
  // can be replaced with junk and every one of those tests still passes. The
  // rendered markdown is a product (docs/COMMANDS.md, and what `docs <cmd>`
  // prints), so asserting it whole is the point rather than a coverage trick.
  const optionOnly: CommandManifestEntry = {
    name: 'list-things',
    summary: 'List the things.',
    category: 'drive',
    graphMethod: 'GET',
    graphPathTemplate: '/me/things',
    graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/thing-list',
    options: [{ name: 'top', key: 'top', required: false, description: 'Page size.' }],
    example: 'ask-marcel-office list-things',
  };

  it('renders an options-only command as exactly this markdown, blank lines included', () => {
    expect(renderCommandMarkdown(optionOnly)).toBe(
      [
        '# `list-things`',
        '',
        'List the things.',
        '',
        `- **Category:** ${CATEGORY_LABELS['drive']}`,
        '- **Graph endpoint:** `GET /me/things`',
        '- **Microsoft Learn:** https://learn.microsoft.com/en-us/graph/api/thing-list',
        '',
        '## Options',
        '',
        '| Flag | Description |',
        '|------|-------------|',
        '| `--top` | Page size. |',
        '',
        '## Example',
        '',
        '```bash',
        'ask-marcel-office list-things',
        '```',
      ].join('\n')
    );
  });

  it('renders positional arguments and a request body as exactly this markdown', () => {
    const withBody: CommandManifestEntry = {
      name: 'send-thing',
      summary: 'Send a thing.',
      category: 'mail',
      graphMethod: 'POST',
      graphPathTemplate: '/me/sendThing',
      graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/thing-send',
      options: [],
      positionalArguments: [{ name: 'target', required: true, description: 'Who to send it to.' }],
      bodyTemplate: '{"a":1}',
      example: 'ask-marcel-office send-thing bob',
    };

    expect(renderCommandMarkdown(withBody)).toBe(
      [
        '# `send-thing`',
        '',
        'Send a thing.',
        '',
        `- **Category:** ${CATEGORY_LABELS['mail']}`,
        '- **Graph endpoint:** `POST /me/sendThing`',
        '- **Microsoft Learn:** https://learn.microsoft.com/en-us/graph/api/thing-send',
        '',
        '## Positional arguments',
        '',
        '| Argument | Required | Description |',
        '|----------|----------|-------------|',
        '| `<target>` | yes | Who to send it to. |',
        '',
        '## Request body',
        '',
        '```json',
        '{"a":1}',
        '```',
        '',
        '## Example',
        '',
        '```bash',
        'ask-marcel-office send-thing bob',
        '```',
      ].join('\n')
    );
  });

  it('renders the README tables as exactly this markdown, pinning the section join and the required-params separator', () => {
    const twoParams: CommandManifestEntry = {
      name: 'get-thing',
      summary: 'Get a thing.',
      category: 'drive',
      graphMethod: 'GET',
      graphPathTemplate: '/me/things/{id}',
      graphDocsUrl: 'https://learn.microsoft.com/x',
      options: [
        { name: 'drive-id', key: 'driveId', required: true, description: 'Drive.' },
        { name: 'item-id', key: 'itemId', required: true, description: 'Item.' },
      ],
      example: 'ask-marcel-office get-thing',
    };
    const mailOne: CommandManifestEntry = { ...twoParams, name: 'list-mail', summary: 'List mail.', category: 'mail', options: [] };

    const rendered = renderReadmeTables({ package: 'p', version: '1', generatedAt: 'now', commands: [twoParams, mailOne] });

    expect(rendered).toBe(
      [
        `### ${CATEGORY_LABELS['drive']}`,
        '',
        '| Command | Description | Required params | Graph endpoint |',
        '|---------|-------------|-----------------|----------------|',
        '| `get-thing` | Get a thing. | `--drive-id`, `--item-id` | `GET /me/things/{id}` |',
        '',
        `### ${CATEGORY_LABELS['mail']}`,
        '',
        '| Command | Description | Required params | Graph endpoint |',
        '|---------|-------------|-----------------|----------------|',
        '| `list-mail` | List mail. | _(none)_ | `GET /me/things/{id}` |',
      ].join('\n')
    );
  });

  it('gives every category its own non-empty label, so none renders as a blank heading', () => {
    const seen = new Set<string>();
    for (const category of CATEGORY_ORDER) {
      const label = CATEGORY_LABELS[category];
      expect(label.trim().length).toBeGreaterThan(0);
      expect(seen.has(label)).toBe(false);
      seen.add(label);
    }
    expect(seen.size).toBe(CATEGORY_ORDER.length);
  });

  it('omits the elevated-token line, and the preferMaxPageSize hint, when the entry does not call for them', () => {
    const plain = renderCommandMarkdown(optionOnly);
    expect(plain).not.toContain('Needs elevated token');
    const paginated = renderCommandMarkdown({ ...optionOnly, pagination: true });
    expect(paginated).not.toContain('odata.maxpagesize');
  });
});
