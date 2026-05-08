import { describe, expect, it } from 'bun:test';
import type { CommandManifest, CommandManifestEntry } from './docs-render.ts';
import { renderCommandMarkdown, renderReadmeTables } from './docs-render.ts';

const calendarEvent: CommandManifestEntry = {
  name: 'get-calendar-event',
  summary: 'Fetch a single calendar event by ID.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '/me/events/{event-id}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-get',
  options: [{ name: 'event-id', key: 'eventId', required: true, description: 'The Graph event ID.' }],
  example: "ask-marcel get-calendar-event --event-id 'AAMk...'",
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
  example: 'ask-marcel list-drives',
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
    expect(md).toContain("ask-marcel get-calendar-event --event-id 'AAMk...'");
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

  it('appends an _(aliases: ...)_ suffix on options with aliases so the markdown surface matches the manifest (audit v1.0.0 §3.2)', () => {
    const aliased: CommandManifestEntry = {
      ...calendarEvent,
      options: [
        {
          name: 'event-id',
          key: 'eventId',
          required: true,
          description: 'The Graph event ID.',
          aliases: [
            { name: 'id', key: 'id' },
            { name: 'evt-id', key: 'evtId' },
          ],
        },
      ],
    };
    const md = renderCommandMarkdown(aliased);
    expect(md).toContain('_(aliases: `--id`, `--evt-id`)_');
  });

  it('still renders options without aliases unchanged (no suffix) so the format is unambiguous', () => {
    const md = renderCommandMarkdown(calendarEvent);
    expect(md).toContain('| `--event-id` | The Graph event ID. |');
    expect(md).not.toContain('aliases:');
  });
});
