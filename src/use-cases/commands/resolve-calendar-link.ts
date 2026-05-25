import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

// Deep-link parser for Microsoft Outlook calendar item URLs. Outlook
// emits two shapes for a single event:
//
//   Path-style (modern):
//     https://outlook.office.com/calendar/item/AAMkA...
//     https://outlook.office365.com/calendar/item/AAMkA...
//
//   OWA query-style:
//     https://outlook.office.com/owa/?itemid=AAMkA...&exvsurl=1&path=/calendar/item
//
// Pure transformation — no HTTP. Pair with `get-calendar-event` to fetch
// the event body once the link is resolved. Mail message links share the
// OWA query host; this command rejects them with a pointer to
// `resolve-mail-link` (mirror of `resolve-mail-link`'s calendar rejection).
const schema = z.object({
  url: z.string().min(1).url(),
});

const OUTLOOK_HOSTS: ReadonlyArray<string> = ['outlook.office.com', 'outlook.office365.com', 'outlook.live.com'];

type Resolved = {
  readonly eventId: string;
};

type ParseOutcome = { readonly kind: 'ok'; readonly value: Resolved } | { readonly kind: 'mail' } | { readonly kind: 'unknown' };

const extractFromCalendarPath = (path: string): string | null => {
  // `/calendar/item/<id>` — the third segment is the id.
  const segments = path.split('/').filter((s) => s.length > 0);
  if (segments.length < 3 || segments[0] !== 'calendar' || segments[1] !== 'item') return null;
  const candidate = segments[2];
  return candidate !== undefined && candidate !== '' ? candidate : null;
};

const extractFromOwaCalendarQuery = (parsed: URL): string | null => {
  const pathParam = parsed.searchParams.get('path');
  if (pathParam === null || !pathParam.startsWith('/calendar')) return null;
  const itemId = parsed.searchParams.get('itemid') ?? parsed.searchParams.get('ItemID') ?? parsed.searchParams.get('itemId');
  return itemId !== null && itemId !== '' ? itemId : null;
};

const isMailLink = (parsed: URL): boolean => {
  if (parsed.pathname.startsWith('/mail/')) return true;
  if (parsed.pathname.startsWith('/owa')) {
    const pathParam = parsed.searchParams.get('path');
    // OWA WITHOUT a calendar `path` query is mail by default.
    if (pathParam === null || pathParam.startsWith('/mail')) return true;
  }
  return false;
};

const parse = (raw: string): ParseOutcome => {
  // No try/catch — Zod's `.url()` refinement on the input schema already
  // validated the URL format before this parser runs. Also satisfies the
  // atelier rule restricting try/catch to `src/infra/**`.
  const url = new URL(raw);
  if (!OUTLOOK_HOSTS.includes(url.hostname)) return { kind: 'unknown' };

  // Path-style first (unambiguous), then OWA with explicit calendar path,
  // then mail-link detection (rejected with pointer).
  const pathId = extractFromCalendarPath(url.pathname);
  if (pathId !== null) return { kind: 'ok', value: { eventId: decodeURIComponent(pathId) } };
  const owaId = url.pathname.startsWith('/owa') ? extractFromOwaCalendarQuery(url) : null;
  if (owaId !== null) return { kind: 'ok', value: { eventId: decodeURIComponent(owaId) } };
  if (isMailLink(url)) return { kind: 'mail' };

  return { kind: 'unknown' };
};

const execute: Command['execute'] = async (_graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const outcome = parse(parsed.data.url);
  if (outcome.kind === 'ok') return ok(outcome.value);
  if (outcome.kind === 'mail') {
    return err({
      type: 'validation_error',
      message: '--url looks like an Outlook mail message link, not a calendar item link.',
      code: 'cli_reject_mail_link_on_calendar_resolver',
    });
  }
  return err({
    type: 'validation_error',
    message: `--url: not an Outlook calendar item link. Expected shapes: \`https://outlook.office.com/calendar/item/AAMkA...\` (path-style), or \`https://outlook.office.com/owa/?itemid=AAMkA...&path=/calendar/item\` (OWA query). Hosts accepted: ${OUTLOOK_HOSTS.join(', ')}.`,
  });
};

const meta: CommandMeta = {
  summary:
    'Parse a Microsoft Outlook calendar item link (the URL emitted by the "Copy link" / share action on a calendar event) into its `eventId`. Pure transformation — no Graph call. Pipe the result into `get-calendar-event` to fetch the event body. For Outlook mail message links use `resolve-mail-link` instead — this command rejects them with a pointer.',
  category: 'calendar',
  graphMethod: 'GET',
  graphPathTemplate: '{url}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/event-get',
  options: [
    {
      name: 'url',
      key: 'url',
      required: true,
      description:
        'Outlook web URL for a single calendar item. Accepted hosts: `outlook.office.com`, `outlook.office365.com`, `outlook.live.com`. Accepted shapes: path-style (`/calendar/item/AAMkA...`) and OWA query-style with calendar path (`/owa/?itemid=AAMkA...&path=/calendar/item`). Mail links (`/mail/...` or `/owa/?itemid=...` without `path=/calendar`) are rejected with `cli_reject_mail_link_on_calendar_resolver` — use `resolve-mail-link` for those.',
    },
  ],
  example: "ask-marcel resolve-calendar-link --url 'https://outlook.office.com/calendar/item/AAMkAGI2THVS...'",
  responseShape: '`{ eventId: string }`. `eventId` is URL-decoded and ready to pass to `get-calendar-event --event-id <id>`.',
};

export { execute, meta, schema };
