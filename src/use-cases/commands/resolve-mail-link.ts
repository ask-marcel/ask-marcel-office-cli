import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

// Deep-link parser for Microsoft Outlook web links. Outlook emits several
// URL shapes for a single message depending on which client surfaced the
// "Copy link" / "Share" action:
//
//   Classic OWA query-style:
//     https://outlook.office365.com/owa/?ItemID=AAMkA...&exvsurl=1&viewmodel=ReadMessageItem
//     https://outlook.office.com/owa/?itemid=AAMkA...&exvsurl=1&path=/mail/inbox
//
//   Modern path-style:
//     https://outlook.office.com/mail/inbox/id/AAMkA...
//     https://outlook.office365.com/mail/AAMkA...
//
// Pure transformation — no HTTP. Pair with `get-mail-message` to fetch the
// body once the link is resolved. Calendar links share the OWA query host
// but carry `path=/calendar/item` (or live under `/calendar/item/...`); this
// command rejects those with a pointer to `resolve-calendar-link` so the
// LLM never silently treats a calendar invite as a mail message.
const schema = z.object({
  url: z.string().min(1).url(),
});

const OUTLOOK_HOSTS: ReadonlyArray<string> = ['outlook.office.com', 'outlook.office365.com', 'outlook.live.com'];

type Resolved = {
  readonly messageId: string;
};

type ParseOutcome = { readonly kind: 'ok'; readonly value: Resolved } | { readonly kind: 'calendar' } | { readonly kind: 'unknown' };

// Extract the ID from either an OWA-style query (`?itemid=` / `?ItemID=`)
// or a modern path-style (`/mail/.../id/<id>` or `/mail/<id>`).
const extractFromOwaQuery = (parsed: URL): string | null => {
  const itemId = parsed.searchParams.get('itemid') ?? parsed.searchParams.get('ItemID') ?? parsed.searchParams.get('itemId');
  return itemId !== null && itemId !== '' ? itemId : null;
};

const extractFromMailPath = (path: string): string | null => {
  // `/mail/<folder>/id/<id>` or `/mail/<id>` — accept both. The "id"
  // segment marker is the modern shape; the bare `/mail/<id>` is the
  // legacy short link.
  const segments = path.split('/').filter((s) => s.length > 0);
  if (segments.length === 0 || segments[0] !== 'mail') return null;
  const idMarkerIdx = segments.indexOf('id');
  if (idMarkerIdx !== -1 && idMarkerIdx + 1 < segments.length) {
    const candidate = segments[idMarkerIdx + 1];
    return candidate !== undefined && candidate !== '' ? candidate : null;
  }
  // `/mail/<id>` — the second segment is the id, but only if it looks
  // like an Outlook message id (starts with `AAMk` or is otherwise long).
  // Folder names like `inbox` / `sentitems` are short and lowercase; ids
  // are base64url-ish. Reject short alphabetic-only segments.
  const second = segments[1];
  if (second === undefined || second.length < 20) return null;
  return second;
};

const isCalendarLink = (parsed: URL): boolean => {
  if (parsed.pathname.startsWith('/calendar/')) return true;
  const pathParam = parsed.searchParams.get('path');
  if (pathParam !== null && pathParam.startsWith('/calendar')) return true;
  return false;
};

const parse = (raw: string): ParseOutcome => {
  // No try/catch — Zod's `.url()` refinement on the input schema already
  // validated the URL format before this parser runs, so `new URL(raw)`
  // cannot throw. Also keeps the file try/catch-free per atelier rule
  // (try/catch is restricted to `src/infra/**`).
  const url = new URL(raw);
  if (!OUTLOOK_HOSTS.includes(url.hostname)) return { kind: 'unknown' };
  if (isCalendarLink(url)) return { kind: 'calendar' };

  // Try OWA query-style first, then modern path-style.
  const owaId = url.pathname.startsWith('/owa') ? extractFromOwaQuery(url) : null;
  if (owaId !== null) return { kind: 'ok', value: { messageId: decodeURIComponent(owaId) } };
  const pathId = extractFromMailPath(url.pathname);
  if (pathId !== null) return { kind: 'ok', value: { messageId: decodeURIComponent(pathId) } };

  return { kind: 'unknown' };
};

const execute: Command['execute'] = async (_graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const outcome = parse(parsed.data.url);
  if (outcome.kind === 'ok') return ok(outcome.value);
  if (outcome.kind === 'calendar') {
    return err({
      type: 'validation_error',
      message: '--url looks like an Outlook calendar link, not a mail message link.',
      code: 'cli_reject_calendar_link_on_mail_resolver',
    });
  }
  return err({
    type: 'validation_error',
    message: `--url: not an Outlook mail link. Expected shapes: \`https://outlook.office.com/owa/?itemid=AAMkA...\`, \`https://outlook.office.com/mail/inbox/id/AAMkA...\`, or \`https://outlook.office.com/mail/AAMkA...\`. Hosts accepted: ${OUTLOOK_HOSTS.join(', ')}.`,
  });
};

const meta: CommandMeta = {
  summary:
    'Parse a Microsoft Outlook web mail link (the URL emitted by the "Copy link" / address-bar share of an email) into its `messageId`. Pure transformation — no Graph call. Pipe the result into `get-mail-message` to fetch the body, or `convert-mail-to-markdown` to render it. For Outlook calendar links use `resolve-calendar-link` instead — this command rejects them with a pointer.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate: '{url}',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/message-get',
  options: [
    {
      name: 'url',
      key: 'url',
      required: true,
      description:
        'Outlook web URL for a single mail message. Accepted hosts: `outlook.office.com`, `outlook.office365.com`, `outlook.live.com`. Accepted shapes: OWA query-style (`/owa/?itemid=AAMkA...` or `/owa/?ItemID=AAMkA...`), modern path-style (`/mail/inbox/id/AAMkA...`), legacy short (`/mail/AAMkA...`). Calendar links (`/calendar/item/...` or `?path=/calendar/item`) are rejected with `cli_reject_calendar_link_on_mail_resolver` — use `resolve-calendar-link` for those.',
    },
  ],
  example: "ask-marcel resolve-mail-link --url 'https://outlook.office.com/mail/inbox/id/AAMkAGI2THVS...'",
  responseShape: '`{ messageId: string }`. `messageId` is URL-decoded and ready to pass to `get-mail-message --message-id <id>` or `convert-mail-to-markdown --message-id <id>`.',
};

export { execute, meta, schema };
