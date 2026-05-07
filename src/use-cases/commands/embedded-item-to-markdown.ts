import { htmlToMarkdown } from '../../infra/turndown-adapter.ts';

/**
 * Render an itemAttachment's inner resource (a `microsoft.graph.message`,
 * `event`, or `contact`) directly to markdown — no RTF round-trip, no
 * Graph conversion call.
 *
 * Each renderer emits a small header block (`**Subject:** …`, etc.) and
 * appends the body run through htmlToMarkdown when the inner resource
 * has one. Callers (the convert-mail-attachment-to-markdown command)
 * choose the right renderer by branching on `item['@odata.type']`.
 */

type EmailAddress = { readonly name?: string; readonly address?: string };
type Recipient = { readonly emailAddress?: EmailAddress };
type Body = { readonly contentType?: string; readonly content?: string };

const formatAddress = (a: EmailAddress | undefined): string | undefined => {
  if (!a?.address) return undefined;
  return a.name ? `${a.name} <${a.address}>` : a.address;
};

const formatRecipients = (rs: ReadonlyArray<Recipient> | undefined): string | undefined => {
  if (!rs || rs.length === 0) return undefined;
  const parts = rs.map((r) => formatAddress(r.emailAddress)).filter((s): s is string => s !== undefined);
  return parts.length > 0 ? parts.join(', ') : undefined;
};

const renderBody = (body: Body | undefined): string => {
  if (!body?.content) return '';
  if (body.contentType !== 'html') return body.content;
  const result = htmlToMarkdown(body.content);
  return result.ok ? result.value : body.content;
};

const headerLine = (label: string, value: string | undefined): string | undefined => (value !== undefined ? `**${label}:** ${value}` : undefined);

type EmbeddedMessage = {
  readonly subject?: string;
  readonly from?: Recipient;
  readonly toRecipients?: ReadonlyArray<Recipient>;
  readonly sentDateTime?: string;
  readonly body?: Body;
};

const embeddedMessageToMarkdown = (m: EmbeddedMessage): string => {
  const headers = [
    headerLine('Subject', m.subject),
    headerLine('From', formatAddress(m.from?.emailAddress)),
    headerLine('To', formatRecipients(m.toRecipients)),
    headerLine('Sent', m.sentDateTime),
  ].filter((s): s is string => s !== undefined);
  const body = renderBody(m.body);
  return [headers.join('\n'), body].filter((s) => s !== '').join('\n\n');
};

type EventDateTime = { readonly dateTime?: string; readonly timeZone?: string };
type Location = { readonly displayName?: string };
type EmbeddedEvent = {
  readonly subject?: string;
  readonly start?: EventDateTime;
  readonly end?: EventDateTime;
  readonly location?: Location;
  readonly organizer?: Recipient;
  readonly attendees?: ReadonlyArray<Recipient>;
  readonly body?: Body;
};

const formatEventDate = (d: EventDateTime | undefined): string | undefined => {
  if (!d?.dateTime) return undefined;
  return d.timeZone ? `${d.dateTime} ${d.timeZone}` : d.dateTime;
};

const embeddedEventToMarkdown = (e: EmbeddedEvent): string => {
  const headers = [
    headerLine('Subject', e.subject),
    headerLine('Start', formatEventDate(e.start)),
    headerLine('End', formatEventDate(e.end)),
    headerLine('Location', e.location?.displayName),
    headerLine('Organizer', formatAddress(e.organizer?.emailAddress)),
    headerLine('Attendees', formatRecipients(e.attendees)),
  ].filter((s): s is string => s !== undefined);
  const body = renderBody(e.body);
  return [headers.join('\n'), body].filter((s) => s !== '').join('\n\n');
};

type EmbeddedContact = {
  readonly displayName?: string;
  readonly emailAddresses?: ReadonlyArray<{ address?: string }>;
  readonly businessPhones?: ReadonlyArray<string>;
  readonly mobilePhone?: string;
  readonly companyName?: string;
  readonly jobTitle?: string;
};

const embeddedContactToMarkdown = (c: EmbeddedContact): string => {
  const emails = c.emailAddresses?.map((e) => e.address).filter((s): s is string => s !== undefined && s !== '') ?? [];
  const businessPhone = c.businessPhones && c.businessPhones.length > 0 ? c.businessPhones[0] : undefined;
  return [
    headerLine('Name', c.displayName),
    headerLine('Emails', emails.length > 0 ? emails.join(', ') : undefined),
    headerLine('Business phone', businessPhone),
    headerLine('Mobile phone', c.mobilePhone),
    headerLine('Company', c.companyName),
    headerLine('Title', c.jobTitle),
  ]
    .filter((s): s is string => s !== undefined)
    .join('\n');
};

export { embeddedContactToMarkdown, embeddedEventToMarkdown, embeddedMessageToMarkdown };
export type { EmbeddedContact, EmbeddedEvent, EmbeddedMessage };
