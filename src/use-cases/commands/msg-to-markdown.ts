import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import { extractMsg } from '../../infra/msg-reader-adapter.ts';
import type { MsgAttachment, MsgRecipient, MsgRecipientKind, ParsedMsg } from '../../infra/msg-reader-adapter.ts';
import { htmlToMarkdown } from '../../infra/turndown-adapter.ts';

/**
 * Render an Outlook `.msg` file to markdown: an H1 subject, a From/To/Cc/Bcc/Date
 * header block, the message body (the plain-text body when present, else the HTML
 * body run through turndown), and an `## Attachments` section where each attachment
 * is converted recursively through the SAME dispatch the zip command uses — so a
 * `.docx`/`.pdf`/`.csv` attached to an email becomes markdown inline, exactly as the
 * user asked ("maybe same way as zip"). Unconvertible attachments (images, binaries)
 * are listed with the dispatch's note instead of failing the whole message.
 *
 * `recurse` is the attachment converter injected by `markdown-dispatch` (its own
 * `bytesToMarkdown`, with the recursion depth incremented). Injecting it — rather
 * than importing `bytesToMarkdown` here — keeps this use-case off the dispatch's
 * import cycle and lets every rendering branch be tested with a plain fake. `depth`
 * caps `.msg`-inside-`.msg` nesting: past {@link MAX_MSG_DEPTH} levels an embedded
 * message's attachments are listed but not expanded.
 */

const MAX_MSG_DEPTH = 3;

type MsgToMarkdownOptions = { readonly depth?: number };
type MsgAttachmentConverter = (bytes: Uint8Array, filename: string) => Promise<Result<unknown, GraphError>>;

const formatAddress = (name: string | undefined, email: string | undefined): string | undefined => {
  if (name !== undefined && name !== '') return email !== undefined && email !== '' ? `${name} <${email}>` : name;
  return email !== undefined && email !== '' ? email : undefined;
};

const headerLine = (label: string, value: string | undefined): string | undefined => (value !== undefined && value !== '' ? `**${label}:** ${value}` : undefined);

const recipientsByKind = (recipients: readonly MsgRecipient[], kind: MsgRecipientKind): string | undefined => {
  const parts = recipients
    .filter((r) => r.kind === kind)
    .map((r) => formatAddress(r.name, r.email))
    .filter((s): s is string => s !== undefined);
  return parts.length > 0 ? parts.join(', ') : undefined;
};

const renderHeader = (msg: ParsedMsg): string =>
  [
    headerLine('From', formatAddress(msg.senderName, msg.senderEmail)),
    headerLine('To', recipientsByKind(msg.recipients, 'to')),
    headerLine('Cc', recipientsByKind(msg.recipients, 'cc')),
    headerLine('Bcc', recipientsByKind(msg.recipients, 'bcc')),
    headerLine('Recipients', recipientsByKind(msg.recipients, 'unknown')),
    headerLine('Date', msg.date),
  ]
    .filter((s): s is string => s !== undefined)
    .join('\n');

const renderBody = (msg: ParsedMsg): string => {
  if (msg.body !== undefined && msg.body.trim() !== '') return msg.body.trim();
  if (msg.bodyHtml !== undefined && msg.bodyHtml !== '') {
    const converted = htmlToMarkdown(msg.bodyHtml);
    return converted.ok ? converted.value : msg.bodyHtml;
  }
  return '';
};

const renderAttachment = async (attachment: MsgAttachment, depth: number, recurse: MsgAttachmentConverter): Promise<string> => {
  const name = attachment.fileName !== undefined && attachment.fileName !== '' ? attachment.fileName : 'unnamed';
  const heading = `### ${name}`;
  if (attachment.content === undefined) return `${heading}\n\n_(no readable content)_`;
  if (depth >= MAX_MSG_DEPTH) return `${heading}\n\n_(embedded message too deeply nested — attachment not expanded)_`;
  const converted = await recurse(attachment.content, name);
  if (!converted.ok) return `${heading}\n\n_${converted.error.message}_`;
  const envelope = converted.value as { readonly text?: string };
  return `${heading}\n\n${envelope.text ?? ''}`;
};

const renderAttachments = async (attachments: readonly MsgAttachment[], depth: number, recurse: MsgAttachmentConverter): Promise<string> => {
  if (attachments.length === 0) return '';
  const sections = await Promise.all(attachments.map((attachment) => renderAttachment(attachment, depth, recurse)));
  return ['## Attachments', ...sections].join('\n\n');
};

const renderMsg = async (msg: ParsedMsg, depth: number, recurse: MsgAttachmentConverter): Promise<string> => {
  const sections: string[] = [];
  if (msg.subject !== undefined && msg.subject !== '') sections.push(`# ${msg.subject}`);
  const header = renderHeader(msg);
  if (header !== '') sections.push(header);
  const body = renderBody(msg);
  if (body !== '') sections.push(body);
  const attachments = await renderAttachments(msg.attachments, depth, recurse);
  if (attachments !== '') sections.push(attachments);
  return sections.join('\n\n');
};

const msgToMarkdown = async (bytes: Uint8Array, opts: MsgToMarkdownOptions, recurse: MsgAttachmentConverter): Promise<Result<unknown, GraphError>> => {
  const parsed = await extractMsg(bytes);
  if (!parsed.ok) return parsed;
  const text = await renderMsg(parsed.value, opts.depth ?? 0, recurse);
  return ok({ contentType: 'text/markdown', size: new TextEncoder().encode(text).byteLength, text });
};

export { MAX_MSG_DEPTH, msgToMarkdown, renderMsg };
export type { MsgAttachmentConverter, MsgToMarkdownOptions };
