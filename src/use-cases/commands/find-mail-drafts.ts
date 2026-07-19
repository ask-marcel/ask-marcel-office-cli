import { z } from 'zod';
import { err, ok, type Result } from '../../domain/result.ts';
import type { GraphError } from '../../infra/graph-client.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { matchExistingDrafts } from './draft-dedup.ts';
import { formatZodError } from './format-zod-error.ts';

const schema = z.object({
  subject: z.string().min(1),
  toRecipients: z.string().min(1).optional(),
});

const SCAN_LIMIT = 50;
// Hardcoded path: the space in `$orderby` is percent-encoded by hand, as there
// is no query builder on this route. Newest-first so the window holds the drafts
// a caller most likely just created.
const SCAN_PATH = `/me/mailFolders/drafts/messages?$top=${SCAN_LIMIT}&$orderby=lastModifiedDateTime%20desc&$select=id,subject,toRecipients,ccRecipients,conversationId,lastModifiedDateTime,webLink`;

const recipientSchema = z.object({ emailAddress: z.object({ address: z.string().optional() }).optional() });
const draftSchema = z.object({
  id: z.string(),
  subject: z.string().optional(),
  conversationId: z.string().optional(),
  toRecipients: z.array(recipientSchema).optional(),
  ccRecipients: z.array(recipientSchema).optional(),
  lastModifiedDateTime: z.string().optional(),
  webLink: z.string().optional(),
});
const draftListSchema = z.object({ value: z.array(draftSchema).optional() });

const parseRecipients = (raw: string | undefined): ReadonlyArray<string> | undefined => {
  if (raw === undefined) return undefined;
  return raw
    .split(',')
    .map((address) => address.trim())
    .filter((address) => address.length > 0);
};

const execute: Command['execute'] = async (graph, params): Promise<Result<unknown, GraphError>> => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const { subject, toRecipients } = parsed.data;

  const listed = await graph.get(SCAN_PATH);
  if (!listed.ok) return listed;

  const drafts = draftListSchema.safeParse(listed.value);
  if (!drafts.success) {
    return err({
      type: 'api_error',
      status: 500,
      code: 'drafts_list_unreadable',
      message:
        'The Drafts folder listing came back in an unreadable shape, so drafts could not be de-duplicated. Retry, or list the folder with `list-mail-folder-messages --mail-folder-id drafts`.',
    });
  }

  const scanned = drafts.data.value ?? [];
  const recipients = parseRecipients(toRecipients);
  const matches = matchExistingDrafts(scanned, { subject, ...(recipients === undefined ? {} : { recipients }) });
  const conversationIds = [...new Set(matches.flatMap((match) => (match.conversationId === undefined ? [] : [match.conversationId])))];

  return ok({ matches, conversationIds, scanned: scanned.length, scanLimit: SCAN_LIMIT });
};

const meta: CommandMeta = {
  summary:
    'Find existing drafts on a mail thread WITHOUT trusting a conversationId $filter. Reply and forward drafts do not always inherit the inbound message conversationId (a thread can split across several), and Graph $filter on the Drafts folder is not read-your-writes consistent, so filtering Drafts by conversationId misses drafts. This command instead scans the 50 most recently modified drafts and matches them client-side on a normalized subject (stripping RE:/FW: and localized reply/forward prefixes) plus, optionally, a shared recipient. Use it before create-reply-draft to avoid creating a duplicate: if a match comes back, revise it with update-mail-draft instead of making a new one. Read-only.',
  category: 'mail',
  graphMethod: 'GET',
  graphPathTemplate:
    '/me/mailFolders/drafts/messages?$top=50&$orderby=lastModifiedDateTime desc&$select=id,subject,toRecipients,ccRecipients,conversationId,lastModifiedDateTime,webLink (a read-only scan of the 50 most recently modified drafts; each is matched CLIENT-SIDE on a normalized {subject} and, when given, a shared {to-recipients} address, so neither value is sent to Graph)',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/mailfolder-list-messages',
  options: [
    {
      name: 'subject',
      key: 'subject',
      required: true,
      description:
        'The thread subject to match. Reply and forward prefixes (RE:, FW:, and localized variants) are stripped on both sides before comparing, so "Contoso Q3" matches a draft titled "RE: Contoso Q3".',
    },
    {
      name: 'to-recipients',
      key: 'toRecipients',
      required: false,
      description:
        'Optional comma-separated recipient addresses. When given, a draft matches only if it shares at least one To or Cc address with this list, narrowing subject-only collisions.',
    },
  ],
  example: 'ask-marcel-office find-mail-drafts --subject "Contoso Q3 budget" --to-recipients "kim@example.com"',
  scopesRequired: ['Mail.Read'],
  responseShape:
    '`{ matches: message[], conversationIds: string[], scanned: number, scanLimit: number }`. `matches` are the drafts whose normalized subject (RE:/FW:/localized prefixes stripped) equals yours, each carrying `{ id, subject, toRecipients, ccRecipients, conversationId, lastModifiedDateTime, webLink }`; pass a match `id` to update-mail-draft to revise it instead of creating a duplicate. `conversationIds` is the de-duplicated union of conversationIds across the matches (a thread can span several). `scanned` is how many drafts were examined and `scanLimit` the cap (50): when `scanned` equals `scanLimit`, older drafts may exist beyond the window.',
};

export { execute, meta, schema };
