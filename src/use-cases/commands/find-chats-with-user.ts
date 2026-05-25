import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Command, CommandMeta } from './command-types.ts';
import { formatZodError } from './format-zod-error.ts';

// Cross-chat member search — collapses the "all my conversations with
// person X" workflow into one call. Walks every page of
// `list-teams-chats-with-messages`'s paginated `/chats` endpoint and
// returns any chat whose `members[]` carries a substring match against
// the query — on display-name, email, given-name, surname, MRI, or
// object-id. Critical for dual-identity people whose org and guest
// MRIs live in different chats.
const schema = z.object({
  name: z.string().min(1),
  maxPages: z
    .string()
    .regex(/^[1-9]\d*$/, 'must be a positive integer')
    .optional(),
  pageSize: z
    .string()
    .regex(/^[1-9]\d*$/, 'must be a positive integer')
    .optional(),
});

const QUERY_BASE = 'enableMembershipSummary=true&supportsAdditionalSystemGeneratedFolders=true&supportsSliceItems=true&enableEngageCommunities=false';

// Match a query against a member by checking every text-bearing field for a
// substring, after folding both sides:
//   - Unicode NFD normalization, then stripping combining-mark codepoints
//     (so `é` ↔ `e`, `ç` ↔ `c`, etc.).
//   - Lowercasing.
//
// Audit Alex-session §D follow-up: a real dual-identity user had the
// CORPORATE-MRI member entry's displayName populated as the email
// (`alex.kim@example.com`, no accent) while the GUEST-MRI entry's
// displayName carried the accented "Alex Kim". A search for `Alex`
// returned only the guest chat; the corporate 1:1 was invisible because
// `é`.toLowerCase() and `e` are different bytes. Folding diacritics on
// both sides makes `Alex` ↔ `Alex` ↔ `JANE` ↔ `alex.kim@example.com`
// all match against the same query.
type Member = {
  readonly mri?: string;
  readonly objectId?: string;
  readonly displayName?: string;
  readonly email?: string;
  readonly userPrincipalName?: string;
  readonly givenName?: string;
  readonly surname?: string;
  readonly jobTitle?: string;
  readonly userSubType?: string;
};

const fold = (s: string): string =>
  s
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
    .toLowerCase();

const memberMatches = (member: Member, queryFolded: string): boolean => {
  const haystacks: ReadonlyArray<string | undefined> = [
    member.displayName,
    member.email,
    member.userPrincipalName,
    member.givenName,
    member.surname,
    member.mri,
    member.objectId,
    member.jobTitle,
  ];
  return haystacks.some((h) => typeof h === 'string' && fold(h).includes(queryFolded));
};

type Chat = {
  readonly id?: string;
  readonly title?: string | null;
  readonly chatType?: string;
  readonly threadType?: string;
  readonly members?: ReadonlyArray<Member>;
  readonly lastMessage?: { readonly composeTime?: string };
};
type ChatsResponse = { readonly chats?: ReadonlyArray<Chat>; readonly continuationToken?: string; readonly hasMoreData?: boolean };

// Project a chat into the find-chats-with-user response shape: only the
// fields a downstream consumer needs to act (chat-id to call other
// commands, plus enough member context to confirm it's the right chat).
type MatchedMember = {
  readonly mri?: string;
  readonly displayName?: string;
  readonly email?: string;
  readonly userSubType?: string;
};
type MatchedChat = {
  readonly chatId: string;
  readonly title: string | null;
  readonly chatType?: string;
  readonly threadType?: string;
  readonly memberCount: number;
  readonly lastMessageAt?: string;
  readonly matchedMembers: ReadonlyArray<MatchedMember>;
};

const projectMember = (m: Member): MatchedMember => {
  const out: { mri?: string; displayName?: string; email?: string; userSubType?: string } = {};
  if (m.mri !== undefined) out.mri = m.mri;
  if (m.displayName !== undefined) out.displayName = m.displayName;
  if (m.email !== undefined) out.email = m.email;
  if (m.userSubType !== undefined) out.userSubType = m.userSubType;
  return out;
};

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const queryFolded = fold(parsed.data.name);
  const pageSize = parsed.data.pageSize ?? '100';
  const maxPages = Number(parsed.data.maxPages ?? '10');

  const matched: Array<MatchedChat> = [];
  // Dedup keys by chat id in case a chat shows up twice (defense in depth —
  // the paginated endpoint emits unique chats per page, but a future cursor
  // shape change shouldn't silently double-count).
  const seenChatIds = new Set<string>();
  let continuationToken: string | undefined;
  let pagesFetched = 0;
  let chatsScanned = 0;

  for (let page = 0; page < maxPages; page += 1) {
    const qs = new URLSearchParams({ pageSize });
    if (continuationToken !== undefined) qs.set('continuationToken', continuationToken);
    const result = await graph.teamsChat(`/api/v3/teams/users/me/chats?${qs.toString()}&${QUERY_BASE}`);
    if (!result.ok) return result;
    const body = result.value as ChatsResponse;
    const chats = body.chats ?? [];
    pagesFetched += 1;
    chatsScanned += chats.length;
    for (const chat of chats) {
      if (chat.id === undefined || seenChatIds.has(chat.id)) continue;
      const members = chat.members ?? [];
      const matchedMembers = members.filter((m) => memberMatches(m, queryFolded));
      if (matchedMembers.length === 0) continue;
      seenChatIds.add(chat.id);
      matched.push({
        chatId: chat.id,
        title: chat.title ?? null,
        chatType: chat.chatType,
        threadType: chat.threadType,
        memberCount: members.length,
        ...(chat.lastMessage?.composeTime !== undefined ? { lastMessageAt: chat.lastMessage.composeTime } : {}),
        matchedMembers: matchedMembers.map(projectMember),
      });
    }
    if (body.hasMoreData !== true || body.continuationToken === undefined) {
      continuationToken = undefined;
      break;
    }
    continuationToken = body.continuationToken;
  }

  return ok({
    name: parsed.data.name,
    matches: matched,
    matchCount: matched.length,
    pagesFetched,
    chatsScanned,
    hasMore: continuationToken !== undefined,
    nextContinuationToken: continuationToken,
  });
};

const meta: CommandMeta = {
  summary:
    'Find every Microsoft Teams chat that includes a member matching `--name` (substring search across display-name, email, given-name, surname, MRI, and object-id). Both sides are Unicode-folded (NFD + combining-mark strip) and lowercased before comparison, so `--name Alex` matches `Alex Kim` AND `alex.kim@example.com` AND `JANE` — important because a dual-identity user often carries the accented display-name on one identity and the un-accented email on the other. Walks the paginated chat-list substrate up to `--max-pages` and returns matching chats with their `matchedMembers[]`. Collapses the canonical "all conversations with person X" workflow into a single call AND surfaces dual-identity people (e.g. someone with both an org MRI and a guest-tenant MRI). **Best-effort, may break on Microsoft client updates** — the chat substrate is not in the public Microsoft Graph API.',
  category: 'chats',
  graphMethod: 'GET',
  graphPathTemplate: 'https://teams.microsoft.com/api/csa/{region}/api/v3/teams/users/me/chats',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chat-list',
  options: [
    {
      name: 'name',
      key: 'name',
      required: true,
      description:
        "Substring to search across each chat member's `displayName`, `email`, `userPrincipalName`, `givenName`, `surname`, `mri`, `objectId`, and `jobTitle`. Both the query and each field are NFD-normalized + diacritics-stripped + lowercased before comparison, so `Alex` ↔ `Alex` ↔ `JANE` are equivalent and a query for the accented name still matches a member whose displayName is the un-accented email. Use the full name or an unambiguous fragment. Quoted multi-word values match on the joined substring, not per-token.",
    },
    {
      name: 'max-pages',
      key: 'maxPages',
      required: false,
      description:
        'Safety cap on the chat-list walk (positive integer; default 10). Each page returns up to `--page-size` chats. Raise carefully on busy accounts — every page is one HTTP round-trip.',
    },
    {
      name: 'page-size',
      key: 'pageSize',
      required: false,
      description: 'Chats per page (positive integer; default 100, same value Teams web uses). Server may silently cap.',
    },
  ],
  example: "ask-marcel find-chats-with-user --name 'Alex Kim'",
  responseShape:
    "`{ name, matches: [{ chatId, title, chatType, threadType, memberCount, lastMessageAt?, matchedMembers: [{ mri, displayName, email, userSubType }] }], matchCount, pagesFetched, chatsScanned, hasMore, nextContinuationToken? }`. `matchedMembers` always carries the matching entries' identifying fields — pass `chatId` into `list-teams-chat-history` to read message bodies. `hasMore: true` means `--max-pages` was hit before exhausting the chat list; chain with the existing `--continuation-token` flag on `list-teams-chats-with-messages` if you need to scan further (this command does not advertise a `--continuation-token` because resuming a partial search is rare; users either widen `--max-pages` or refine `--name`).",
  stability: 'experimental',
};

export { execute, meta, schema };
