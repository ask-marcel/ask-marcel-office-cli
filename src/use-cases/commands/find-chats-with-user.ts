import { z } from 'zod';
import { err, ok } from '../../domain/result.ts';
import type { Result } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
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
// a real dual-identity user had the
// CORPORATE-MRI member entry's displayName populated as the email
// (`alex.kim@example.com`, no accent) while the GUEST-MRI entry's
// displayName carried the accented "Alex Kim". A search for `Alex`
// returned only the guest chat; the corporate 1:1 was invisible because
// `é`.toLowerCase() and `e` are different bytes. Folding diacritics on
// both sides makes `Alex` ↔ `Alex` ↔ `ALEX` ↔ `alex.kim@example.com`
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

// ── Cross-tenant hydration ──────────────────────────────────────────────
// The substrate summary roster (`enableMembershipSummary=true`) resolves
// names only for identities this tenant can see; an externally-homed member
// comes back bare — `{mri:'8:orgid:<homeOid>'}`, no name/email — so a name
// query can never match it. For every bare DIRECT (1:1) chat we then call
// `GET /chats/{id}/members`, which DOES resolve the external member, and
// re-run the matcher — surfacing 1:1s with a counterpart the roster left bare.
const isNameResolvable = (m: Member): boolean => {
  const fields = [m.displayName, m.email, m.userPrincipalName, m.givenName, m.surname];
  return fields.some((f) => typeof f === 'string' && f.trim() !== '');
};

const bareMemberCount = (members: ReadonlyArray<Member>): number => members.filter((m) => !isNameResolvable(m)).length;

// A Microsoft Graph `conversationMember` as returned by `/chats/{id}/members`.
type GraphConversationMember = { readonly displayName?: string; readonly email?: string; readonly userId?: string };
type GraphMembersResponse = { readonly value?: ReadonlyArray<GraphConversationMember> };

// Map a hydrated Graph member into the matcher's shape, reconstructing the
// substrate MRI (`8:orgid:<userId>`) so a projected match still carries the id
// a caller feeds back to get-chat / list-chat-members.
const fromGraphMember = (g: GraphConversationMember): Member => ({
  ...(g.displayName !== undefined ? { displayName: g.displayName } : {}),
  ...(g.email !== undefined ? { email: g.email } : {}),
  ...(typeof g.userId === 'string' ? { mri: `8:orgid:${g.userId}`, objectId: g.userId } : {}),
});

const toMatchedChat = (chatId: string, chat: Chat, members: ReadonlyArray<Member>): MatchedChat => ({
  chatId,
  title: chat.title ?? null,
  chatType: chat.chatType,
  threadType: chat.threadType,
  memberCount: (chat.members ?? []).length,
  ...(chat.lastMessage?.composeTime !== undefined ? { lastMessageAt: chat.lastMessage.composeTime } : {}),
  matchedMembers: members.map(projectMember),
});

type BareChat = { readonly chatId: string; readonly chat: Chat; readonly bareCount: number };
type ScanState = {
  readonly matched: Array<MatchedChat>;
  readonly bareUnmatched: ReadonlyArray<BareChat>;
  readonly pagesFetched: number;
  readonly chatsScanned: number;
  readonly continuationToken: string | undefined;
};
type ScanAcc = { readonly seen: Set<string>; readonly matched: Array<MatchedChat>; readonly bareUnmatched: Array<BareChat> };

// Cheap pass over one chat: match its summary roster, else record it (with its
// count of bare members) as a hydration candidate. Dedup both outcomes by id.
const collectChat = (chat: Chat, queryFolded: string, acc: ScanAcc): void => {
  if (chat.id === undefined || acc.seen.has(chat.id)) return;
  const members = chat.members ?? [];
  const hits = members.filter((m) => memberMatches(m, queryFolded));
  if (hits.length > 0) {
    acc.seen.add(chat.id);
    acc.matched.push(toMatchedChat(chat.id, chat, hits));
    return;
  }
  const bareCount = bareMemberCount(members);
  if (bareCount > 0) {
    acc.seen.add(chat.id);
    acc.bareUnmatched.push({ chatId: chat.id, chat, bareCount });
  }
};

const scanChatPages = async (graph: GraphClient, queryFolded: string, pageSize: string, maxPages: number): Promise<Result<ScanState, GraphError>> => {
  const acc: ScanAcc = { seen: new Set<string>(), matched: [], bareUnmatched: [] };
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
    for (const chat of chats) collectChat(chat, queryFolded, acc);
    if (body.hasMoreData !== true || body.continuationToken === undefined) {
      continuationToken = undefined;
      break;
    }
    continuationToken = body.continuationToken;
  }
  return ok({ matched: acc.matched, bareUnmatched: acc.bareUnmatched, pagesFetched, chatsScanned, continuationToken });
};

const sumBareCounts = (bare: ReadonlyArray<BareChat>): number => bare.reduce((n, b) => n + b.bareCount, 0);

// Hydrate one bare chat: `null` = Graph couldn't resolve it (stays unresolved);
// otherwise the (possibly empty) list of hydrated members matching the query.
//
// Uses the BASIC token (`graph.get`), not the elevated one `list-chat-members`
// uses: `/chats/{id}/members` needs `ChatMember.Read`, which the Teams web-client
// token grants (verified live), so basic suffices — and the elevated token can't
// refresh on the command path, so depending on it would make this fail whenever
// it is stale. If the scope is ever absent the call 403s, `null` is returned, and
// the chat is surfaced via `unresolvedMemberCount` / `hint` — graceful, not silent.
const hydrateMatches = async (graph: GraphClient, chatId: string, queryFolded: string): Promise<ReadonlyArray<Member> | null> => {
  const res = await graph.get(`/chats/${chatId}/members`);
  if (!res.ok) return null;
  const value = (res.value as GraphMembersResponse).value ?? [];
  return value.map(fromGraphMember).filter((m) => memberMatches(m, queryFolded));
};

// A direct (1:1) chat is the canonical target of a "find my chat with person X"
// search, and a cross-tenant counterpart there shows up bare — so these ALWAYS
// get hydrated, even when the cheap pass matched the same person elsewhere (the
// dual-identity case: resolved in a meeting, bare in the 1:1). Bare members in
// group/meeting chats are left to the `hint`.
//
// The substrate marks a 1:1 by the structural `@unq.gbl.spaces` id suffix
// (`19:<oidA>_<oidB>@unq.gbl.spaces`); group/meeting chats use `@thread.v2`.
// chatType is NOT a reliable discriminator here — the substrate reports a 1:1
// as `chatType: 'chat'` (not Graph's `'oneOnOne'`) and meetings as `'meeting'`,
// so the id suffix is the invariant we key on.
const isDirectChat = (chatId: string): boolean => chatId.endsWith('@unq.gbl.spaces');

type HydrationOutcome = { readonly chatsHydrated: number; readonly unresolvedMemberCount: number };

// Resolve every bare direct chat in parallel and fold any name matches into
// `matched`. A chat Graph cannot resolve (error) and every bare member left in
// a non-direct chat stay counted in `unresolvedMemberCount` — never silent.
const hydrateBareDirect = async (graph: GraphClient, queryFolded: string, bareUnmatched: ReadonlyArray<BareChat>, matched: Array<MatchedChat>): Promise<HydrationOutcome> => {
  const direct = bareUnmatched.filter((b) => isDirectChat(b.chatId));
  let unresolvedMemberCount = sumBareCounts(bareUnmatched.filter((b) => !isDirectChat(b.chatId)));
  const results = await Promise.all(direct.map(async (b) => ({ bare: b, hits: await hydrateMatches(graph, b.chatId, queryFolded) })));
  for (const { bare, hits } of results) {
    if (hits === null) {
      unresolvedMemberCount += bare.bareCount;
      continue;
    }
    if (hits.length > 0) matched.push(toMatchedChat(bare.chatId, bare.chat, hits));
  }
  return { chatsHydrated: direct.length, unresolvedMemberCount };
};

const HINT =
  'No chat member matched by name, but at least one chat has a cross-tenant member the Teams roster left unresolved — an externally-homed counterpart often appears only as a bare object-id. Direct 1:1 chats were deep-probed; members in group/meeting chats were not. Retry searching by their object-id (pass it as `--name <object-id>`), or, if you have the chat URL, read it directly with `get-chat` / `list-teams-chat-messages --chat-id 19:<their-oid>_<your-oid>@unq.gbl.spaces`.';

const execute: Command['execute'] = async (graph, params) => {
  const parsed = schema.safeParse(params);
  if (!parsed.success) return err({ type: 'validation_error', message: formatZodError(parsed.error) });
  const queryFolded = fold(parsed.data.name);
  const pageSize = parsed.data.pageSize ?? '100';
  const maxPages = Number(parsed.data.maxPages ?? '20');

  const scan = await scanChatPages(graph, queryFolded, pageSize, maxPages);
  if (!scan.ok) return scan;
  const { matched, bareUnmatched, pagesFetched, chatsScanned, continuationToken } = scan.value;

  // Always hydrate bare 1:1 chats: a dual-identity counterpart can be resolved
  // in a group/meeting yet bare in the direct chat, so this must run even when
  // the cheap pass already found a match elsewhere.
  const { chatsHydrated, unresolvedMemberCount } = await hydrateBareDirect(graph, queryFolded, bareUnmatched, matched);

  return ok({
    name: parsed.data.name,
    matches: matched,
    matchCount: matched.length,
    pagesFetched,
    chatsScanned,
    chatsHydrated,
    unresolvedMemberCount,
    hasMore: continuationToken !== undefined,
    nextContinuationToken: continuationToken,
    ...(matched.length === 0 && unresolvedMemberCount > 0 ? { hint: HINT } : {}),
  });
};

const meta: CommandMeta = {
  summary:
    'Find every Microsoft Teams chat that includes a member matching `--name` (substring search across display-name, email, given-name, surname, MRI, and object-id). Both sides are Unicode-folded (NFD + combining-mark strip) and lowercased before comparison, so `--name Alex` matches `Alex Kim` AND `alex.kim@example.com` AND `ALEX` — important because a dual-identity user often carries the accented display-name on one identity and the un-accented email on the other. Walks the paginated chat-list substrate up to `--max-pages` and returns matching chats with their `matchedMembers[]`. Collapses the canonical "all conversations with person X" workflow into a single call AND surfaces dual-identity people (e.g. someone with both an org MRI and a guest-tenant MRI). Cross-tenant resolution: the summary roster returns externally-homed counterparts as a bare object-id (no name/email), which a name search cannot match; for every bare DIRECT (1:1) chat the command hydrates the roster via the per-chat members endpoint and re-matches — so an external counterpart who is bare in your 1:1 is still found, even when they were already resolved in some meeting (the dual-identity case). Bare members in group/meeting chats are not deep-probed; when nothing matches and such members exist it returns a `hint` plus `unresolvedMemberCount` rather than a confident empty result. **Best-effort, may break on Microsoft client updates** — the chat substrate is not in the public Microsoft Graph API.',
  category: 'chats',
  needsSubstrateToken: 'chatsvcagg',
  graphMethod: 'GET',
  graphPathTemplate: 'https://teams.microsoft.com/api/csa/{region}/api/v3/teams/users/me/chats',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/chat-list',
  options: [
    {
      name: 'name',
      key: 'name',
      required: true,
      description:
        "Substring to search across each chat member's `displayName`, `email`, `userPrincipalName`, `givenName`, `surname`, `mri`, `objectId`, and `jobTitle`. Both the query and each field are NFD-normalized + diacritics-stripped + lowercased before comparison, so `Alex` ↔ `Alex` ↔ `ALEX` are equivalent and a query for the accented name still matches a member whose displayName is the un-accented email. Use the full name or an unambiguous fragment. Quoted multi-word values match on the joined substring, not per-token.",
    },
    {
      name: 'max-pages',
      key: 'maxPages',
      required: false,
      description:
        'Safety cap on the chat-list walk (positive integer; default 20). The walk stops early once the substrate reports no more pages, so the cap only bites on accounts with more chats than `--max-pages × --page-size`; raise it (and watch `hasMore`) if a known chat is missed. Each page is one HTTP round-trip.',
    },
    {
      name: 'page-size',
      key: 'pageSize',
      required: false,
      description: 'Chats per page (positive integer; default 100, same value Teams web uses). Server may silently cap.',
    },
  ],
  example: "ask-marcel-office find-chats-with-user --name 'Alex Kim'",
  responseShape:
    "`{ name, matches: [{ chatId, title, chatType, threadType, memberCount, lastMessageAt?, matchedMembers: [{ mri, displayName, email, userSubType }] }], matchCount, pagesFetched, chatsScanned, chatsHydrated, unresolvedMemberCount, hasMore, nextContinuationToken?, hint? }`. `matchedMembers` always carries the matching entries' identifying fields — pass `chatId` into `list-teams-chat-history` to read message bodies. `chatsHydrated` counts the per-chat members lookups spent resolving bare cross-tenant members in direct (1:1) chats. `unresolvedMemberCount` is how many cross-tenant members are still unresolved by name (bare members in group/meeting chats, which are not deep-probed, plus any 1:1 hydration that errored); when `matchCount` is 0 and this is non-zero, a `hint` is present explaining the likely cause and the object-id / read-by-chat-id remedy — so an empty result is never silently confident. `hasMore: true` means `--max-pages` was hit before exhausting the chat list; chain with the existing `--continuation-token` flag on `list-teams-chats-with-messages` if you need to scan further (this command does not advertise a `--continuation-token` because resuming a partial search is rare; users either widen `--max-pages` or refine `--name`).",
  stability: 'experimental',
};

export { execute, meta, schema };
