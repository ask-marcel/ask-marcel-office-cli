import type { Result } from '../../domain/result.ts';
import { ok } from '../../domain/result.ts';
import type { GraphClient, GraphError } from '../../infra/graph-client.ts';
import type { CommandOptionMeta } from './command-types.ts';

/**
 * What the Teams chat substrate leaves out of a message and an agent needs on
 * every read: a deep link, a normalised name for the system entries (call
 * started or ended, a recording or transcript posted, a member added or
 * removed, a topic change), and a way to keep only the messages that mention
 * the signed-in user. The csa (v1) route spells the fields in camelCase, the
 * IC3 route in lowercase; both are read here.
 */

type SubstrateMessage = Readonly<Record<string, unknown>>;
type SubstrateFilter = { readonly skipSystem: boolean; readonly mentionsOf?: string };

const CHAT_DEEP_LINK = 'https://teams.microsoft.com/l/message';

const typeOf = (m: SubstrateMessage): string => String(m['messageType'] ?? m['messagetype'] ?? '');

/** The event a system entry stands for, or `undefined` for a message somebody wrote. */
const eventOf = (m: SubstrateMessage): string | undefined => {
  const type = typeOf(m);
  if (type === 'Event/Call') return String(m['content'] ?? '').includes('<ended/>') ? 'call-ended' : 'call-started';
  if (type === 'RichText/Media_CallRecording') return 'recording-posted';
  if (type === 'RichText/Media_CallTranscript') return 'transcript-posted';
  if (!type.startsWith('ThreadActivity/')) return undefined;
  const kind = type.slice('ThreadActivity/'.length);
  if (kind === 'MemberJoined' || kind === 'AddMember') return 'member-added';
  if (kind === 'MemberLeft' || kind === 'DeleteMember') return 'member-removed';
  if (kind === 'TopicUpdate') return 'topic-changed';
  return `thread-activity:${kind}`;
};

/** The Teams deep link of a message, the shape the Teams client itself copies. */
const webUrlOf = (chatId: string, m: SubstrateMessage): string => `${CHAT_DEEP_LINK}/${encodeURIComponent(chatId)}/${encodeURIComponent(String(m['id'] ?? ''))}`;

/** The message with `webUrl` and, for a system entry, `event`. */
const enrichSubstrateMessage = (chatId: string, m: SubstrateMessage): SubstrateMessage => {
  const event = eventOf(m);
  return { ...m, webUrl: webUrlOf(chatId, m), ...(event === undefined ? {} : { event }) };
};

const mentionMris = (m: SubstrateMessage): ReadonlyArray<string> => {
  const props = m['properties'];
  if (props === null || typeof props !== 'object') return [];
  const raw = (props as Record<string, unknown>)['mentions'];
  let list: unknown = raw;
  if (typeof raw === 'string') {
    try {
      list = JSON.parse(raw) as unknown;
    } catch {
      return [];
    }
  }
  if (!Array.isArray(list)) return [];
  return list.map((x) => String((x as { mri?: unknown })?.mri ?? ''));
};

/** Whether the message @mentions the user whose directory object id is given (`8:orgid:<id>` in the substrate). */
const mentionsUser = (m: SubstrateMessage, userObjectId: string): boolean => mentionMris(m).some((mri) => mri.endsWith(`:${userObjectId}`));

/** Applies `--skip-system` and `--mentions-me` to a list of messages. */
const filterSubstrateMessages = (messages: ReadonlyArray<SubstrateMessage>, filter: SubstrateFilter): ReadonlyArray<SubstrateMessage> =>
  messages.filter((m) => (!filter.skipSystem || eventOf(m) === undefined) && (filter.mentionsOf === undefined || mentionsUser(m, filter.mentionsOf)));

export { enrichSubstrateMessage, eventOf, filterSubstrateMessages, mentionsUser, webUrlOf };
export type { SubstrateFilter, SubstrateMessage };

// --- the two flags the chat listings share --------------------------------

type SubstrateFilterParams = { readonly skipSystem?: 'true' | 'false'; readonly mentionsMe?: 'true' | 'false' };

const SUBSTRATE_FILTER_OPTIONS: ReadonlyArray<CommandOptionMeta> = [
  {
    name: 'skip-system',
    key: 'skipSystem',
    required: false,
    description:
      'Pass `--skip-system true` to drop the system entries (call started and ended, recording and transcript cards, members added or removed, topic changes) and keep only what people wrote. Every entry carries a normalised `event` field naming its kind, and the count dropped comes back as `omitted`.',
    argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
  },
  {
    name: 'mentions-me',
    key: 'mentionsMe',
    required: false,
    description:
      'Pass `--mentions-me true` to keep only the messages that @mention the signed-in user (one extra `/me` read resolves the identity). Combines with `--skip-system`.',
    argumentHint: { kind: 'magicValue', values: ['true', 'false'] },
  },
];

/** Resolves the two flags into a filter; `--mentions-me` costs one `/me` read for the caller's directory id. */
const substrateFilterFor = async (graph: GraphClient, params: SubstrateFilterParams): Promise<Result<SubstrateFilter, GraphError>> => {
  const skipSystem = params.skipSystem === 'true';
  if (params.mentionsMe !== 'true') return ok({ skipSystem });
  const me = await graph.get('/me?$select=id');
  if (!me.ok) return me;
  return ok({ skipSystem, mentionsOf: String((me.value as { id?: unknown }).id ?? '') });
};

export { SUBSTRATE_FILTER_OPTIONS, substrateFilterFor };
