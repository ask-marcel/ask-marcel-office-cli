import { parseJson } from '../../domain/json.ts';
import type { Result } from '../../domain/result.ts';
import { ok, unwrapOr } from '../../domain/result.ts';
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

type People = ReadonlyMap<string, string>;
type Entry = Readonly<Record<string, unknown>>;

const isEntry = (value: unknown): value is Entry => value !== null && typeof value === 'object' && !Array.isArray(value);

// `properties.files` comes as a JSON string and `properties.emotions` as an
// array (probed 2026-09-27); either shape is read, anything else is empty.
const propertyList = (m: SubstrateMessage, key: string): ReadonlyArray<Entry> => {
  const props = m['properties'];
  const raw = isEntry(props) ? props[key] : undefined;
  const value = typeof raw === 'string' ? unwrapOr(parseJson(raw), undefined) : raw;
  return Array.isArray(value) ? value.filter(isEntry) : [];
};

const text = (value: unknown): string | undefined => (typeof value === 'string' ? value : undefined);
const isFiniteNumber = (value: unknown): value is number => Number.isFinite(value);

/** The object id at the end of an `8:orgid:<id>` mri, or of a sender URL ending in one. */
const objectIdOf = (mri: string): string => mri.slice(mri.lastIndexOf(':') + 1);

/** Who wrote each message of a read: the names a reaction's bare mri is matched against. */
const substratePeople = (messages: ReadonlyArray<SubstrateMessage>): People => {
  const people = new Map<string, string>();
  for (const m of messages) {
    const from = text(m['from']);
    const name = text(m['imDisplayName'] ?? m['imdisplayname']);
    if (from !== undefined && name !== undefined) people.set(objectIdOf(from), name);
  }
  return people;
};

/** The files shared in a message, from `properties.files`: name, type, and the links to resolve. */
const filesOf = (m: SubstrateMessage): ReadonlyArray<Entry> =>
  propertyList(m, 'files').flatMap((f) => {
    const name = text(f['fileName']);
    if (name === undefined) return [];
    const info = isEntry(f['fileInfo']) ? f['fileInfo'] : {};
    return [
      Object.fromEntries(Object.entries({ name, type: text(f['fileType']), url: text(info['fileUrl']), shareUrl: text(info['shareUrl']) }).filter(([, v]) => v !== undefined)),
    ];
  });

/** Who reacted with what and when, oldest first, from `properties.emotions`. */
const reactionsOf = (m: SubstrateMessage, people: People): ReadonlyArray<{ readonly type: string; readonly by: string; readonly at: string }> =>
  propertyList(m, 'emotions')
    .flatMap((e) => {
      const type = text(e['key']);
      const users = Array.isArray(e['users']) ? e['users'].filter(isEntry) : [];
      return type === undefined ? [] : users.flatMap((u) => reactionOf(type, u, people));
    })
    .toSorted((a, b) => a.at.localeCompare(b.at));

const reactionOf = (type: string, user: Entry, people: People): ReadonlyArray<{ readonly type: string; readonly by: string; readonly at: string }> => {
  const mri = text(user['mri']);
  const time = user['time'];
  if (mri === undefined || !isFiniteNumber(time)) return [];
  const id = objectIdOf(mri);
  return [{ type, by: people.get(id) ?? `user ${id}`, at: new Date(time).toISOString() }];
};

/** The message with `webUrl`, `event` for a system entry, and `files` and `reactions` when it has any. */
const enrichSubstrateMessage = (chatId: string, m: SubstrateMessage, people: People = substratePeople([m])): SubstrateMessage => {
  const event = eventOf(m);
  const files = filesOf(m);
  const reactions = reactionsOf(m, people);
  return {
    ...m,
    webUrl: webUrlOf(chatId, m),
    ...(event === undefined ? {} : { event }),
    ...(files.length === 0 ? {} : { files }),
    ...(reactions.length === 0 ? {} : { reactions }),
  };
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

export { enrichSubstrateMessage, eventOf, filterSubstrateMessages, mentionsUser, substratePeople, webUrlOf };
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
