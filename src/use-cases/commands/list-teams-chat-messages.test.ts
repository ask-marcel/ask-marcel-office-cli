import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['list-teams-chat-messages'];
if (!command) throw new Error('list-teams-chat-messages is not registered');

const CHAT = '19:abc@thread.v2';
const OID = '72f4fcd7-78d5-4e9c-b198-388b77646bd4';
const mentions = JSON.stringify([{ itemid: 0, mri: `8:orgid:${OID}` }]);
const page = {
  messages: [
    { id: '3', messageType: 'RichText/Html', content: '<p>later</p>' },
    { id: '2', messageType: 'Event/Call', content: '<ended/>' },
    { id: '1', messageType: 'RichText/Html', content: '<p>ping</p>', properties: { mentions } },
  ],
  messageToken: 't',
};

const graphWith = (overrides: Partial<GraphClient> = {}): { graph: GraphClient; gets: string[] } => {
  const gets: string[] = [];
  const graph = fakeGraphClient({
    teamsChat: async () => ok(page),
    get: async (path) => {
      gets.push(path);
      return ok({ id: OID });
    },
    ...overrides,
  });
  return { graph, gets };
};

type Envelope = { readonly messages: ReadonlyArray<Record<string, unknown>>; readonly messageToken?: string; readonly omitted?: number };

describe('reading a chat through the substrate with links and events', () => {
  it('adds a deep link to every message and an event name to the system entries, keeping the envelope', async () => {
    const { graph, gets } = graphWith();
    const result = await command.execute(graph, { chatId: CHAT });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const envelope = result.value as Envelope;
    expect(envelope.messageToken).toBe('t');
    expect(envelope.omitted).toBeUndefined();
    expect(envelope.messages.map((m) => m['webUrl'])).toEqual([
      'https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/3',
      'https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/2',
      'https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/1',
    ]);
    expect(envelope.messages.map((m) => m['event'])).toEqual([undefined, 'call-ended', undefined]);
    expect(gets).toEqual([]);
  });

  it('drops the system entries with --skip-system true and counts them', async () => {
    const { graph } = graphWith();
    const result = await command.execute(graph, { chatId: CHAT, skipSystem: 'true' });
    if (!result.ok) throw new Error('expected ok');
    const envelope = result.value as Envelope;
    expect(envelope.messages.map((m) => m['id'])).toEqual(['3', '1']);
    expect(envelope.omitted).toBe(1);
  });

  it('keeps only the messages that mention the signed-in user with --mentions-me true, resolving the identity once', async () => {
    const { graph, gets } = graphWith();
    const result = await command.execute(graph, { chatId: CHAT, mentionsMe: 'true', skipSystem: 'false' });
    if (!result.ok) throw new Error('expected ok');
    const envelope = result.value as Envelope;
    expect(envelope.messages.map((m) => m['id'])).toEqual(['1']);
    expect(envelope.omitted).toBe(2);
    expect(gets).toEqual(['/me?$select=id']);
  });

  it('fails before the chat read when the identity cannot be resolved for --mentions-me', async () => {
    const chatReads: number[] = [];
    const { graph } = graphWith({
      get: async () => err({ type: 'api_error', status: 401, message: 'expired' }),
      teamsChat: async () => {
        chatReads.push(1);
        return ok(page);
      },
    });
    const result = await command.execute(graph, { chatId: CHAT, mentionsMe: 'true' });
    expect(result.ok).toBe(false);
    expect(chatReads).toEqual([]);
  });

  it('hands a substrate error back unchanged and tolerates a page without messages', async () => {
    const failing = graphWith({ teamsChat: async () => err({ type: 'api_error', status: 503, message: 'down' }) });
    const result = await command.execute(failing.graph, { chatId: CHAT });
    expect(result.ok).toBe(false);
    const empty = graphWith({ teamsChat: async () => ok({ messageToken: 'x' }) });
    const emptyResult = await command.execute(empty.graph, { chatId: CHAT });
    if (!emptyResult.ok) throw new Error('expected ok');
    expect((emptyResult.value as Envelope).messages).toEqual([]);
  });

  it('advertises the two flags after the chat id', () => {
    expect(command.meta.options.map((o) => o.name)).toEqual(['chat-id', 'skip-system', 'mentions-me']);
  });
});
