import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['list-teams-chat-history'];
if (!command) throw new Error('list-teams-chat-history is not registered');

const CHAT = '19:abc@thread.v2';
const OID = '72f4fcd7-78d5-4e9c-b198-388b77646bd4';
const mentions = JSON.stringify([{ itemid: 0, mri: `8:orgid:${OID}` }]);
const ic3Page = {
  messages: [
    { id: '3', messagetype: 'RichText/Html', content: '<p>later</p>', imdisplayname: 'Alex Kim', extra: 'dropped by the slim projection' },
    { id: '2', messagetype: 'ThreadActivity/MemberJoined', content: '{}' },
    { id: '1', messagetype: 'Text', content: 'ping', properties: { mentions } },
  ],
};

type Envelope = { readonly messages: ReadonlyArray<Record<string, unknown>>; readonly omitted?: number; readonly projection: string };

describe('the IC3 history read with links, events and filters', () => {
  it('projects webUrl and event into the slim shape', async () => {
    const graph = fakeGraphClient({ teamsChatIc3: async () => ok(ic3Page) });
    const result = await command.execute(graph, { chatId: CHAT });
    if (!result.ok) throw new Error('expected ok');
    const envelope = result.value as Envelope;
    expect(envelope.projection).toBe('slim');
    expect(envelope.messages[0]).toEqual({
      id: '3',
      messagetype: 'RichText/Html',
      imdisplayname: 'Alex Kim',
      content: '<p>later</p>',
      webUrl: 'https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/3',
    });
    expect(envelope.messages[1]).toMatchObject({ id: '2', event: 'member-added' });
    expect(envelope.omitted).toBeUndefined();
  });

  it('keeps the raw shape plus the two fields with --full true', async () => {
    const graph = fakeGraphClient({ teamsChatIc3: async () => ok(ic3Page) });
    const result = await command.execute(graph, { chatId: CHAT, full: 'true' });
    if (!result.ok) throw new Error('expected ok');
    expect((result.value as Envelope).messages[0]).toMatchObject({
      extra: 'dropped by the slim projection',
      webUrl: 'https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/3',
    });
  });

  it('applies --skip-system and --mentions-me to the accumulated history and counts what was dropped', async () => {
    const graph = fakeGraphClient({ teamsChatIc3: async () => ok(ic3Page), get: async () => ok({ id: OID }) });
    const skipped = await command.execute(graph, { chatId: CHAT, skipSystem: 'true' });
    if (!skipped.ok) throw new Error('expected ok');
    expect((skipped.value as Envelope).messages.map((m) => m['id'])).toEqual(['3', '1']);
    expect((skipped.value as Envelope).omitted).toBe(1);
    const mine = await command.execute(graph, { chatId: CHAT, mentionsMe: 'true' });
    if (!mine.ok) throw new Error('expected ok');
    expect((mine.value as Envelope).messages.map((m) => m['id'])).toEqual(['1']);
    expect((mine.value as Envelope).omitted).toBe(2);
  });

  it('advertises the two flags last', () => {
    const names = command.meta.options.map((o) => o.name);
    expect(names.slice(-2)).toEqual(['skip-system', 'mentions-me']);
  });
});

describe('the edges of the history read', () => {
  it('accepts the two flags spelled false, and refuses any other spelling', async () => {
    const graph = fakeGraphClient({ teamsChatIc3: async () => ok(ic3Page) });
    const off = await command.execute(graph, { chatId: CHAT, skipSystem: 'false', mentionsMe: 'false' });
    if (!off.ok) throw new Error('expected ok');
    expect((off.value as Envelope).messages).toHaveLength(3);
    const badFlags: ReadonlyArray<Record<string, string>> = [{ skipSystem: 'yes' }, { mentionsMe: 'yes' }];
    for (const bad of badFlags) {
      const result = await command.execute(graph, { chatId: CHAT, ...bad });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.type).toBe('validation_error');
    }
  });

  it('fails before reading the history when the identity behind --mentions-me cannot be resolved, and passes an IC3 failure through', async () => {
    const reads: number[] = [];
    const noMe = fakeGraphClient({
      get: async () => err({ type: 'api_error', status: 401, message: 'expired' }),
      teamsChatIc3: async () => {
        reads.push(1);
        return ok(ic3Page);
      },
    });
    const denied = await command.execute(noMe, { chatId: CHAT, mentionsMe: 'true' });
    expect(denied.ok).toBe(false);
    expect(reads).toEqual([]);
    const broken = fakeGraphClient({ teamsChatIc3: async () => err({ type: 'api_error', status: 503, message: 'down' }) });
    const failed = await command.execute(broken, { chatId: CHAT });
    expect(failed.ok).toBe(false);
    if (!failed.ok) expect(failed.error.message).toBe('down');
  });

  it('treats a page without a messages list as empty, and stops at an empty first page even when a syncState is offered', async () => {
    const calls: string[] = [];
    const graph = fakeGraphClient({
      teamsChatIc3: async (path) => {
        calls.push(path);
        return ok({ messages: [], _metadata: { syncState: 'https://teams.microsoft.com/api/chatsvc/emea/v1/users/ME/conversations/c/messages?syncState=next' } });
      },
    });
    const result = await command.execute(graph, { chatId: CHAT, maxPages: '5' });
    if (!result.ok) throw new Error('expected ok');
    expect((result.value as Envelope).messages).toEqual([]);
    expect(calls).toHaveLength(1);
    const bare = fakeGraphClient({ teamsChatIc3: async () => ok({}) });
    const empty = await command.execute(bare, { chatId: CHAT });
    if (!empty.ok) throw new Error('expected ok');
    expect((empty.value as Envelope).messages).toEqual([]);
    expect((empty.value as { hasMore: boolean }).hasMore).toBe(false);
  });

  it('keeps a body exactly at --max-content-chars untouched and cuts one character longer', async () => {
    const graph = fakeGraphClient({
      teamsChatIc3: async () =>
        ok({
          messages: [
            { id: 'a', messagetype: 'Text', content: 'abcde' },
            { id: 'b', messagetype: 'Text', content: 'abcdef' },
          ],
        }),
    });
    const result = await command.execute(graph, { chatId: CHAT, maxContentChars: '5' });
    if (!result.ok) throw new Error('expected ok');
    const [exact, longer] = (result.value as Envelope).messages;
    expect(exact).toMatchObject({ content: 'abcde' });
    expect(exact?.['truncated']).toBeUndefined();
    expect(longer).toMatchObject({ content: 'abcde', truncated: true, originalContentChars: 6 });
  });

  it('refuses page-size, max-pages and max-content-chars that are not whole positive numbers, naming the rule', async () => {
    const graph = fakeGraphClient({ teamsChatIc3: async () => ok(ic3Page) });
    const badNumbers: ReadonlyArray<Record<string, string>> = [
      { pageSize: '0' },
      { pageSize: '1a' },
      { pageSize: 'a1' },
      { maxPages: '2x' },
      { maxPages: 'x2' },
      { maxContentChars: '3.5' },
      { maxContentChars: ' 3' },
    ];
    for (const bad of badNumbers) {
      const result = await command.execute(graph, { chatId: CHAT, ...bad });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.error.type).toBe('validation_error');
        expect(result.error.message).toContain('must be a positive integer');
      }
    }
  });
});
