import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const CHAT = '19:abc@thread.v2';
const ROBIN = '11111111-2222-3333-4444-555555555555';
const ALEX = '66666666-7777-8888-9999-000000000000';
const FILE_URL = 'https://contoso-my.sharepoint.com/personal/robin_chen_contoso_com/Documents/Microsoft%20Teams%20Chat%20Files/order-export.xlsx';
const SHARE_URL = 'https://contoso-my.sharepoint.com/:x:/g/personal/robin_chen_contoso_com/EabcDEF';

// The substrate's own shapes (probed 2026-09-27): `properties.files` is a JSON
// string, `properties.emotions` an array; a reactor is an `8:orgid:<id>` mri.
const MESSAGES = [
  {
    id: '2',
    messageType: 'RichText/Html',
    from: `https://notifications.skype.net/v1/users/ME/contacts/8:orgid:${ALEX}`,
    imDisplayName: 'Alex Kim',
    content: '<p>thanks</p>',
  },
  {
    id: '1',
    messageType: 'RichText/Html',
    from: `https://notifications.skype.net/v1/users/ME/contacts/8:orgid:${ROBIN}`,
    imDisplayName: 'Robin Chen',
    content: '<p>the export</p>',
    properties: {
      files: JSON.stringify([
        {
          fileName: 'order-export.xlsx',
          fileType: 'xlsx',
          itemid: 'guid-1',
          fileInfo: { fileUrl: FILE_URL, shareUrl: SHARE_URL, siteUrl: 'https://contoso-my.sharepoint.com/personal/robin_chen_contoso_com/' },
        },
        { fileType: 'png' },
        { fileName: 42 },
        null,
        { fileName: 'notes.txt' },
      ]),
      emotions: [
        { key: 'heart', users: [{ mri: '8:orgid:99999999-8888-7777-6666-555555555555', time: 1_790_500_000_000, value: '1' }] },
        { key: 'like', users: [{ mri: `8:orgid:${ALEX}`, time: 1_790_400_000_000, value: '1' }, { mri: `8:orgid:${ROBIN}`, time: 'soon' }, null, { time: 1_790_300_000_000 }] },
        { key: 'laugh', users: [{ mri: `8:orgid:${ROBIN}`, time: 1_790_450_000_000 }] },
        { users: [{ mri: `8:orgid:${ALEX}`, time: 1_790_460_000_000 }] },
      ],
    },
  },
  { id: '0', messageType: 'RichText/Html', content: '<p>older</p>', properties: { files: '[]', emotions: [] } },
  { id: '-1', messageType: 'RichText/Html', content: '<p>oldest</p>', properties: null },
];

describe('files and reactions on chat messages', () => {
  it('parses the files shared in a message and who reacted when, oldest first, naming people from the same read', async () => {
    const command = commands['list-teams-chat-messages'];
    if (!command) throw new Error('list-teams-chat-messages is not registered');
    const result = await command.execute(fakeGraphClient({ teamsChat: async () => ok({ messages: MESSAGES }) }), { chatId: CHAT });
    if (!result.ok) throw new Error(result.error.message);
    const [alexMessage, robinMessage, older] = (result.value as { messages: ReadonlyArray<Record<string, unknown>> }).messages;
    expect(robinMessage?.['files']).toStrictEqual([{ name: 'order-export.xlsx', type: 'xlsx', url: FILE_URL, shareUrl: SHARE_URL }, { name: 'notes.txt' }]);
    expect(robinMessage?.['reactions']).toEqual([
      { type: 'like', by: 'Alex Kim', at: '2026-09-26T05:20:00.000Z' },
      { type: 'laugh', by: 'Robin Chen', at: '2026-09-26T19:13:20.000Z' },
      { type: 'heart', by: 'user 99999999-8888-7777-6666-555555555555', at: '2026-09-27T09:06:40.000Z' },
    ]);
    expect(robinMessage).not.toHaveProperty('event');
    expect(alexMessage).not.toHaveProperty('files');
    expect(alexMessage).not.toHaveProperty('reactions');
    expect(older).not.toHaveProperty('files');
    expect(older).not.toHaveProperty('reactions');
  });

  it('reads a single message the same way, and tolerates a files property that is not JSON', async () => {
    const command = commands['get-teams-chat-message'];
    if (!command) throw new Error('get-teams-chat-message is not registered');
    const one = await command.execute(fakeGraphClient({ teamsChat: async () => ok(MESSAGES[1]) }), { chatId: CHAT, messageId: '1' });
    if (!one.ok) throw new Error(one.error.message);
    expect((one.value as { reactions: ReadonlyArray<{ by: string }> }).reactions.map((r) => r.by)).toEqual([
      'user 66666666-7777-8888-9999-000000000000',
      'Robin Chen',
      'user 99999999-8888-7777-6666-555555555555',
    ]);
    const broken = await command.execute(fakeGraphClient({ teamsChat: async () => ok({ id: '5', properties: { files: '{not json', emotions: 'nope' } }) }), {
      chatId: CHAT,
      messageId: '5',
    });
    if (!broken.ok) throw new Error(broken.error.message);
    expect(broken.value).not.toHaveProperty('files');
    expect(broken.value).not.toHaveProperty('reactions');
  });
});

describe('files and reactions in the deep history read', () => {
  it('survive the slim projection, which still drops the raw properties', async () => {
    const command = commands['list-teams-chat-history'];
    if (!command) throw new Error('list-teams-chat-history is not registered');
    const ic3 = { ...MESSAGES[1], messagetype: 'RichText/Html', imdisplayname: 'Robin Chen', imDisplayName: undefined };
    const byAlex = { ...MESSAGES[0], messagetype: 'RichText/Html', imdisplayname: 'Alex Kim', imDisplayName: undefined };
    const result = await command.execute(fakeGraphClient({ teamsChatIc3: async () => ok({ messages: [ic3, byAlex] }) }), { chatId: CHAT });
    if (!result.ok) throw new Error(result.error.message);
    const [message] = (result.value as { messages: ReadonlyArray<Record<string, unknown>> }).messages;
    expect(message?.['files']).toEqual([{ name: 'order-export.xlsx', type: 'xlsx', url: FILE_URL, shareUrl: SHARE_URL }, { name: 'notes.txt' }]);
    expect((message?.['reactions'] as ReadonlyArray<{ by: string }>).map((r) => r.by)).toEqual(['Alex Kim', 'Robin Chen', 'user 99999999-8888-7777-6666-555555555555']);
    expect(message).not.toHaveProperty('properties');
  });
});
