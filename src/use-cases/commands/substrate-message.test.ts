import { describe, expect, it } from 'bun:test';
import { enrichSubstrateMessage, eventOf, filterSubstrateMessages, mentionsUser, webUrlOf } from './substrate-message.ts';

const CHAT = '19:abc@thread.v2';
const post = (over: Record<string, unknown> = {}): Record<string, unknown> => ({ id: '1700000000000', messageType: 'RichText/Html', content: '<p>hello</p>', ...over });

describe('naming the system entries of a chat', () => {
  it('reads the call, recording, transcript and membership entries of the csa route, and the lowercase IC3 spelling', () => {
    expect(eventOf(post({ messageType: 'Event/Call', content: '<partlist></partlist><meetingDetails/>' }))).toBe('call-started');
    expect(eventOf(post({ messageType: 'Event/Call', content: '<ended/><partlist count="5"></partlist>' }))).toBe('call-ended');
    expect(eventOf(post({ messageType: 'RichText/Media_CallRecording' }))).toBe('recording-posted');
    expect(eventOf(post({ messageType: 'RichText/Media_CallTranscript' }))).toBe('transcript-posted');
    expect(eventOf(post({ messageType: 'ThreadActivity/MemberJoined' }))).toBe('member-added');
    expect(eventOf(post({ messageType: 'ThreadActivity/AddMember' }))).toBe('member-added');
    expect(eventOf(post({ messageType: 'ThreadActivity/MemberLeft' }))).toBe('member-removed');
    expect(eventOf(post({ messageType: 'ThreadActivity/DeleteMember' }))).toBe('member-removed');
    expect(eventOf(post({ messageType: 'ThreadActivity/TopicUpdate' }))).toBe('topic-changed');
    expect(eventOf(post({ messageType: 'ThreadActivity/MeetingPolicyUpdated' }))).toBe('thread-activity:MeetingPolicyUpdated');
    expect(eventOf({ id: '1', messagetype: 'ThreadActivity/MemberJoined' })).toBe('member-added');
  });

  it('leaves a written message, in either spelling, without an event', () => {
    expect(eventOf(post())).toBeUndefined();
    expect(eventOf(post({ messageType: 'Text' }))).toBeUndefined();
    expect(eventOf({ id: '1', messagetype: 'RichText/Html' })).toBeUndefined();
    expect(eventOf({ id: '1' })).toBeUndefined();
  });
});

describe('the deep link and the enriched message', () => {
  it('builds the Teams message link with the chat id encoded and adds it to the message, plus the event when there is one', () => {
    expect(webUrlOf(CHAT, post())).toBe('https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/1700000000000');
    expect(enrichSubstrateMessage(CHAT, post())).toEqual({ ...post(), webUrl: 'https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/1700000000000' });
    expect(enrichSubstrateMessage(CHAT, post({ messageType: 'ThreadActivity/TopicUpdate' }))).toMatchObject({ event: 'topic-changed' });
    expect(webUrlOf(CHAT, { messageType: 'Text' })).toBe('https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/');
  });
});

describe('who a message mentions', () => {
  const OID = '72f4fcd7-78d5-4e9c-b198-388b77646bd4';
  const mentions = JSON.stringify([{ itemid: 0, mri: `8:orgid:${OID}`, mentionType: 'person', displayName: 'Alex Kim' }]);

  it('reads the mentions JSON the substrate stores as a string, and as an array', () => {
    expect(mentionsUser(post({ properties: { mentions } }), OID)).toBe(true);
    expect(mentionsUser(post({ properties: { mentions: JSON.parse(mentions) as unknown } }), OID)).toBe(true);
    expect(mentionsUser(post({ properties: { mentions } }), 'someone-else')).toBe(false);
  });

  it('treats a message without properties, without mentions, with broken mentions JSON or with a non-list as mentioning nobody', () => {
    expect(mentionsUser(post(), OID)).toBe(false);
    expect(mentionsUser(post({ properties: null }), OID)).toBe(false);
    expect(mentionsUser(post({ properties: {} }), OID)).toBe(false);
    expect(mentionsUser(post({ properties: { mentions: '{not json' } }), OID)).toBe(false);
    expect(mentionsUser(post({ properties: { mentions: '{"mri":"8:orgid:x"}' } }), OID)).toBe(false);
    expect(mentionsUser(post({ properties: { mentions: [null, 3] } }), OID)).toBe(false);
  });

  it('filters by system entries, by mentions, by both, or not at all', () => {
    const system = post({ id: 's', messageType: 'Event/Call', content: '<ended/>' });
    const mine = post({ id: 'm', properties: { mentions } });
    const other = post({ id: 'o' });
    const all = [system, mine, other];
    expect(filterSubstrateMessages(all, { skipSystem: false })).toEqual(all);
    expect(filterSubstrateMessages(all, { skipSystem: true }).map((m) => m['id'])).toEqual(['m', 'o']);
    expect(filterSubstrateMessages(all, { skipSystem: false, mentionsOf: OID }).map((m) => m['id'])).toEqual(['m']);
    expect(
      filterSubstrateMessages([system, post({ id: 'sm', messageType: 'ThreadActivity/TopicUpdate', properties: { mentions } }), mine], { skipSystem: true, mentionsOf: OID }).map(
        (m) => m['id']
      )
    ).toEqual(['m']);
  });
});
