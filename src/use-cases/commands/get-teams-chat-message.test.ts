import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['get-teams-chat-message'];
if (!command) throw new Error('get-teams-chat-message is not registered');

describe('reading one chat message through the substrate', () => {
  it('adds the deep link, and the event for a system entry', async () => {
    const graph = fakeGraphClient({ teamsChat: async () => ok({ id: '7', messageType: 'RichText/Media_CallRecording', content: 'x' }) });
    const result = await command.execute(graph, { chatId: '19:abc@thread.v2', messageId: '7' });
    if (!result.ok) throw new Error('expected ok');
    expect(result.value).toEqual({
      id: '7',
      messageType: 'RichText/Media_CallRecording',
      content: 'x',
      webUrl: 'https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/7',
      event: 'recording-posted',
    });
  });

  it('hands a substrate error back unchanged', async () => {
    const graph = fakeGraphClient({ teamsChat: async () => err({ type: 'api_error', status: 404, message: 'gone' }) });
    const result = await command.execute(graph, { chatId: '19:abc@thread.v2', messageId: '7' });
    expect(result.ok).toBe(false);
  });
});
