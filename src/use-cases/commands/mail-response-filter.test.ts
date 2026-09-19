import { describe, expect, it } from 'bun:test';
import { err, ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';
import { excludeMeetingResponses } from './mail-response-filter.ts';

const invite = { id: 'i', '@odata.type': '#microsoft.graph.eventMessageRequest', subject: 'Kick-off' };
const accepted = { id: 'a', '@odata.type': '#microsoft.graph.eventMessageResponse', subject: 'Accepted: Kick-off' };
const plain = { id: 'p', subject: 'Budget' };

describe('dropping meeting responses from a mail page', () => {
  it('keeps invites and plain mail, drops the responses, and counts them', () => {
    const result = excludeMeetingResponses(ok({ value: [invite, accepted, plain, accepted], '@odata.nextLink': 'n' }));
    if (!result.ok) throw new Error('expected ok');
    expect(result.value).toEqual({ value: [invite, plain], '@odata.nextLink': 'n', excludedMeetingResponses: 2 });
  });

  it('leaves a page with nothing to drop, a body without a value list, and an error untouched', () => {
    const clean = ok({ value: [invite, plain] });
    expect(excludeMeetingResponses(clean)).toEqual(ok({ value: [invite, plain] }));
    const single = ok({ id: 'x' });
    expect(excludeMeetingResponses(single)).toBe(single);
    const failure = err({ type: 'api_error' as const, status: 500, message: 'boom' });
    expect(excludeMeetingResponses(failure)).toBe(failure);
    expect(excludeMeetingResponses(ok({ value: [null, 3, 'x'] }))).toEqual(ok({ value: [null, 3, 'x'] }));
  });
});

describe('the flag on the two mail listings', () => {
  for (const [name, params, expectedPath] of [
    [
      'list-mail-messages',
      {},
      '/me/messages?$select=id%2Csubject%2Cfrom%2CtoRecipients%2CccRecipients%2CreceivedDateTime%2ChasAttachments%2CisRead%2Cimportance%2CbodyPreview%2CconversationId%2CwebLink',
    ],
    ['list-mail-folder-messages', { mailFolderId: 'f1' }, '/me/mailFolders/f1/messages'],
  ] as const) {
    it(`${name} drops the responses only when asked, and never sends the flag to Graph`, async () => {
      const command = commands[name];
      if (!command) throw new Error(`${name} is not registered`);
      const paths: string[] = [];
      const graph = fakeGraphClient({
        get: async (path) => {
          paths.push(path);
          return ok({ value: [invite, accepted, plain] });
        },
      });
      const on = await command.execute(graph, { ...params, excludeMeetingResponses: 'true' });
      if (!on.ok) throw new Error('expected ok');
      expect(on.value).toEqual({ value: [invite, plain], excludedMeetingResponses: 1 });
      const off = await command.execute(graph, { ...params, excludeMeetingResponses: 'false' });
      if (!off.ok) throw new Error('expected ok');
      expect(off.value).toEqual({ value: [invite, accepted, plain] });
      const absent = await command.execute(graph, params);
      if (!absent.ok) throw new Error('expected ok');
      expect(absent.value).toEqual({ value: [invite, accepted, plain] });
      expect(paths).toEqual([expectedPath, expectedPath, expectedPath]);
      expect(command.meta.options.map((o) => o.name)).toContain('exclude-meeting-responses');
    });
  }
});

describe('a flag value that is neither true nor false', () => {
  for (const name of ['list-mail-messages', 'list-mail-folder-messages'] as const) {
    it(`${name} refuses it before calling Graph`, async () => {
      const command = commands[name];
      if (!command) throw new Error(`${name} is not registered`);
      const paths: string[] = [];
      const graph = fakeGraphClient({
        get: async (path) => {
          paths.push(path);
          return ok({ value: [] });
        },
      });
      const result = await command.execute(graph, { mailFolderId: 'f1', excludeMeetingResponses: 'maybe' });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error.type).toBe('validation_error');
      expect(paths).toEqual([]);
    });
  }
});
