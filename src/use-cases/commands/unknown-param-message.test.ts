import { describe, expect, it } from 'bun:test';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const messageFor = async (params: Record<string, string>): Promise<string> => {
  const command = commands['get-mail-message'];
  if (!command) throw new Error('get-mail-message is not registered');
  const result = await command.execute(fakeGraphClient(), params);
  if (result.ok) throw new Error('expected a rejection');
  return result.error.message;
};

describe('the rejection of a parameter the command does not declare', () => {
  it('names the flag, the flag it was likely meant to be, and the supported flags in order', async () => {
    expect(await messageFor({ messageId: 'm1', selct: 'id' })).toBe(
      '--selct is not a parameter of this command, so it would have been ignored rather than applied. Did you mean `--select`? Supported: --expand, --message-id, --select.'
    );
  });

  it('names each unknown flag with its own suggestion', async () => {
    expect(await messageFor({ messageId: 'm1', selct: 'id', expnd: 'x' })).toBe(
      '--selct, --expnd are not parameters of this command, so it would have been ignored rather than applied. For `--selct`, did you mean `--select`? For `--expnd`, did you mean `--expand`? Supported: --expand, --message-id, --select.'
    );
  });
});
