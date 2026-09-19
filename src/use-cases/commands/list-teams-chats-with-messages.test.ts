import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

const command = commands['list-teams-chats-with-messages'];
if (!command) throw new Error('list-teams-chats-with-messages is not registered');

describe('the page size of the chat listing', () => {
  it('sends a whole positive page size and refuses anything else, naming the rule', async () => {
    const paths: string[] = [];
    const graph = fakeGraphClient({
      teamsChat: async (path) => {
        paths.push(path);
        return ok({ chats: [] });
      },
    });
    const fine = await command.execute(graph, { pageSize: '12' });
    expect(fine.ok).toBe(true);
    expect(paths[0]?.includes('pageSize=12')).toBe(true);
    for (const bad of ['0', '1a', 'a1', '12 ', '']) {
      const result = await command.execute(graph, { pageSize: bad });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.error.type).toBe('validation_error');
        if (bad !== '') expect(result.error.message).toContain('must be a positive integer');
      }
    }
    expect(paths).toHaveLength(1);
  });
});
