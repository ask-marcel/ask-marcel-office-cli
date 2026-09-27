import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { commands } from './index.ts';

// Every call is recorded: an instant that no calendar holds must be refused
// before any request is built, never crash and never roll over to a real day.
const recordingGraph = (calls: string[]): ReturnType<typeof fakeGraphClient> =>
  fakeGraphClient({
    get: async (path) => {
      calls.push(path);
      return ok({ value: [] });
    },
    post: async (path) => {
      calls.push(path);
      return ok({ value: [] });
    },
    teamsChatIc3: async (path) => {
      calls.push(path);
      return ok({ messages: [] });
    },
  });

const run = async (name: string, params: Record<string, string>, calls: string[]): Promise<{ ok: boolean; type?: string; message?: string }> => {
  const command = commands[name];
  if (!command) throw new Error(`${name} is not registered`);
  const result = await command.execute(recordingGraph(calls), params);
  return result.ok ? { ok: true } : { ok: false, type: result.error.type, message: result.error.message };
};

describe('an instant that no calendar holds', () => {
  it('is refused as a validation error on --due-before instead of crashing the task listing', async () => {
    const calls: string[] = [];
    const refused = await run('list-incomplete-todo-tasks', { todoTaskListId: 'L1', dueBefore: '2026-13-01T00:00:00Z' }, calls);
    expect(refused).toMatchObject({ ok: false, type: 'validation_error' });
    expect(refused.message).toContain('2026-13-01T00:00:00Z');
    expect(calls).toEqual([]);
  });

  it('is refused on list-changed-files --since rather than rolled over from 30 February to 2 March', async () => {
    const calls: string[] = [];
    expect(await run('list-changed-files', { since: '2026-02-30T00:00:00Z' }, calls)).toMatchObject({ ok: false, type: 'validation_error' });
    expect(calls).toEqual([]);
  });

  it('is refused on the chat history --since, so the substrate never receives a NaN start time', async () => {
    const calls: string[] = [];
    expect(await run('list-teams-chat-history', { chatId: '19:abc@thread.v2', since: '2026-04-31T08:00:00Z' }, calls)).toMatchObject({
      ok: false,
      type: 'validation_error',
    });
    expect(calls).toEqual([]);
  });

  it('leaves a real instant untouched, fractions and the last second of a leap day included', async () => {
    const calls: string[] = [];
    expect(await run('list-incomplete-todo-tasks', { todoTaskListId: 'L1', dueBefore: '2028-02-29T23:59:59.5Z' }, calls)).toEqual({ ok: true });
    expect(decodeURIComponent(calls[0] ?? '')).toContain("dueDateTime/dateTime lt '2028-02-29T23:59:59'");
  });
});
