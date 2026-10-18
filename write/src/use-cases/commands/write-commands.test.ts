import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import type { WriteGraph } from '@ask-marcel/office-core/infra/write-graph.ts';
import { fakeWriteGraph } from '@ask-marcel/office-core/test-helpers/write-graph-fake.ts';
import { commands } from './write-commands.ts';

// The registry of the write package `@ask-marcel/office-write` (package split,
// phase 5): its seven commands, each wrapped once so that an unknown parameter
// is refused and the command runs on the write graph alone.

const WRITE_COMMANDS = [
  'convert-calendar-event-attachment-to-pdf',
  'convert-group-post-attachment-to-pdf',
  'convert-mail-attachment-to-pdf',
  'create-forward-draft',
  'create-mail-draft',
  'create-reply-draft',
  'update-mail-draft',
];

// A write graph that records each member that a command calls.
const recordingWriteGraph = (called: string[]): WriteGraph => {
  const answer = (member: string): (() => Promise<ReturnType<typeof ok<object>>>) => {
    const recorded = async (): Promise<ReturnType<typeof ok<object>>> => {
      called.push(member);
      return ok({});
    };
    return recorded;
  };
  return fakeWriteGraph(Object.fromEntries(['delete', 'fetchUrl', 'get', 'getBinary', 'patch', 'post', 'put'].map((member) => [member, answer(member)])));
};

describe('the registry of the write package', () => {
  it('holds the seven write commands, and no read command', () => {
    expect(Object.keys(commands).toSorted((a, b) => a.localeCompare(b))).toEqual(WRITE_COMMANDS);
    expect(Object.values(commands).filter((command) => command.meta.effect === 'read')).toEqual([]);
  });

  for (const name of WRITE_COMMANDS) {
    it(`refuses an unknown parameter of \`${name}\` before it calls Graph`, async () => {
      const called: string[] = [];
      const result = await commands[name]?.execute(recordingWriteGraph(called), { draftBodyy: 'x' });
      expect(result?.ok === false && result.error.code).toBe('unknown_parameter');
      expect(called).toEqual([]);
    });
  }

  it('runs a draft on the write graph alone', async () => {
    const called: string[] = [];
    const result = await commands['create-mail-draft']?.execute(recordingWriteGraph(called), { subject: 'Agenda', bodyContent: 'Hi', toRecipients: 'robin.chen@contoso.com' });
    expect(result?.ok).toBe(true);
    expect(called).toEqual(['post']);
  });
});
