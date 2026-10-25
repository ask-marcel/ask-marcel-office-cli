import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { fakeReadGraph } from '@ask-marcel/office-core/test-helpers/read-graph-fake.ts';
import type { Command, ReadGraph } from './index.ts';
import { commands } from './index.ts';

// The public library of the read package. A consumer runs a command of
// `commands` on a `ReadGraph`, and names the command with `Command`. These
// exported types must fit each other: when `Command` asks for more graph than
// `ReadGraph` gives, the typecheck of the read package fails on this file.

const PROFILE = { id: 'u-1', displayName: 'Ada Lovelace' };

const profileGraph = (): ReadGraph => fakeReadGraph({ get: async () => ok(PROFILE) });

const runCommand = (command: Command, graph: ReadGraph): ReturnType<Command['execute']> => command.execute(graph, {});

const currentUser = (): Command => {
  const command = commands['get-current-user'];
  if (command === undefined) throw new Error('the registry has no get-current-user');
  return command;
};

describe('the read library', () => {
  it('runs get-current-user on a ReadGraph through the Command type that it exports', async () => {
    expect(await runCommand(currentUser(), profileGraph())).toEqual(ok(PROFILE));
  });
});
