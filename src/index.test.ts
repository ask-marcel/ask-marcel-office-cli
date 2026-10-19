import { describe, expect, it } from 'bun:test';
import { ok } from '@ask-marcel/office-core/domain/result.ts';
import { fakeGraphClient } from '@ask-marcel/office-core/test-helpers/graph-client-fake.ts';
import type { Command, CommandExecute, GraphClient } from './index.ts';
import { commands } from './index.ts';

// The public library of the root package. A consumer runs a command of
// `commands` on a `GraphClient`, and names the command with `Command` or its
// run with `CommandExecute`. These exported types must fit each other: when
// `Command` asks for more graph than `GraphClient` gives, the root typecheck
// fails on this file.

const PROFILE = { id: 'u-1', displayName: 'Ada Lovelace' };

const profileGraph = (): GraphClient => fakeGraphClient({ get: async () => ok(PROFILE) });

const runCommand = (command: Command, graph: GraphClient): ReturnType<CommandExecute> => command.execute(graph, {});

const runExecute = (execute: CommandExecute, graph: GraphClient): ReturnType<CommandExecute> => execute(graph, {});

const currentUser = (): Command => {
  const command = commands['get-current-user'];
  if (command === undefined) throw new Error('the registry has no get-current-user');
  return command;
};

describe('the library of the root package', () => {
  it('runs get-current-user on its own GraphClient through the Command type that it exports', async () => {
    expect(await runCommand(currentUser(), profileGraph())).toEqual(ok(PROFILE));
  });

  it('runs get-current-user on its own GraphClient through the CommandExecute type that it exports', async () => {
    expect(await runExecute(currentUser().execute, profileGraph())).toEqual(ok(PROFILE));
  });
});
