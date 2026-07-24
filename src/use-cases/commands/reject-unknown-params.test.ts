import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeGraphClient } from '../../test-helpers/graph-client-fake.ts';
import { createFileSystemFake } from '../../test-helpers/filesystem-fake.ts';
import { commands } from './index.ts';

/*
 * The whole point of the 2026-07-24 unification: a parameter a command does not
 * declare fails the SAME way no matter which surface asked. Commander already
 * rejected unknown flags on the CLI, but the Zod schemas STRIP unknown keys, so
 * MCP `run-command` params and direct library `commands[x].execute(...)` calls
 * silently swallowed them and returned data that looked like it obeyed. The
 * 2026-07-24 audit found that on 5 of 7 delta commands alone.
 */
describe('a parameter the command does not declare is refused on every surface', () => {
  it('refuses an unknown key on a Graph-backed command instead of stripping it and calling Graph anyway', async () => {
    let called = false;
    const graph = fakeGraphClient({
      get: async () => {
        called = true;
        return ok({ value: [] });
      },
    });

    const result = await commands['list-todo-tasks-delta'].execute(graph, { todoTaskListId: 'L', filter: "status eq 'x'" });

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('validation_error');
      expect(result.error.message).toContain('--filter');
    }
    // Refused BEFORE the request: a silently-ignored filter is the bug.
    expect(called).toBe(false);
  });

  it('names the flags the command does accept, so the caller can retry without a second round-trip', async () => {
    const result = await commands['list-mail-folder-messages-delta'].execute(fakeGraphClient(), { mailFolderId: 'inbox', orderby: 'x' });

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.message).toContain('--select');
      expect(result.error.message).toContain('--mail-folder-id');
    }
  });

  it('carries a machine-readable errorCode so an agent can branch without substring-matching', async () => {
    const result = await commands['get-current-user'].execute(fakeGraphClient(), { nope: 'x' });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('unknown_parameter');
  });

  it('refuses an unknown key on a local-filesystem command too, which never reaches Graph', async () => {
    const fs = createFileSystemFake();
    fs.seed('/work/data.csv', 'name,age\nAlice,30');

    const command = commands['convert-local-file-to-markdown'];
    const result = await command.executeLocal?.(fs, { path: '/work/data.csv', bogus: 'x' });

    expect(result?.ok).toBe(false);
    if (result && !result.ok) expect(result.error.code).toBe('unknown_parameter');
  });

  it('still accepts every key the command declares, so the guard never blocks a legitimate call', async () => {
    const graph = fakeGraphClient({ get: async () => ok({ value: [] }) });

    const result = await commands['list-mail-folder-messages-delta'].execute(graph, { mailFolderId: 'inbox', select: 'id', filter: 'isRead eq false', top: '5' });

    expect(result.ok).toBe(true);
  });
});
