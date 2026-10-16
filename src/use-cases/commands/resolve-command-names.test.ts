import { describe, expect, it } from 'bun:test';
import { ok } from '../../domain/result.ts';
import { fakeCommand } from '../../test-helpers/command-package-fake.ts';
import { resolveCommand } from './resolve-command.ts';

// The name lookup the MCP gateway's tools share (core, package split phase 3).

describe('a command looked up by name', () => {
  const registry = { 'list-widgets': fakeCommand(), 'find-widgets': fakeCommand(), 'get-widget': fakeCommand() };

  it('is found by its exact name', () => {
    expect(resolveCommand(registry, 'get-widget')).toEqual(ok({ name: 'get-widget', command: registry['get-widget'] }));
  });

  it('is refused by an unknown name or a prototype key, with every name of the registry in alphabetical order', () => {
    for (const name of ['list-widget', 'constructor']) {
      expect(resolveCommand(registry, name)).toEqual({ ok: false, error: { type: 'unknown_command', name, available: ['find-widgets', 'get-widget', 'list-widgets'] } });
    }
  });
});
