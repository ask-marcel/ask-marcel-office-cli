import { describe, expect, it } from 'bun:test';
import { commands } from './index.ts';
import { resolveCommand } from './resolve-command.ts';

describe('resolving a command name an agent supplied', () => {
  it('finds a command asked for by its canonical registry name', () => {
    const result = resolveCommand(commands, 'list-mail-messages');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.name).toBe('list-mail-messages');
    expect(result.value.command.meta.category).toBe('mail');
  });

  it('finds a command asked for by a deprecated name it was renamed from, reporting the canonical name back', () => {
    // An agent that learned `download-onedrive-file-content` before the rename
    // still resolves; the CANONICAL name comes back so everything downstream
    // (docs rendering, the mutates gate) keys off one spelling.
    const result = resolveCommand(commands, 'download-onedrive-file-content');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.name).toBe('download-drive-item-content');
  });

  it('resolves every deprecated alias the registry advertises, so the manifest never names a spelling that fails to run', () => {
    const aliasPairs = Object.entries(commands).flatMap(([canonical, cmd]) => (cmd.meta.commandAliases ?? []).map((alias) => ({ canonical, alias })));
    expect(aliasPairs.length).toBeGreaterThan(0);
    for (const { canonical, alias } of aliasPairs) {
      const result = resolveCommand(commands, alias);
      expect(result.ok).toBe(true);
      if (!result.ok) return;
      expect(result.value.name).toBe(canonical);
    }
  });

  it('rejects a name that is neither a command nor an alias, listing the available names so the agent can retry without a second round-trip', () => {
    const result = resolveCommand(commands, 'list-mail-messagez');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.type).toBe('unknown_command');
    expect(result.error.name).toBe('list-mail-messagez');
    expect(result.error.available).toContain('list-mail-messages');
  });

  it('offers the available names in alphabetical order so the rejection reads predictably', () => {
    const result = resolveCommand(commands, 'nope');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.available).toEqual([...result.error.available].toSorted((a, b) => a.localeCompare(b)));
  });

  it('offers deprecated aliases among the available names, since an agent may legitimately invoke one', () => {
    const result = resolveCommand(commands, 'nope');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.available).toContain('download-onedrive-file-content');
  });

  it('prefers a canonical command over any alias that shadows it, so a rename can never hijack a live command name', () => {
    // Guards the lookup ORDER. `meta.test.ts` already forbids an alias that
    // collides with a canonical name, so this pins the resolution rule rather
    // than a live collision: canonical wins.
    const base = commands['get-current-user'];
    const registry = {
      'real-command': { ...base, meta: { ...base.meta, commandAliases: undefined } },
      decoy: { ...base, meta: { ...base.meta, commandAliases: ['real-command'] } },
    } as unknown as Parameters<typeof resolveCommand>[0];
    const result = resolveCommand(registry, 'real-command');
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.name).toBe('real-command');
  });

  it.each(['constructor', 'toString', 'hasOwnProperty', '__proto__'])(
    'rejects the inherited Object member `%s` instead of handing back a prototype value dressed up as a command',
    (inherited) => {
      // Without the `Object.hasOwn` gate, `registry['constructor']` resolves to
      // `Object` and this returns ok({ command: Object }) — a non-command that
      // would blow up at the first `.meta` read downstream.
      const result = resolveCommand(commands, inherited);
      expect(result.ok).toBe(false);
      if (result.ok) return;
      expect(result.error.type).toBe('unknown_command');
    }
  );

  it('advertises only names that actually resolve, so the rejection never sends an agent after a name that fails too', () => {
    const result = resolveCommand(commands, 'nope');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.available.length).toBeGreaterThan(0);
    for (const advertised of result.error.available) {
      expect(resolveCommand(commands, advertised).ok).toBe(true);
    }
  });

  it('rejects an empty command name rather than treating it as a lookup miss on a real command', () => {
    const result = resolveCommand(commands, '');
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.error.name).toBe('');
  });
});
