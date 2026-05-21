import { describe, expect, it } from 'bun:test';
import { commands } from './index.ts';
import { GRAPH_SCOPES_BY_COMMAND, lookupScopes } from './graph-scopes.ts';

// Commands that intentionally have no Graph scope mapping (meta / cursor /
// search-passthrough). They don't call a single fixed endpoint, or their
// scope requirement is inherited from elsewhere.
const COMMANDS_WITHOUT_SCOPES: ReadonlySet<string> = new Set([
  'next-page', // inherits from cursor target
  'microsoft-search-query', // varies by entity types
  'scopes-check', // no Graph call
  // chatsvcagg-tier commands — not Graph; auth gates server-side on the
  // captured Teams client identity, not on delegated Graph scopes.
  'list-teams-chats-with-messages',
  'list-teams-chat-messages',
  'list-teams-chat-history',
  'get-teams-chat-message',
]);

describe('graph-scopes — central scope map', () => {
  it('has an entry for every registered command except the meta/cursor exceptions (audit round-8 Wave C)', () => {
    const missing: string[] = [];
    for (const name of Object.keys(commands)) {
      if (COMMANDS_WITHOUT_SCOPES.has(name)) continue;
      const scopes = lookupScopes(name);
      if (scopes === undefined || scopes.length === 0) missing.push(name);
    }
    expect(missing).toEqual([]);
  });

  it('returns undefined for commands not in the registry (lookupScopes is total against the registry only)', () => {
    expect(lookupScopes('nonexistent-command')).toBeUndefined();
  });

  it('every scope listed is a non-empty string in PascalCase.Verb form', () => {
    const scopePattern = /^[A-Z][A-Za-z]+(\.[A-Z][A-Za-z]+)*$/;
    for (const [name, scopes] of Object.entries(GRAPH_SCOPES_BY_COMMAND)) {
      for (const scope of scopes) {
        if (!scopePattern.test(scope)) {
          throw new Error(`Command "${name}" has malformed scope "${scope}" — expected PascalCase.Verb (e.g. "Files.Read", "Chat.ReadBasic")`);
        }
      }
    }
  });
});
