import { describe, expect, it } from 'bun:test';
import { findErrorHint } from './error-hints.ts';

describe('a message that vanished between the listing and the read', () => {
  it('learns that a move changes the id, and how to find the message again by its thread', () => {
    const hint = findErrorHint('ErrorItemNotFound: The specified object was not found in the store.', 'ErrorItemNotFound')?.hint ?? '';
    expect(hint).toContain('An Outlook message id also changes when the message moves to another folder');
    expect(hint).toContain('list-conversation-messages --conversation-id <its conversationId>');
  });
});
