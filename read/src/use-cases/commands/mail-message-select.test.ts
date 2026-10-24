import { describe, expect, it } from 'bun:test';
import { MAIL_MESSAGE_DEFAULT_SELECT } from './mail-message-select.ts';

describe('the slim projection every mail read shares', () => {
  it('carries the fields an agent links and triages on, webLink included, as one comma-separated list without spaces', () => {
    const fields = MAIL_MESSAGE_DEFAULT_SELECT.split(',');
    expect(fields).toContain('id');
    expect(fields).toContain('webLink');
    expect(fields).toContain('conversationId');
    expect(MAIL_MESSAGE_DEFAULT_SELECT).not.toContain(' ');
    expect(new Set(fields).size).toBe(fields.length);
  });
});
