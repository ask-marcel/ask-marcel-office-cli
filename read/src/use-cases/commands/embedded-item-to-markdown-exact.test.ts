import { describe, expect, it } from 'bun:test';
import { err } from '@ask-marcel/office-core/domain/result.ts';
import { CONVERTERS } from '../../infra/converters.ts';
import type { Converters } from '../ports/converters.ts';
import { embeddedContactToMarkdown, embeddedEventToMarkdown, embeddedMessageToMarkdown } from './embedded-item-to-markdown.ts';

// Converters whose HTML conversion fails, so a body falls back to its raw content.
const FAILING: Converters = { ...CONVERTERS, htmlToMarkdown: () => err({ type: 'validation_error', message: 'bad html' }) };

describe('an embedded message, rendered exactly', () => {
  it('joins its headers by a line and its body by a blank line', () => {
    const md = embeddedMessageToMarkdown(CONVERTERS, {
      subject: 'Plan',
      toRecipients: [{ emailAddress: { name: 'Ana', address: 'ana@contoso.com' } }, { emailAddress: { name: 'No address' } }],
      body: { contentType: 'text', content: 'a_b *c*' },
    });
    expect(md).toBe('**Subject:** Plan\n**To:** Ana <ana@contoso.com>\n\na_b *c*');
  });

  it('gives no To header when no recipient has an address', () => {
    expect(embeddedMessageToMarkdown(CONVERTERS, { subject: 'Plan', toRecipients: [{ emailAddress: { name: 'No address' } }] })).toBe('**Subject:** Plan');
  });

  it('gives no To header for an empty recipient list', () => {
    expect(embeddedMessageToMarkdown(CONVERTERS, { subject: 'Plan', toRecipients: [] })).toBe('**Subject:** Plan');
  });

  it('gives the headers alone when the body has no content', () => {
    expect(embeddedMessageToMarkdown(CONVERTERS, { subject: 'Plan', body: { contentType: 'html' } })).toBe('**Subject:** Plan');
  });

  it('gives the body alone when there is no header', () => {
    expect(embeddedMessageToMarkdown(CONVERTERS, { body: { contentType: 'text', content: 'only body' } })).toBe('only body');
  });

  it('keeps the raw HTML body when the conversion fails', () => {
    expect(embeddedMessageToMarkdown(FAILING, { body: { contentType: 'html', content: '<p>raw</p>' } })).toBe('<p>raw</p>');
  });
});

describe('an embedded event, rendered exactly', () => {
  it('joins its headers by a line and its body by a blank line', () => {
    const md = embeddedEventToMarkdown(CONVERTERS, {
      subject: 'Review',
      location: { displayName: 'Room 1' },
      attendees: [{ emailAddress: { address: 'bo@contoso.com' } }, { emailAddress: {} }],
      body: { contentType: 'text', content: 'agenda' },
    });
    expect(md).toBe('**Subject:** Review\n**Location:** Room 1\n**Attendees:** bo@contoso.com\n\nagenda');
  });

  it('gives the body alone when there is no header', () => {
    expect(embeddedEventToMarkdown(CONVERTERS, { body: { contentType: 'text', content: 'agenda' } })).toBe('agenda');
  });

  it('gives the headers alone when there is no body', () => {
    expect(embeddedEventToMarkdown(CONVERTERS, { subject: 'Review' })).toBe('**Subject:** Review');
  });
});

describe('an embedded contact, rendered exactly', () => {
  it('lists only the addresses it has, and its fields one to a line', () => {
    const md = embeddedContactToMarkdown({
      displayName: 'Ana',
      emailAddresses: [{ address: 'ana@contoso.com' }, { address: '' }, {}, { address: 'a2@contoso.com' }],
      businessPhones: [],
      companyName: 'Contoso',
    });
    expect(md).toBe('**Name:** Ana\n**Emails:** ana@contoso.com, a2@contoso.com\n**Company:** Contoso');
  });

  it('gives the first business phone', () => {
    expect(embeddedContactToMarkdown({ businessPhones: ['+1 555', '+1 556'] })).toBe('**Business phone:** +1 555');
  });
});
