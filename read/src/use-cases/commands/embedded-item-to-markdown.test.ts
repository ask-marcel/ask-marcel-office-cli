import { describe, expect, it } from 'bun:test';
import { embeddedContactToMarkdown, embeddedEventToMarkdown, embeddedMessageToMarkdown } from './embedded-item-to-markdown.ts';

describe('embeddedMessageToMarkdown — render an itemAttachment of type message as headers + body markdown', () => {
  it('renders the standard headers (From, To, Subject, Sent) plus the body run through htmlToMarkdown', () => {
    const md = embeddedMessageToMarkdown({
      subject: 'Re: Q3 plan',
      from: { emailAddress: { name: 'Alice', address: 'alice@contoso.com' } },
      toRecipients: [{ emailAddress: { name: 'Bob', address: 'bob@contoso.com' } }],
      sentDateTime: '2026-04-30T15:00:00Z',
      body: { contentType: 'html', content: '<p>Looks <strong>good</strong>.</p>' },
    });
    expect(md).toContain('**Subject:** Re: Q3 plan');
    expect(md).toContain('**From:** Alice <alice@contoso.com>');
    expect(md).toContain('**To:** Bob <bob@contoso.com>');
    expect(md).toContain('**Sent:** 2026-04-30T15:00:00Z');
    expect(md).toContain('Looks **good**.');
  });

  it('handles multiple toRecipients and missing displayNames', () => {
    const md = embeddedMessageToMarkdown({
      subject: 'multi',
      toRecipients: [{ emailAddress: { address: 'a@x' } }, { emailAddress: { address: 'b@x' } }],
      body: { contentType: 'html', content: '<p>x</p>' },
    });
    expect(md).toContain('**To:** a@x, b@x');
  });

  it('renders text-content bodies as-is when contentType is `text` instead of `html`', () => {
    const md = embeddedMessageToMarkdown({
      subject: 'plain',
      body: { contentType: 'text', content: 'Hello\nworld' },
    });
    expect(md).toContain('Hello');
    expect(md).toContain('world');
  });

  it('still produces a header block when the message has no body', () => {
    const md = embeddedMessageToMarkdown({ subject: 'empty', from: { emailAddress: { address: 'x@y' } } });
    expect(md).toContain('**Subject:** empty');
    expect(md).toContain('**From:** x@y');
  });

  it('omits headers whose source field is missing entirely', () => {
    const md = embeddedMessageToMarkdown({ subject: 'only-subject' });
    expect(md).toContain('**Subject:** only-subject');
    expect(md).not.toContain('**From:**');
    expect(md).not.toContain('**To:**');
    expect(md).not.toContain('**Sent:**');
  });
});

describe('embeddedEventToMarkdown — render an itemAttachment of type event', () => {
  it('renders subject, start/end, location, organizer, attendees, and body', () => {
    const md = embeddedEventToMarkdown({
      subject: 'Quarterly Review',
      start: { dateTime: '2026-05-01T09:00:00', timeZone: 'UTC' },
      end: { dateTime: '2026-05-01T10:00:00', timeZone: 'UTC' },
      location: { displayName: 'Boardroom 4' },
      organizer: { emailAddress: { name: 'Alice', address: 'alice@contoso.com' } },
      attendees: [{ emailAddress: { address: 'bob@contoso.com' } }, { emailAddress: { address: 'carol@contoso.com' } }],
      body: { contentType: 'html', content: '<p>Agenda enclosed.</p>' },
    });
    expect(md).toContain('**Subject:** Quarterly Review');
    expect(md).toContain('**Start:** 2026-05-01T09:00:00 UTC');
    expect(md).toContain('**End:** 2026-05-01T10:00:00 UTC');
    expect(md).toContain('**Location:** Boardroom 4');
    expect(md).toContain('**Organizer:** Alice <alice@contoso.com>');
    expect(md).toContain('**Attendees:** bob@contoso.com, carol@contoso.com');
    expect(md).toContain('Agenda enclosed.');
  });

  it('omits the location header when the event has no location.displayName', () => {
    const md = embeddedEventToMarkdown({
      subject: 'no-loc',
      start: { dateTime: '2026-05-01T09:00:00', timeZone: 'UTC' },
      end: { dateTime: '2026-05-01T10:00:00', timeZone: 'UTC' },
    });
    expect(md).not.toContain('**Location:**');
  });

  it('omits attendees when the array is empty', () => {
    const md = embeddedEventToMarkdown({ subject: 'solo', attendees: [] });
    expect(md).not.toContain('**Attendees:**');
  });
});

describe('embeddedContactToMarkdown — render an itemAttachment of type contact', () => {
  it('renders displayName, emails, phones, and company info', () => {
    const md = embeddedContactToMarkdown({
      displayName: 'Alice Doe',
      emailAddresses: [{ address: 'alice@contoso.com' }, { address: 'alice@personal.com' }],
      businessPhones: ['+1-555-0100'],
      mobilePhone: '+1-555-0101',
      companyName: 'Contoso',
      jobTitle: 'VP Engineering',
    });
    expect(md).toContain('**Name:** Alice Doe');
    expect(md).toContain('**Emails:** alice@contoso.com, alice@personal.com');
    expect(md).toContain('**Business phone:** +1-555-0100');
    expect(md).toContain('**Mobile phone:** +1-555-0101');
    expect(md).toContain('**Company:** Contoso');
    expect(md).toContain('**Title:** VP Engineering');
  });

  it('omits fields that are missing or empty', () => {
    const md = embeddedContactToMarkdown({ displayName: 'Bob' });
    expect(md).toContain('**Name:** Bob');
    expect(md).not.toContain('**Emails:**');
    expect(md).not.toContain('**Business phone:**');
    expect(md).not.toContain('**Mobile phone:**');
    expect(md).not.toContain('**Company:**');
    expect(md).not.toContain('**Title:**');
  });
});
