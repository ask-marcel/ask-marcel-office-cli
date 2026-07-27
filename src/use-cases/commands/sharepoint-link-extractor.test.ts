import { describe, expect, it } from 'bun:test';
import { buildShareToken, extractSharepointUrls } from './sharepoint-link-extractor.ts';

describe('extractSharepointUrls — surface every *.sharepoint.com link in an email body', () => {
  it('finds a single href to a SharePoint document', () => {
    const html = '<p>See <a href="https://contoso.sharepoint.com/sites/Marketing/Shared%20Documents/Q3.docx">the deck</a>.</p>';
    expect(extractSharepointUrls(html)).toEqual(['https://contoso.sharepoint.com/sites/Marketing/Shared%20Documents/Q3.docx']);
  });

  it('finds multiple distinct hrefs to different SharePoint files', () => {
    const html = '<a href="https://contoso.sharepoint.com/a.docx">a</a><br>' + '<a href="https://contoso.sharepoint.com/sites/team/b.xlsx">b</a>';
    const urls = extractSharepointUrls(html);
    expect(urls).toEqual(['https://contoso.sharepoint.com/a.docx', 'https://contoso.sharepoint.com/sites/team/b.xlsx']);
  });

  it('also catches *-my.sharepoint.com (personal sites)', () => {
    const html = '<a href="https://contoso-my.sharepoint.com/personal/alice/Documents/notes.docx">notes</a>';
    expect(extractSharepointUrls(html)).toEqual(['https://contoso-my.sharepoint.com/personal/alice/Documents/notes.docx']);
  });

  it('deduplicates identical URLs that appear multiple times', () => {
    const html = '<a href="https://contoso.sharepoint.com/a.docx">first</a>' + '<a href="https://contoso.sharepoint.com/a.docx">again</a>';
    expect(extractSharepointUrls(html)).toEqual(['https://contoso.sharepoint.com/a.docx']);
  });

  it('ignores non-SharePoint links (graph.microsoft.com, attacker.com, mailto, …)', () => {
    const html = '<a href="https://graph.microsoft.com/v1.0/me">graph</a>' + '<a href="https://attacker.example.com/x">bad</a>' + '<a href="mailto:alice@contoso.com">mail</a>';
    expect(extractSharepointUrls(html)).toEqual([]);
  });

  it('returns an empty array when the body is empty or has no links', () => {
    expect(extractSharepointUrls('')).toEqual([]);
    expect(extractSharepointUrls('<p>just text, no links</p>')).toEqual([]);
  });

  it('handles single-quoted href attribute', () => {
    const html = "<a href='https://contoso.sharepoint.com/a.docx'>x</a>";
    expect(extractSharepointUrls(html)).toEqual(['https://contoso.sharepoint.com/a.docx']);
  });

  it('strips a trailing fragment (#anchor) so the URL resolves cleanly via /shares/{token}', () => {
    const html = '<a href="https://contoso.sharepoint.com/a.docx#page=2">x</a>';
    expect(extractSharepointUrls(html)).toEqual(['https://contoso.sharepoint.com/a.docx']);
  });

  it('also surfaces SharePoint URLs that appear as bare text rather than as an href', () => {
    const html = '<p>The deck is at https://contoso.sharepoint.com/sites/team/deck.pptx and please review.</p>';
    expect(extractSharepointUrls(html)).toEqual(['https://contoso.sharepoint.com/sites/team/deck.pptx']);
  });
});

describe('buildShareToken — encode a SharePoint URL for the /shares/{token} resolver', () => {
  it('returns u! + base64url(url) (RFC 4648, no padding) per Graph spec', () => {
    // Spec: https://learn.microsoft.com/en-us/graph/api/shares-get
    // Algorithm: u! + base64(url) with - replacing + and _ replacing /, then strip = padding.
    const token = buildShareToken('https://contoso.sharepoint.com/sites/Marketing/Shared%20Documents/Q3.docx');
    expect(token.startsWith('u!')).toBe(true);
    expect(token).not.toContain('+');
    expect(token).not.toContain('/');
    expect(token).not.toContain('=');
  });

  it('round-trips a URL when manually decoded back', () => {
    const url = 'https://contoso.sharepoint.com/a.docx';
    const token = buildShareToken(url);
    // Strip 'u!', restore base64 alphabet, restore padding, decode.
    const stripped = token.slice(2).replaceAll('-', '+').replaceAll('_', '/');
    const padded = stripped + '='.repeat((4 - (stripped.length % 4)) % 4);
    expect(atob(padded)).toBe(url);
  });
});

describe('buildShareToken — non-ASCII URLs encode as UTF-8, not Latin-1', () => {
  // Independent oracle: base64url of the URL's UTF-8 bytes, via Buffer (not btoa).
  const referenceToken = (url: string): string => `u!${Buffer.from(url, 'utf-8').toString('base64').replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '')}`;
  // The old, buggy form: base64 of the Latin-1 code units. Only defined for chars <= U+00FF.
  const latin1Token = (url: string): string => `u!${btoa(url).replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '')}`;

  it('encodes a OneDrive path with accented characters (Réunião) as its UTF-8 bytes', () => {
    const url = 'https://contoso-my.sharepoint.com/personal/ana/Documents/Réunião.pdf';
    expect(buildShareToken(url)).toBe(referenceToken(url));
  });

  it('no longer produces the wrong Latin-1 token for an accented path (the bug)', () => {
    const url = 'https://contoso-my.sharepoint.com/personal/ana/Documents/Réunião.pdf';
    expect(buildShareToken(url)).not.toBe(latin1Token(url));
  });

  it('encodes a CJK filename (会議メモ) without throwing', () => {
    const url = 'https://contoso.sharepoint.com/sites/team/会議メモ.pdf';
    expect(buildShareToken(url)).toBe(referenceToken(url));
  });

  it('encodes an emoji filename (🚀) without throwing', () => {
    const url = 'https://contoso.sharepoint.com/sites/team/notes-🚀.pdf';
    expect(buildShareToken(url)).toBe(referenceToken(url));
  });

  it('round-trips a non-ASCII URL back to the original through the UTF-8 decoder', () => {
    const url = 'https://contoso.sharepoint.com/sites/team/会議メモ.pdf';
    const token = buildShareToken(url);
    const stripped = token.slice(2).replaceAll('-', '+').replaceAll('_', '/');
    const padded = stripped + '='.repeat((4 - (stripped.length % 4)) % 4);
    const binary = atob(padded);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    expect(new TextDecoder().decode(bytes)).toBe(url);
  });
});
