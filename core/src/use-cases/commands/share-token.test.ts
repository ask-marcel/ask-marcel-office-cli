import { describe, expect, it } from 'bun:test';
import { buildShareToken } from './share-token.ts';

// The share token read's link resolvers and write's PDF converters both mint
// (core, package split phase 3). The cases are copies of the buildShareToken
// cases in sharepoint-link-extractor.test.ts and commands.test.ts.

describe('buildShareToken — encode a SharePoint URL for the /shares/{token} resolver', () => {
  it('encodes a url as u! + unpadded base64url (+ → -, / → _, = stripped)', () => {
    expect(buildShareToken('https://contoso.sharepoint.com/x?q=~//?>')).toBe('u!aHR0cHM6Ly9jb250b3NvLnNoYXJlcG9pbnQuY29tL3g_cT1-Ly8_Pg');
  });

  it('round-trips a URL when manually decoded back', () => {
    const url = 'https://contoso.sharepoint.com/a.docx';
    const token = buildShareToken(url);
    const stripped = token.slice(2).replaceAll('-', '+').replaceAll('_', '/');
    const padded = stripped + '='.repeat((4 - (stripped.length % 4)) % 4);
    expect(atob(padded)).toBe(url);
  });
});

describe('buildShareToken — non-ASCII URLs encode as UTF-8, not Latin-1', () => {
  // Independent oracle: base64url of the URL's UTF-8 bytes, via Buffer (not btoa).
  const referenceToken = (url: string): string => `u!${Buffer.from(url, 'utf-8').toString('base64').replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '')}`;

  it('encodes a OneDrive path with accented characters (Réunião) as its UTF-8 bytes', () => {
    const url = 'https://contoso-my.sharepoint.com/personal/ana/Documents/Réunião.pdf';
    expect(buildShareToken(url)).toBe(referenceToken(url));
  });

  it('encodes a CJK filename (会議メモ) and an emoji filename (🚀) without throwing', () => {
    for (const url of ['https://contoso.sharepoint.com/sites/team/会議メモ.pdf', 'https://contoso.sharepoint.com/sites/team/notes-🚀.pdf']) {
      expect(buildShareToken(url)).toBe(referenceToken(url));
    }
  });
});
