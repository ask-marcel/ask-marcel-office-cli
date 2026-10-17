import { describe, expect, it } from 'bun:test';
import { decodeUtf8Text, isPlainTextFilename } from './text-passthrough.ts';

describe('decodeUtf8Text — content-sniff: valid UTF-8 returns the text, binary returns undefined', () => {
  it('returns the decoded string for valid UTF-8 bytes, including multi-byte characters', () => {
    expect(decodeUtf8Text(new TextEncoder().encode('hello — café 文'))).toBe('hello — café 文');
  });

  it('returns an empty string for empty input (a zero-length document is valid UTF-8)', () => {
    expect(decodeUtf8Text(new Uint8Array([]))).toBe('');
  });

  it('returns undefined for bytes that are not valid UTF-8 (binary)', () => {
    expect(decodeUtf8Text(new Uint8Array([0xff, 0xfe, 0xfd]))).toBeUndefined();
    expect(decodeUtf8Text(new Uint8Array([0x80]))).toBeUndefined();
    expect(decodeUtf8Text(new Uint8Array([0xc3, 0x28]))).toBeUndefined();
  });
});

describe('isPlainTextFilename — Graph cannot convert these formats; the conversion commands must short-circuit and return raw bytes', () => {
  it('recognises the plain-text extensions Graph will reject', () => {
    expect(isPlainTextFilename('notes.txt')).toBe(true);
    expect(isPlainTextFilename('README.md')).toBe(true);
    expect(isPlainTextFilename('CHANGELOG.markdown')).toBe(true);
    expect(isPlainTextFilename('index.html')).toBe(true);
    expect(isPlainTextFilename('legacy.htm')).toBe(true);
    expect(isPlainTextFilename('payload.json')).toBe(true);
    expect(isPlainTextFilename('feed.xml')).toBe(true);
    expect(isPlainTextFilename('server.log')).toBe(true);
    expect(isPlainTextFilename('config.yaml')).toBe(true);
    expect(isPlainTextFilename('config.yml')).toBe(true);
    expect(isPlainTextFilename('site.css')).toBe(true);
    expect(isPlainTextFilename('main.js')).toBe(true);
    expect(isPlainTextFilename('client.ts')).toBe(true);
    expect(isPlainTextFilename('deploy.sh')).toBe(true);
    expect(isPlainTextFilename('schema.sql')).toBe(true);
  });

  it('treats Office source formats as convertible (returns false)', () => {
    expect(isPlainTextFilename('q3-budget.docx')).toBe(false);
    expect(isPlainTextFilename('forecast.xlsx')).toBe(false);
    expect(isPlainTextFilename('deck.pptx')).toBe(false);
    expect(isPlainTextFilename('legacy.doc')).toBe(false);
    expect(isPlainTextFilename('table.csv')).toBe(false);
    expect(isPlainTextFilename('outline.rtf')).toBe(false);
  });

  it('matches case-insensitively on the extension', () => {
    expect(isPlainTextFilename('NOTES.TXT')).toBe(true);
    expect(isPlainTextFilename('readme.MD')).toBe(true);
    expect(isPlainTextFilename('Index.Html')).toBe(true);
  });

  it('returns false for filenames with no dot at all', () => {
    expect(isPlainTextFilename('Makefile')).toBe(false);
    expect(isPlainTextFilename('README')).toBe(false);
  });

  it('returns false for unknown extensions Graph might or might not handle', () => {
    expect(isPlainTextFilename('archive.zip')).toBe(false);
    expect(isPlainTextFilename('photo.jpg')).toBe(false);
    expect(isPlainTextFilename('manuscript.epub')).toBe(false);
  });

  it('uses the LAST dot to identify the extension (handles dotted filenames)', () => {
    expect(isPlainTextFilename('archive.tar.gz')).toBe(false); // gz is not in the whitelist
    expect(isPlainTextFilename('config.local.yaml')).toBe(true); // yaml is whitelisted
  });
});
