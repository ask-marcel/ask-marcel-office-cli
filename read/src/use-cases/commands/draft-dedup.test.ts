import { describe, expect, it } from 'bun:test';
import { matchExistingDrafts, normalizeThreadSubject } from './draft-dedup.ts';

type TestDraft = {
  readonly id: string;
  readonly subject: string;
  readonly conversationId: string;
  readonly toRecipients: ReadonlyArray<{ readonly emailAddress: { readonly address: string } }>;
  readonly ccRecipients?: ReadonlyArray<{ readonly emailAddress: { readonly address: string } }>;
};

const draft = (id: string, subject: string, to: ReadonlyArray<string>, conversationId = id): TestDraft => ({
  id,
  subject,
  conversationId,
  toRecipients: to.map((address) => ({ emailAddress: { address } })),
});

describe('normalizeThreadSubject reduces reply and forward prefixes to one thread key', () => {
  it('treats an RE: reply and an FW: forward of the same subject as the same thread', () => {
    expect(normalizeThreadSubject('RE: Contoso Q3')).toBe(normalizeThreadSubject('FW: Contoso Q3'));
  });

  it('collapses a stack of repeated reply prefixes to the bare subject', () => {
    expect(normalizeThreadSubject('RE: RE: RE: Budget')).toBe(normalizeThreadSubject('Budget'));
  });

  it('strips localized reply and forward prefixes (AW:, WG:, TR:)', () => {
    const bare = normalizeThreadSubject('Rapport');
    expect(normalizeThreadSubject('AW: Rapport')).toBe(bare);
    expect(normalizeThreadSubject('WG: Rapport')).toBe(bare);
    expect(normalizeThreadSubject('TR: Rapport')).toBe(bare);
  });

  it('strips a numbered reply prefix such as AW[2]:', () => {
    expect(normalizeThreadSubject('AW[2]: Rapport')).toBe(normalizeThreadSubject('Rapport'));
  });

  it('strips a parenthesized reply count such as FW(3):', () => {
    expect(normalizeThreadSubject('FW(3): Rapport')).toBe(normalizeThreadSubject('Rapport'));
  });

  it('strips the remaining localized reply and forward prefixes (SV, VS, ODP, RV, R, V)', () => {
    const bare = normalizeThreadSubject('Rapport');
    for (const prefix of ['SV', 'VS', 'ODP', 'RV', 'R', 'V']) {
      expect(normalizeThreadSubject(`${prefix}: Rapport`)).toBe(bare);
    }
  });

  it('ignores case and surrounding whitespace differences', () => {
    expect(normalizeThreadSubject('  re:   Hello World  ')).toBe(normalizeThreadSubject('Hello World'));
  });

  it('does not strip a real word that merely starts with a prefix letter', () => {
    expect(normalizeThreadSubject('Report ready')).toBe('report ready');
  });
});

describe('matchExistingDrafts finds drafts on one thread without trusting conversationId', () => {
  it('matches drafts on the same subject even when their conversationIds differ', () => {
    const drafts = [draft('a', 'RE: Contoso', ['kim@example.com'], 'conv-x'), draft('b', 'RE: Contoso', ['kim@example.com'], 'conv-y')];
    const matched = matchExistingDrafts(drafts, { subject: 'Contoso' });
    expect(matched.map((d) => d.id)).toEqual(['a', 'b']);
  });

  it('keeps only drafts that share a recipient when recipients are given', () => {
    const drafts = [draft('a', 'RE: Contoso', ['kim@example.com']), draft('b', 'RE: Contoso', ['other@example.com'])];
    const matched = matchExistingDrafts(drafts, { subject: 'Contoso', recipients: ['KIM@example.com'] });
    expect(matched.map((d) => d.id)).toEqual(['a']);
  });

  it('matches on a cc recipient as well as a to recipient', () => {
    const withCc: TestDraft = { ...draft('a', 'RE: Contoso', ['someone@example.com']), ccRecipients: [{ emailAddress: { address: 'kim@example.com' } }] };
    const matched = matchExistingDrafts([withCc], { subject: 'Contoso', recipients: ['kim@example.com'] });
    expect(matched.map((d) => d.id)).toEqual(['a']);
  });

  it('matches on subject alone when no recipients are supplied', () => {
    const drafts = [draft('a', 'FW: Contoso', ['anyone@example.com'])];
    expect(matchExistingDrafts(drafts, { subject: 'RE: Contoso' }).map((d) => d.id)).toEqual(['a']);
  });

  it('returns nothing for an empty draft list', () => {
    expect(matchExistingDrafts([], { subject: 'Contoso' })).toEqual([]);
  });

  it('excludes drafts on a different thread', () => {
    const drafts = [draft('a', 'RE: Contoso', ['kim@example.com']), draft('b', 'RE: Fabrikam', ['kim@example.com'])];
    expect(matchExistingDrafts(drafts, { subject: 'Contoso' }).map((d) => d.id)).toEqual(['a']);
  });
});
