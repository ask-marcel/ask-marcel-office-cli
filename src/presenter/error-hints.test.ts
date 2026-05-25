import { describe, expect, it } from 'bun:test';
import { findErrorHint } from './error-hints.ts';

describe('findErrorHint — Graph error translation (Audit Alex-session §2)', () => {
  it('maps `ErrorInvalidIdMalformed` to an actionable hint pointing the LLM at the right list-* command for sourcing IDs', () => {
    const result = findErrorHint('ErrorInvalidIdMalformed: Id is malformed.', 'ErrorInvalidIdMalformed');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('Source IDs from a sibling');
    expect(result?.hint).toContain('list-mail-messages');
  });

  it('maps `itemNotFound` to a "well-formed but missing — re-fetch" hint distinct from malformed IDs', () => {
    const result = findErrorHint('The specified object was not found in the store.', 'itemNotFound');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('well-formed but the resource is missing');
  });

  it('detects "Missing scope" anywhere in the message (not just as a structured code) and points at scopes-check + the appid scope ceiling', () => {
    const result = findErrorHint("Missing scope permissions on the request. API: 'Read.All' on resource '/me/...'", 'Forbidden');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('scopes-check');
    expect(result?.hint).toContain('fixed scope ceiling');
  });

  it('maps `accessDenied` to a permissions-explanation hint that covers mailbox, SharePoint, and chat-substrate cases', () => {
    const result = findErrorHint('Access denied.', 'accessDenied');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('delegated read access');
  });

  it('maps `InvalidAuthenticationToken` to a `login` instruction (the actionable remedy)', () => {
    const result = findErrorHint('Lifetime validation failed.', 'InvalidAuthenticationToken');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('ask-marcel login');
    expect(result?.hint).toContain('expiresInSeconds');
  });

  it('maps `TooManyRequests` to a throttling-recovery hint', () => {
    const result = findErrorHint('Throttled.', 'TooManyRequests');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('Retry-After');
  });

  it('catches the canonical KQL quoting trap ("An identifier was expected at position 0") and tells the LLM not to double-quote --query', () => {
    const result = findErrorHint('BadRequest: An identifier was expected at position 0.', 'BadRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('extra double-quotes');
  });

  it('catches the `$skip is not supported` family and points at next-page', () => {
    const result = findErrorHint('invalidRequest: $skip is not supported on this API.', 'invalidRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('next-page');
  });

  it('catches the $search + $filter / $orderby incompatibility', () => {
    const result = findErrorHint('Cannot combine $search and $filter.', 'BadRequest');
    expect(result?.source).toBe('graph');
    expect(result?.hint).toContain('relevance ranking');
  });
});

describe('findErrorHint — CLI-side rewrites and rejections', () => {
  it('matches errorCodes starting with `cli_rewrite_` and tags them as source=cli (the LLM should read the message text for the remedy, since the CLI already inlined it there)', () => {
    const result = findErrorHint('The --event-id is not a recurring series — find a seriesMaster ...', 'cli_rewrite_expand_series_not_recurring');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('read past the headline');
  });

  it('matches the search-mail-messages --filter+--query rejection (`cli_reject_search_with_filter`) with the actionable remedy in `hint`, not duplicated in `error` (Audit Alex-session §2 field-inversion fix)', () => {
    const result = findErrorHint('--filter is incompatible with $search on /me/messages — Graph rejects the combination with `SearchWithFilter`.', 'cli_reject_search_with_filter');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('list-mail-messages --filter');
    expect(result?.hint).toContain('KQL query string');
  });
});

describe('findErrorHint — Commander.js parser errors (Audit Alex-session §2 follow-up)', () => {
  it('maps `commander.unknownOption` to an actionable hint pointing at `<command> --help` and the terse manifest', () => {
    const result = findErrorHint("unknown option '--notarealflag'", 'commander.unknownOption');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('ask-marcel <command> --help');
    expect(result?.hint).toContain('help-json --terse --category');
  });

  it('maps `commander.missingMandatoryOptionValue` to a "required flag missing — re-read --help" hint', () => {
    const result = findErrorHint("required option '--message-id <value>' not specified", 'commander.missingMandatoryOptionValue');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('required CLI flag is missing');
  });

  it('maps `commander.optionMissingArgument` (same shape as missingMandatoryOptionValue but Commander emits it differently for --foo with no value) to the same hint family', () => {
    const result = findErrorHint("option '--message-id <value>' argument missing", 'commander.optionMissingArgument');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('required CLI flag is missing');
  });

  it('maps `commander.missingArgument` (positional) to a hint about positional args (mainly `docs <command>` / `help <command>`)', () => {
    const result = findErrorHint("missing required argument 'command'", 'commander.missingArgument');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('positional argument');
  });

  it('maps `commander.unknownCommand` to discovery-surface advice (`help-json --terse --category mail`)', () => {
    const result = findErrorHint("unknown command 'discover-person'", 'commander.unknownCommand');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('help-json --terse');
  });

  it('maps `commander.invalidArgument` (typed-value validation failure, e.g. --top abc) to an expected-shape hint', () => {
    const result = findErrorHint("Invalid argument 'abc' for option '--top'", 'commander.invalidArgument');
    expect(result?.source).toBe('cli');
    expect(result?.hint).toContain('value shape per flag');
  });
});

describe('findErrorHint — substrate errors (Audit Alex-session §2 follow-up)', () => {
  it('maps any `substrateHttp{status}_chatsvcagg` code to the experimental-substrate hint with the structured `source: "substrate"` classifier', () => {
    const result = findErrorHint('BadRequest', 'substrateHttp400_chatsvcagg');
    expect(result?.source).toBe('substrate');
    expect(result?.hint).toContain('chatsvcagg');
    expect(result?.hint).toContain('experimental');
    expect(result?.hint).toContain('ask-marcel login');
  });

  it('maps `substrateHttp{status}_ic3` codes through the same rule (single hint covers both substrate identities)', () => {
    const result = findErrorHint('Forbidden', 'substrateHttp403_ic3');
    expect(result?.source).toBe('substrate');
    expect(result?.hint).toContain('substrate');
  });

  it('substrate 5xx still routes through the substrate hint (transient — retry once)', () => {
    const result = findErrorHint('InternalServerError', 'substrateHttp503_chatsvcagg');
    expect(result?.source).toBe('substrate');
    expect(result?.hint).toContain('5xx is typically transient');
  });
});

describe('findErrorHint — validation (Zod / CLI flag parsing)', () => {
  it("treats messages that start with `--<flag>` as Zod validation rejections and tags them source='validation'", () => {
    const result = findErrorHint('--message-id is required', undefined);
    expect(result?.source).toBe('validation');
    expect(result?.hint).toContain('ask-marcel <cmd> --help');
  });

  it('also matches the generic "Validation error" prefix that Zod uses for nested-shape failures', () => {
    const result = findErrorHint('Validation error: top must be a positive integer', undefined);
    expect(result?.source).toBe('validation');
  });
});

describe('findErrorHint — no match', () => {
  it('returns undefined for an unknown Graph error so the bare `error:` line still renders (the historical shape — additive, not replacing)', () => {
    const result = findErrorHint('Some weird new Graph error nobody has seen before.', 'WeirdNewCode');
    expect(result).toBeUndefined();
  });

  it('returns undefined for a free-text error with no code and no recognisable pattern', () => {
    const result = findErrorHint('some unstructured failure', undefined);
    expect(result).toBeUndefined();
  });
});
