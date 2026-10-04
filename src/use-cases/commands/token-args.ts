import type { Result } from '../../domain/result.ts';
import { err, ok } from '../../domain/result.ts';
import { tenantId } from '../../domain/tenant-id.ts';
import type { TokenFingerprint } from '../../domain/token-fingerprint.ts';
import { parseTokenFingerprint } from '../../domain/token-fingerprint.ts';
import type { TokenRequest, TokenTier } from '../ports/token-issuer.ts';

// The one line the helper prints when it hands out no token. `tier` is null when
// the caller named none the helper knows.
type TokenFailureLine = { readonly errorCode: string; readonly tier: TokenTier | null; readonly message: string; readonly remedy: string };

type TokenArgs = { readonly request: TokenRequest; readonly rejected?: TokenFingerprint };

type Flag = '--tier' | '--tenant' | '--reject';

const TIERS: ReadonlyArray<TokenTier> = ['basic', 'elevated', 'chatsvcagg', 'ic3', 'guest'];
const FLAGS: ReadonlyArray<Flag> = ['--tier', '--tenant', '--reject'];

// No message echoes what the caller passed: a bearer pasted where a fingerprint
// belongs would otherwise land on stdout in a failure line.
const FLAGS_MESSAGE = 'The token helper takes --tier, --tenant and --reject, each at most once and each followed by its value.';
const TIER_MESSAGE = `--tier must be one of: ${TIERS.join(', ')}.`;
const REJECT_MESSAGE = '--reject must be the fingerprint of the refused token: its SHA-256 as 64 lowercase hex digits.';
const USAGE_REMEDY =
  'Call it as `ask-marcel-office token --tier <basic|elevated|chatsvcagg|ic3|guest> [--tenant <guid>] [--reject <fingerprint>]`; --tenant goes with the guest tier only.';

const invalid = (message: string, tier: TokenTier | null): Result<never, TokenFailureLine> => err({ errorCode: 'invalid_arguments', tier, message, remedy: USAGE_REMEDY });

const isFlag = (name: string | undefined): name is Flag => FLAGS.some((flag) => flag === name);

// `--flag value` pairs and nothing else.
const readFlags = (argv: ReadonlyArray<string>): Partial<Record<Flag, string>> | undefined => {
  const flags: Partial<Record<Flag, string>> = {};
  for (let at = 0; at < argv.length; at += 2) {
    const name = argv[at];
    const value = argv[at + 1];
    if (!isFlag(name) || Object.hasOwn(flags, name) || value === undefined) return undefined;
    flags[name] = value;
  }
  return flags;
};

// A guest token comes from a partner tenant's authority, so that tier and only
// that tier names one.
const requestFor = (tier: TokenTier, tenant: string | undefined): Result<TokenRequest, string> => {
  if (tier !== 'guest') return tenant === undefined ? ok({ tier }) : err('--tenant goes with the guest tier only.');
  if (tenant === undefined) return err('The guest tier needs --tenant <guid>, the partner tenant that issues the token.');
  const id = tenantId(tenant);
  return id.ok ? ok({ tier, tenant: id.value }) : err('--tenant must be a tenant GUID (8-4-4-4-12 hex digits).');
};

const parseTokenArgs = (argv: ReadonlyArray<string>): Result<TokenArgs, TokenFailureLine> => {
  const flags = readFlags(argv);
  if (flags === undefined) return invalid(FLAGS_MESSAGE, null);
  const tier = TIERS.find((known) => known === flags['--tier']);
  if (tier === undefined) return invalid(TIER_MESSAGE, null);
  const request = requestFor(tier, flags['--tenant']);
  if (!request.ok) return invalid(request.error, tier);
  const raw = flags['--reject'];
  if (raw === undefined) return ok({ request: request.value });
  const rejected = parseTokenFingerprint(raw);
  return rejected.ok ? ok({ request: request.value, rejected: rejected.value }) : invalid(REJECT_MESSAGE, tier);
};

export { parseTokenArgs };
export type { TokenArgs, TokenFailureLine };
