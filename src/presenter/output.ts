import type { Logger } from '../use-cases/ports/logger.ts';

type SuccessEnvelope = { readonly ok: true; readonly data: unknown; readonly nextLink?: string; readonly count?: number };

const isPlainRecord = (value: unknown): value is Record<string, unknown> => value !== null && typeof value === 'object' && !Array.isArray(value);

const wrap = (data: unknown): SuccessEnvelope => {
  if (!isPlainRecord(data)) return { ok: true, data };
  const nextLink = data['@odata.nextLink'];
  const count = data['@odata.count'];
  const cleaned: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(data)) {
    if (key === '@odata.nextLink' || key === '@odata.count') continue;
    cleaned[key] = value;
  }
  const envelope: SuccessEnvelope = {
    ok: true,
    data: cleaned,
    ...(typeof nextLink === 'string' ? { nextLink } : {}),
    ...(typeof count === 'number' ? { count } : {}),
  };
  return envelope;
};

const render = (data: unknown, logger: Logger): void => {
  logger.info('output_rendered', {});
  process.stdout.write(`${JSON.stringify(wrap(data))}\n`);
};

const renderError = (message: string): void => {
  process.stdout.write(`${JSON.stringify({ ok: false, error: message })}\n`);
};

export { render, renderError };
