import type { Logger, LogMeta } from '../use-cases/ports/logger.ts';

export type LoggerFake = Logger & {
  readonly calls: ReadonlyArray<{ readonly level: 'info' | 'warn' | 'error'; readonly event: string; readonly meta?: LogMeta }>;
};

export const createLoggerFake = (): LoggerFake => {
  const calls: { level: 'info' | 'warn' | 'error'; event: string; meta?: LogMeta }[] = [];
  return {
    calls,
    info: (event, meta) => {
      calls.push({ level: 'info', event, meta });
    },
    warn: (event, meta) => {
      calls.push({ level: 'warn', event, meta });
    },
    error: (event, meta) => {
      calls.push({ level: 'error', event, meta });
    },
  };
};
