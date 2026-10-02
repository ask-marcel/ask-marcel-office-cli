import { describe, expect, it } from 'bun:test';
import { createWinstonLogger } from './logger.ts';

const captureStream = async (stream: 'stdout' | 'stderr', run: () => void | Promise<unknown>): Promise<string> => {
  const target = process[stream];
  const original = target.write.bind(target);
  let captured = '';
  const swap = (chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    return true;
  };
  target.write = swap;
  try {
    await run();
  } finally {
    target.write = original;
  }
  return captured;
};

describe('winston logger adapter', () => {
  it('exposes info, warn, and error methods that conform to the Logger port', () => {
    const logger = createWinstonLogger({ logLevel: 'error' });
    expect(typeof logger.info).toBe('function');
    expect(typeof logger.warn).toBe('function');
    expect(typeof logger.error).toBe('function');
  });

  it('emits an info event as JSON to stderr (stdout reserved for command output)', async () => {
    const logger = createWinstonLogger({ logLevel: 'info' });
    const stderr = await captureStream('stderr', () => logger.info('user_logged_in', { userId: 'u-1' }));
    const stdout = await captureStream('stdout', () => logger.info('also_routed_to_stderr'));
    expect(stderr).toContain('user_logged_in');
    expect(stderr).toContain('u-1');
    expect(stdout).toBe('');
  });

  it('redacts secret-shaped meta keys before emitting', async () => {
    const logger = createWinstonLogger({ logLevel: 'info' });
    const out = await captureStream('stderr', () => logger.info('inbound_request', { token: 'super-secret-jwt' }));
    expect(out).not.toContain('super-secret-jwt');
    expect(out).toContain('[REDACTED]');
  });

  it("redacts the token cache's own field names, whatever their case", async () => {
    const logger = createWinstonLogger({ logLevel: 'info' });
    const meta = { access_token: 'secret-a', accessToken: 'secret-b', refresh_token: 'secret-c', RefreshToken: 'secret-d' };
    const out = await captureStream('stderr', () => logger.info('auth.debug', meta));
    for (const secret of ['secret-a', 'secret-b', 'secret-c', 'secret-d']) expect(out).not.toContain(secret);
  });

  it('emits warn and error levels', async () => {
    const logger = createWinstonLogger({ logLevel: 'warn' });
    const out = await captureStream('stderr', () => {
      logger.warn('disk_almost_full');
      logger.error('disk_full');
    });
    expect(out).toContain('disk_almost_full');
    expect(out).toContain('disk_full');
  });

  it('defaults to error level (silencing info events) when no config or env var is given', async () => {
    const previous = process.env.ASKMARCEL_LOG_LEVEL;
    delete process.env.ASKMARCEL_LOG_LEVEL;
    try {
      const logger = createWinstonLogger();
      const out = await captureStream('stderr', () => {
        logger.info('this_should_be_silent');
        logger.error('this_should_appear');
      });
      expect(out).not.toContain('this_should_be_silent');
      expect(out).toContain('this_should_appear');
    } finally {
      if (previous === undefined) delete process.env.ASKMARCEL_LOG_LEVEL;
      else process.env.ASKMARCEL_LOG_LEVEL = previous;
    }
  });

  it('honours ASKMARCEL_LOG_LEVEL env var when no config logLevel is provided', async () => {
    const previous = process.env.ASKMARCEL_LOG_LEVEL;
    process.env.ASKMARCEL_LOG_LEVEL = 'info';
    try {
      const logger = createWinstonLogger();
      const out = await captureStream('stderr', () => logger.info('env_driven_event'));
      expect(out).toContain('env_driven_event');
    } finally {
      if (previous === undefined) delete process.env.ASKMARCEL_LOG_LEVEL;
      else process.env.ASKMARCEL_LOG_LEVEL = previous;
    }
  });
});
