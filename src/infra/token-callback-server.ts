import { createServer } from 'node:http';
import type { Server } from 'node:http';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { Logger } from '../use-cases/ports/logger.ts';

type TokenCallbackPayload = {
  readonly access_token: string;
  readonly refresh_token?: string;
  readonly elevated_access_token?: string;
  readonly chatsvcagg_access_token?: string;
  readonly ic3_access_token?: string;
  readonly chatsvcagg_region?: string;
};

type TokenCallbackServer = {
  readonly port: number;
  readonly start: () => Promise<Result<TokenCallbackPayload, TokenCallbackError>>;
  readonly stop: () => Promise<void>;
};

type TokenCallbackError =
  | { type: 'bind_failed'; message: string }
  | { type: 'timeout'; message: string }
  | { type: 'invalid_payload'; message: string }
  | { type: 'server_closed'; message: string };

const ALLOWED_ORIGINS = [
  'https://teams.microsoft.com',
  'https://teams.live.com',
  'http://127.0.0.1',
  'http://localhost',
  'chrome-extension://',
];

const createTokenCallbackServer = (logger: Logger, timeoutMs: number = 5 * 60 * 1000): TokenCallbackServer => {
  let server: Server | null = null;
  let resolveCallback: ((result: Result<TokenCallbackPayload, TokenCallbackError>) => void) | null = null;
  let timeoutHandle: ReturnType<typeof setTimeout> | null = null;

  const stop = async (): Promise<void> => {
    if (timeoutHandle) {
      clearTimeout(timeoutHandle);
      timeoutHandle = null;
    }
    if (server) {
      return new Promise<void>((resolve) => {
        server!.close(() => {
          server = null;
          resolve();
        });
      });
    }
  };

  const start = (): Promise<Result<TokenCallbackPayload, TokenCallbackError>> => {
    return new Promise((resolve) => {
      resolveCallback = resolve;

      server = createServer((req, res) => {
        // CORS headers for browser extension
        res.setHeader('Access-Control-Allow-Origin', '*');
        res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
        res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

        if (req.method === 'OPTIONS') {
          res.writeHead(204);
          res.end();
          return;
        }

        if (req.method === 'POST' && req.url === '/token') {
          let body = '';
          req.on('data', (chunk) => {
            body += chunk;
            if (body.length > 1e6) {
              req.destroy();
              res.writeHead(413);
              res.end('Payload too large');
            }
          });

          req.on('end', () => {
            try {
              const payload = JSON.parse(body) as TokenCallbackPayload;

              if (!payload.access_token || typeof payload.access_token !== 'string') {
                res.writeHead(400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: 'missing access_token' }));
                return;
              }

              // Validate origin if present
              const origin = req.headers['origin'];
              if (origin && !ALLOWED_ORIGINS.some((allowed) => origin.startsWith(allowed))) {
                logger.info('token_callback.rejected_origin', { origin });
                res.writeHead(403, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: 'origin not allowed' }));
                return;
              }

              logger.info('token_callback.received', { hasAccessToken: true, hasRefreshToken: !!payload.refresh_token });

              res.writeHead(200, { 'Content-Type': 'application/json' });
              res.end(JSON.stringify({ status: 'ok' }));

              // Resolve the promise and stop the server
              if (resolveCallback) {
                resolveCallback(ok(payload));
                resolveCallback = null;
              }
              void stop();
            } catch (e) {
              const message = e instanceof Error ? e.message : String(e);
              logger.info('token_callback.parse_error', { message });
              res.writeHead(400, { 'Content-Type': 'application/json' });
              res.end(JSON.stringify({ error: 'invalid json' }));
            }
          });
          return;
        }

        // Health check endpoint
        if (req.method === 'GET' && req.url === '/health') {
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ status: 'waiting' }));
          return;
        }

        res.writeHead(404);
        res.end('Not found');
      });

      // Bind to port 0 — OS assigns a random available port
      server.listen(0, '127.0.0.1', () => {
        const address = server!.address();
        if (!address || typeof address === 'string') {
          resolve({ ok: false, error: { type: 'bind_failed', message: 'failed to get assigned port' } });
          return;
        }

        const port = address.port;
        logger.info('token_callback.listening', { port, timeoutMs });

        // Set timeout
        timeoutHandle = setTimeout(() => {
          logger.info('token_callback.timeout', { timeoutMs });
          if (resolveCallback) {
            resolveCallback(err({ type: 'timeout', message: `token callback timed out after ${timeoutMs}ms` }));
            resolveCallback = null;
          }
          void stop();
        }, timeoutMs);
      });

      server.on('error', (e) => {
        const message = e instanceof Error ? e.message : String(e);
        logger.info('token_callback.bind_error', { message });
        resolve(err({ type: 'bind_failed', message }));
      });
    });
  };

  // Return a proxy that exposes the port after start() resolves
  return {
    get port(): number {
      const address = server?.address();
      if (address && typeof address !== 'string') {
        return address.port;
      }
      return 0;
    },
    start,
    stop,
  };
};

export { createTokenCallbackServer };
export type { TokenCallbackError, TokenCallbackPayload, TokenCallbackServer };
