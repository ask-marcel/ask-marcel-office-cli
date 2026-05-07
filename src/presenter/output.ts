import type { Logger } from '../use-cases/ports/logger.ts';

const render = (data: unknown, logger: Logger): void => {
  logger.info('output_rendered', {});
  process.stdout.write(`${JSON.stringify(data)}\n`);
};

const renderError = (message: string): void => {
  process.stdout.write(`${JSON.stringify({ error: message })}\n`);
};

export { render, renderError };
