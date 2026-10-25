import { describe, expect, it } from 'bun:test';
import { createFileSystemFake } from '@ask-marcel/office-core/test-helpers/filesystem-fake.ts';
import { buildDeps } from './build-deps.ts';

describe('buildDeps composition root', () => {
  it('wires logger, graph client and process runner', () => {
    const fs = createFileSystemFake();
    const deps = buildDeps({ home: '/virtual/home', logLevel: 'error', fs });
    expect(typeof deps.logger.info).toBe('function');
    expect(typeof deps.graph.get).toBe('function');
    expect(typeof deps.processRunner.run).toBe('function');
  });

  it('uses default config values when no config is provided', () => {
    const previousLevel = process.env.ASKMARCEL_LOG_LEVEL;
    process.env.ASKMARCEL_LOG_LEVEL = 'error';
    try {
      const deps = buildDeps();
      expect(typeof deps.logger.info).toBe('function');
    } finally {
      if (previousLevel === undefined) delete process.env.ASKMARCEL_LOG_LEVEL;
      else process.env.ASKMARCEL_LOG_LEVEL = previousLevel;
    }
  });
});
