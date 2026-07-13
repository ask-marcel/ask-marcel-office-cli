import { z } from 'zod';
import type { Result } from '../../domain/result.ts';
import type { AuthManager } from '../../infra/auth.ts';

const schema = z.object({}).strict();

const execute = async (auth: AuthManager, options?: { force?: boolean }): Promise<Result<string, import('../../infra/auth.ts').AuthError>> => auth.getAccessToken(options);

export { execute, schema };
