import { z } from 'zod';
import type { Result } from '@ask-marcel/office-core/domain/result.ts';
import type { AuthError, AuthSession } from '../ports/auth-session.ts';

const schema = z.object({}).strict();

const execute = async (auth: AuthSession): Promise<Result<void, AuthError>> => auth.logout();

export { execute, schema };
