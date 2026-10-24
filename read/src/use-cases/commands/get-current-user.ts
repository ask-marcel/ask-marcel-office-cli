import { z } from 'zod';
import { buildSelectableCommand } from '@ask-marcel/office-core/use-cases/commands/build-command.ts';
import type { ReadCommandMeta } from '@ask-marcel/office-core/use-cases/commands/command-types.ts';
import { selectExpandOptions } from '@ask-marcel/office-core/use-cases/commands/odata-query.ts';

// the full `/me` resource carries 20+ fields most LLM
// callers don't need (preferredLanguage, surname, accountEnabled, etc.). Ship a
// slim default that covers the common "who am I" question. Override with
// `--select` for the full profile.
const DEFAULT_SELECT = 'id,displayName,mail,userPrincipalName,jobTitle,officeLocation,mobilePhone';

const baseSchema = z.object({}).strict();
const { execute, schema } = buildSelectableCommand(() => '/me', baseSchema, { defaultSelect: DEFAULT_SELECT });

const meta: ReadCommandMeta = {
  summary:
    "Return the signed-in user's Microsoft Graph profile. The CLI ships a slim default `--select=id,displayName,mail,userPrincipalName,jobTitle,officeLocation,mobilePhone` covering the common identity fields. Pass `--select id,displayName,givenName,surname,preferredLanguage,...` to widen, or `--select '*'` for everything Graph returns.",
  category: 'user',
  graphMethod: 'GET',
  graphPathTemplate: '/me',
  graphDocsUrl: 'https://learn.microsoft.com/en-us/graph/api/user-get',
  options: [...selectExpandOptions],
  example: '{read-bin} get-current-user',
  responseShape: 'single Microsoft Graph `user` resource projected to the default `--select` set (or, when overridden, to the requested fields)',
  effect: 'read',
  scopesRequired: ['User.Read'],
};

export { execute, meta, schema };
