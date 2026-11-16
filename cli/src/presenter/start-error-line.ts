import type { StartError } from '../domain/start-error.ts';

const FACADE = 'ask-marcel-office-cli';
const AGAIN = `Install ${FACADE} again.`;

const cause = (error: StartError): string => {
  if (error.type === 'package_not_installed') return `the package ${error.packageName} is not installed. ${AGAIN}`;
  if (error.type === 'manifest_unreadable') return `the package.json of ${error.packageName} cannot be read. ${AGAIN}`;
  if (error.type === 'no_single_bin') return `the package.json of ${error.packageName} does not name exactly one bin. ${AGAIN}`;
  return `the bin of ${error.packageName} did not start: ${error.message}`;
};

// The one stderr line of a facade bin that could not start: the package, the cause and the next step.
export const startErrorLine = (error: StartError): string => `${FACADE}: ${cause(error)}\n`;
