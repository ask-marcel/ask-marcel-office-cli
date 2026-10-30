/** The three published packages of the split. */
export const PACKAGE_ROLES = ['auth', 'read', 'write'] as const;
export type PackageRole = (typeof PACKAGE_ROLES)[number];
