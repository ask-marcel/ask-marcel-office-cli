// Why a bin of the facade could not start the bin of a package (ADR 0005).
export type StartError =
  | { readonly type: 'package_not_installed'; readonly packageName: string }
  | { readonly type: 'manifest_unreadable'; readonly packageName: string }
  | { readonly type: 'no_single_bin'; readonly packageName: string }
  | { readonly type: 'bin_failed'; readonly packageName: string; readonly message: string };
