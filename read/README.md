# @ask-marcel/office-read

> Not released yet. This package is a skeleton of the package split
> (`docs/plans/2026-10-01-package-split.md` in the repository). Until then, use
> [`ask-marcel-office-cli`](https://www.npmjs.com/package/ask-marcel-office-cli).

`@ask-marcel/office-read` will hold every command that changes nothing in the tenant (about 200), with the conversion of Office files, PDF and mail to markdown. Its bin is `ask-marcel-office-read`.

Today the bin answers `--version` and refuses every command, and the library
exports one type, `Result`. The code moves in with phase 6 of the split.

## Develop

This folder is one Bun workspace of the repository. Run `bun install` at the
repository root once, then from this folder:

| Command | What it does |
|:--|:--|
| `bun test` | The tests of this folder |
| `bun run lint` / `bun run lint:strict` | ESLint (fast / type-aware) |
| `bun run typecheck` | `tsc --noEmit` |
| `bun run coverage` | 100% on every tier of `src/` |
| `bun run mutate:changed` | Stryker on the changed domain and use-case files |
| `bun run build` | `dist/`: the bundles (core inlined) and the declarations |
| `bun run check:packed` | Packs the package and checks the tarball and a clean consumer install |

## License

MIT. See [LICENSE](LICENSE).
