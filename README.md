# StyleSmith

PowerPoint add-in bringing named, reusable styles — for text, shapes, **shape geometry**, and tables — with a design token layer underneath.

**Start here: [`CLAUDE.md`](./CLAUDE.md)** — validated API contract, architectural rules, and the constraints that are not inferable from documentation.

## Contents

| Path | What it is |
|---|---|
| `CLAUDE.md` | ⚠️ **Read first.** Verified Office JS contract, critical rules, architecture, stack, conventions |
| `docs/PRD.md` | Product requirements — problem, users, features, data model, roadmap |
| `docs/UX-SPEC.md` | Panel design, Fluent tokens, IA, flows, states, accessibility |
| `docs/IMPLEMENTATION-PLAN.md` | Stack rationale, project structure, 8 phases with gates |
| `docs/ACCEPTANCE-CRITERIA.md` | Given/when/then for every P0 feature |
| `docs/QA-PROTOCOL.md` | Manual release testing across web, Windows, Mac |
| `reference/verified-api-patterns.ts` | Working code for the APIs that behave unexpectedly |
| `reference/manifest-reference.xml` | Manifest proven to sideload successfully |
| `fixtures/stylesmith-perf-test-deck.pptx` | 60 slides, 2,520 shapes — performance and adoption fixture |

## The four rules that matter most

1. **`adjustments.get(i)` returns a ClientResult** — sync, then read `.value`. It has no `load()`. It is not a collection.
2. **Metadata-only writes vanish on web** unless the document is dirtied.
3. **Always chunk writes.** 420 shapes in one sync hangs the host.
4. **Tags do not survive cross-deck paste.** Linkage is deck-local; adoption is the answer.

## Development

```bash
npm install
npm run dev:certs   # once per machine: installs a trusted local CA for HTTPS sideloading
npm run dev         # serves the task pane on https://localhost:3000
```

Then sideload `manifest/manifest-localhost.xml` into PowerPoint (web or desktop).

| Script | Does |
|---|---|
| `npm run dev` | Vite HTTPS dev server on `localhost:3000` |
| `npm run build` | Type-check (`tsc --noEmit`) then production build to `dist/` |
| `npm run typecheck` | Type-check only |
| `npm run test` | Vitest (unit + the core/→office/ architecture boundary guard) |
| `npm run lint` | ESLint, including the `import/no-restricted-paths` boundary rule |

Manifests live in `manifest/` — `manifest-localhost.xml` (dev, points at `localhost:3000`) and `manifest.xml` (prod).

## Status

Architecture validated by live testing, 23 July 2026. **All build phases (0–7) code-complete** — schema + token engine, Office bridge + storage, apply sweep, Styles view, Tokens view + cascade, adoption engine, and polish/submission scaffolding. 71 unit tests green; production build + AppSource manifest validation pass.

**Remaining before submission is host- and business-dependent** (see [docs/SUBMISSION.md](docs/SUBMISSION.md)): sideload + run the manual protocol on Windows/Mac/web, the real hosting origin and PNG icons, and the accumulated phase gates re-verified on a live host.
