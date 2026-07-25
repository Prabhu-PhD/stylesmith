# StyleSmith — Technical Implementation Plan
**v1.0 · 23 July 2026 · companion to PRD v1.0 and UX Spec v1.0**

---

## 1. Stack decisions and rationale

Each choice is justified by a constraint of this specific product, not by general preference.

| Concern | Decision | Constraint driving it |
|---|---|---|
| **Language** | TypeScript, `strict: true` | The product is a schema engine. Types are the specification. |
| **UI framework** | React 18 | ~8 distinct views with shared state crossing panel/modal boundaries; heavy conditional rendering |
| **Component library** | **Fluent UI v9** (`@fluentui/react-components`) | The UX spec's visual language *is* Fluent tokens. Fluent gives automatic light/dark/high-contrast inheritance from the user's Office theme. Hand-rolling this in vanilla means reimplementing a design system and watching it drift. |
| **Build** | Vite | Needs an HTTPS dev server for sideloading; needs aggressive tree-shaking because a 320px panel cold-starts on every open |
| **State** | Zustand | Shared state (styles, tokens, scan cache, operation progress) accessed from panel and modal. ~1KB, no provider nesting. |
| **Validation** | Zod | Reads versioned JSON from a document possibly written by an older build. Runtime validation and TS types from one definition. |
| **Testing** | Vitest | `core/` is pure logic — token resolution, fingerprinting, migrations — and highly testable without a host |
| **Manifest** | **XML add-in-only** | Directly supported on Windows, Mac, web and mobile. The unified manifest is *not* directly supported on Mac — it relies on an XML manifest generated at AppSource deployment. Mac support is a headline benefit, so XML is the safer route. Also already proven working in our spike. |
| **Hosting** | Static (Vercel) | No backend in v1. Revisit only when licence-key validation arrives. |

### Rejected, with reasons
- **Vanilla TypeScript** — smallest bundle, but reimplementing Fluent's visual language across three Office themes by hand is a large, drift-prone cost that undercuts UX principle #1 ("native, not bolted on").
- **Unified JSON manifest** — no direct Mac support (see above).
- **Redux / MobX** — heavier than this state footprint warrants.
- **Webpack** — Microsoft's generator default, but Vite's dev server and HMR are materially better for iteration.

---

## 2. Project structure

```
stylesmith/
├─ CLAUDE.md                  ← read first
├─ manifest/
│  ├─ manifest.xml            ← production
│  └─ manifest-localhost.xml  ← dev, points at https://localhost:3000
├─ src/
│  ├─ core/                   ← PURE. no Office JS. unit-tested.
│  │  ├─ schema/
│  │  │  ├─ types.ts          ← Style, Token, TokenSet, ValueKind
│  │  │  ├─ schemas.ts        ← Zod definitions
│  │  │  ├─ migrations.ts     ← keyed on schemaVersion
│  │  │  └─ defaults.ts
│  │  ├─ tokens/
│  │  │  ├─ resolve.ts        ← ref → value, alias chains
│  │  │  ├─ cycles.ts         ← cycle detection, depth limit
│  │  │  └─ usage.ts          ← which styles/shapes use a token
│  │  ├─ fingerprint/
│  │  │  ├─ signature.ts      ← formatting → comparable signature
│  │  │  ├─ cluster.ts        ← group shapes by signature
│  │  │  └─ match.ts          ← exact / near match with tolerance
│  │  └─ styles/
│  │     ├─ model.ts          ← CRUD on the style graph
│  │     └─ diff.ts           ← style vs shape delta (drift)
│  ├─ office/                 ← ALL Office JS. nowhere else.
│  │  ├─ bridge.ts            ← interface core/ depends on
│  │  ├─ context.ts           ← PowerPoint.run wrapper, error mapping
│  │  ├─ shapes.ts            ← batched enumeration
│  │  ├─ tags.ts              ← linkage read/write
│  │  ├─ adjustments.ts       ← ⚠️ the ClientResult pattern
│  │  ├─ formatting.ts        ← fill, lineFormat, textFrame
│  │  ├─ tables.ts
│  │  ├─ storage.ts           ← custom XML part + ⚠️ dirty-flag guard
│  │  ├─ theme.ts             ← theme colour slots
│  │  └─ sweep.ts             ← ⚠️ chunked execution + progress + cancel
│  ├─ ui/
│  │  ├─ App.tsx
│  │  ├─ state/store.ts       ← Zustand
│  │  ├─ views/               ← StylesView, TokensView, StyleDetail,
│  │  │                          TokenDetail, AdoptionFlow, ApplyModal
│  │  └─ components/          ← StyleRow, TokenRow, BindingControl,
│  │                             ScopeSelector, ProgressBar, EmptyState
│  └─ index.tsx
├─ tests/
└─ vite.config.ts
```

**The load-bearing rule:** `core/` never imports `office/`. `office/` implements interfaces `core/` declares. Roughly 70% of the codebase becomes testable without a PowerPoint host, and every API quirk stays confined to one directory.

---

## 3. Build phases

Foundation-first. Each phase has a gate that must pass before the next begins.

### Phase 0 — Scaffold and spikes (1–2 days)
Project setup, manifest, sideload working, Fluent theme wired to Office theme.

**Plus the three open spikes**, which do not block later phases but do inform them:
1. Group traversal — can group children be enumerated?
2. Does writing a custom XML part dirty the document?
3. Within-deck duplicate — do tags survive Ctrl+D?

**Gate:** panel loads in PowerPoint web and desktop, renders a Fluent component, correctly follows the Office theme. Spike results recorded in `CLAUDE.md`.

### Phase 1 — Core schema and token engine (2–3 days)
Pure logic, no UI, no Office JS. Zod schemas, types, migration harness, token resolution with alias chains and cycle detection, cached-literal fallback, usage computation.

**Gate:** full unit test coverage on token resolution including cycles, missing tokens, deep chains, and theme-colour values. This is the foundation — it must be provably correct before anything sits on it.

### Phase 2 — Office bridge and storage (2–3 days)
Bridge interface, batched shape enumeration, tag read/write, adjustments using the verified ClientResult pattern, custom XML part storage **with the central dirty-flag guard**, theme colour reading.

**Gate:** round-trip test — write styles and tokens, close, reopen, read back identical. Must pass on both web and desktop.

### Phase 3 — Apply engine (3–4 days)
Single-shape apply across all four layers. Then chunked sweep with progress, cancellation at chunk boundaries, scope resolution, override preservation, and graceful skip on incompatible shape types.

**Gate:** apply a style to 400+ shapes on the 60-slide test deck without hanging, with accurate progress and working cancel. Record throughput and tune chunk size.

### Phase 4 — UI shell and Styles view (4–5 days)
App shell, segmented navigation, Zustand store, style list with preview swatches, style detail with the four layer sections, binding control, apply modal with scope and layer selection.

**Gate:** full create → apply → modify → apply-to-all loop working through the UI. The hero flow, end to end.

### Phase 5 — Tokens view (3–4 days)
Token list grouped by type, token detail with the three value kinds, theme-colour binding, cascade preview showing affected styles and shapes, create-token-from-value, unbind-to-literal.

**Gate:** edit one token → cascade preview shows accurate style and shape counts → apply → all affected shapes update.

### Phase 6 — Adoption engine (4–5 days)
Signature computation, clustering, near-match with tolerance, the three-step adoption flow, unlinked-shape banner, paste detection.

**Gate:** open the 60-slide test deck cold → scan → correctly cluster shapes → create styles → link. The onboarding path, proven.

### Phase 7 — Polish and submission (3–5 days)
Empty states, error states, keyboard shortcuts, accessibility pass, performance tuning, bundle optimisation, AppSource requirements (privacy policy, support URL, screenshots, validation).

**Gate:** AppSource validation passes; full manual test protocol green on Windows, Mac and web.

---

## 4. Critical implementation notes

### 4.1 The adjustments pattern
Isolate it entirely in `office/adjustments.ts`. Nothing else touches it.

```ts
export async function readAdjustments(ctx, shape): Promise<number[]> {
  shape.adjustments.load("count");
  await ctx.sync();
  const n = shape.adjustments.count;
  if (!n) return [];

  const results = Array.from({ length: n }, (_, i) => shape.adjustments.get(i));
  await ctx.sync();                       // ClientResult — sync then read .value
  return results.map(r => r.value);
}

export function queueAdjustmentWrite(shape, index: number, value: number): void {
  shape.adjustments.set(index, value);    // caller syncs at chunk boundary
}
```

### 4.2 The dirty-flag guard
One implementation, in `office/storage.ts`, used by every metadata write.

```ts
async function ensureDirty(ctx, slide): Promise<void> {
  // Re-set a property to its existing value so the host marks the doc modified.
  // Without this, tag- and XML-only writes are never autosaved on web.
}
```

### 4.3 Chunked sweep
```ts
export async function sweep(shapes, apply, opts: {
  chunkSize: number;                       // default 50, tuned in Phase 3
  onProgress: (done: number, total: number) => void;
  signal: AbortSignal;
}): Promise<SweepResult>
```
Cancellation stops at a chunk boundary and reports exactly what completed. A partially applied sweep the user knows about is recoverable; a silent one is not.

### 4.4 Fingerprint signature
Signature must be **stable** (same formatting → same signature) and **tolerant** (near-matches detectable). Suggested: normalise each layer to a canonical form, hash for exact matching, and keep the normalised form for computing per-property deltas on near matches. Tolerance is per-layer and configurable — colour distance differs in kind from a radius delta.

---

## 5. Testing strategy

| Layer | Approach |
|---|---|
| `core/` | Vitest unit tests. Target high coverage — this is pure logic and the foundation. |
| `office/` | Manual test protocols plus a hidden debug panel exposing bridge calls. Office JS is impractical to mock meaningfully. |
| `ui/` | Component tests on presentational pieces; manual for flows. |
| End-to-end | Scripted manual protocol against the 60-slide test deck, run on Windows, Mac and web before each release. |

Keep `cascade-perf-test-deck.pptx` (60 slides, 2,520 shapes) in the repo as the standing performance and adoption fixture.

---

## 6. Risks

| Risk | Mitigation |
|---|---|
| Group traversal impossible | Phase 0 spike. If unreachable, report skipped counts explicitly — never silently miss shapes. |
| Custom XML writes don't dirty the document | Phase 0 spike. If confirmed, the dirty guard extends to every save. |
| Desktop API surface differs from web | Verify at every phase gate on both, not only at the end. |
| `textFrame` lacks bullet formatting | Phase 4 discovery. If absent, state it in marketing rather than let users find it. |
| Table API too thin for real table styles | Phase 4/5 discovery. Table layer is independently nullable, so it can ship later without schema change. |
| Fluent v9 bundle size hurts cold start | Import components individually; lazy-load modal and adoption flows; measure in Phase 4. |
| AppSource rejection | Read submission requirements in Phase 0, not Phase 7. |
