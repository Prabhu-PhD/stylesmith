# CLAUDE.md — StyleSmith

Context file for Claude Code. Read this before writing any code in this repo.

---

## What this is

**StyleSmith** is a PowerPoint add-in that brings named, reusable styles to PowerPoint — for text, shapes, **shape geometry**, and tables. Define a style once, apply it across a deck, change it, and every shape carrying it updates.

The differentiator is the **geometry layer**: styles carry adjustment-handle values (corner radius, arrow proportions). Format Painter does not copy these. No competitor does this.

A **design token layer** sits beneath styles: tokens hold values, styles reference tokens. Change `brand-primary` once and every style bound to it — and every shape carrying those styles — updates.

Companion docs: `StyleSmith-PRD-v1.md`, `StyleSmith-UX-Spec-v1.md`, `IMPLEMENTATION-PLAN.md`, `ACCEPTANCE-CRITERIA.md`.

---

## ⚠️ CRITICAL: Validated Office JS API contract

**Everything in this section was established by live runtime introspection on 23 Jul 2026. Some of it contradicts the obvious/documented patterns. Do not "correct" it from memory — it was verified against a running PowerPoint host.**

### Adjustments is NOT a standard collection

The `Adjustments` object has **no `items`, no `getItemAt`, no `getCount`**. Its members are: `count`, `get`, `set`, `load`, `toJSON`. `_scalarPropertyNames` is `["count"]`.

```ts
// ✅ CORRECT — verified working
shape.adjustments.load("count");
await ctx.sync();
const n = shape.adjustments.count;              // number

const result = shape.adjustments.get(0);        // returns ClientResult<number>
await ctx.sync();                                // MUST sync before reading
const value = result.value;                      // e.g. 0.16667

shape.adjustments.set(0, 0.15);                 // write — verified by read-back
await ctx.sync();
```

```ts
// ❌ WRONG — all of these fail. Do not generate them.
shape.adjustments.getCount();                    // not a function
shape.adjustments.getItemAt(0);                  // not a function
shape.adjustments.items[0];                      // items is undefined
shape.adjustments.load("items");                 // succeeds but yields nothing
shape.adjustments.setItemAt(0, v);               // not a function
shape.adjustments.set({ items: [v] });           // throws InvalidArgument
result.load("value");                            // ClientResult has no load()
```

**`get(i)` returns a `ClientResult`, not a client object.** You do not `load()` it. You `sync()`, then read `.value`. Getting this wrong produces a silent `null` that looks like the API is broken.

**Beware false positives on write.** `adjustments.set(0, 0.4)` with a *broken reader* appears to succeed because nothing throws. Always verify writes by reading the value back.

### Members that DO NOT EXIST on Shape

Confirmed by enumerating all 81 members of `PowerPoint.Shape`:

- **No `shadow`.** No shadow, glow, reflection or 3-D effects in the style model. Do not attempt.
- **No `effects`.**
- **No `geometricShapeType`.** Styles **tune** shapes; they cannot convert a rectangle into a rounded rectangle. Applying a geometry layer to a shape with zero adjustments must skip gracefully and report it.

### Members that DO exist and are used
`adjustments · fill · lineFormat · textFrame · tags · getTable · customXmlParts · creationId · getParentSlide · group · parentGroup · id · name · type · left/top/width/height · rotation · delete · duplicate · setZOrder`

### Host support
PowerPointApi **1.1 – 1.10** confirmed present (1.11+ absent). Target 1.1 as manifest minimum; feature-detect anything above.

---

## ⚠️ CRITICAL: The dirty-flag rule

**Metadata-only writes are silently lost on PowerPoint web.**

Verified: tags written and read back successfully in-session → document closed and reopened → **tags gone**. The same test with a shape nudged first (dirtying the document) → **100% survival**.

A write that changes only tags or only a custom XML part does not mark the document modified, so autosave never fires.

**Rule: every operation that writes only a shape TAG (or other non-settings metadata) must be paired with a change that dirties the document.** Applying a style is safe (it changes fill/border/geometry). The exposure is in **tag-only writes**:
- Linking an already-correctly-formatted shape (tag only, no visual change)
- Adoption, repair and migration routines, wherever they write linkage tags without a visual change

Implement this once, centrally, in `office/storage.ts` — never ad hoc at call sites.

**RESOLVED (Phase 0 spike, 24 Jul 2026):** `Office.context.document.customXmlParts` is **undefined in PowerPoint** (Word/Excel only). Deck-level storage uses **`Office.context.document.settings`** (`set` / `get` / `saveAsync`), and **`settings.saveAsync` self-persists across close/reopen WITHOUT a dirtying change** — verified on web both with and without a shape nudge. So the dirty-flag guard does **not** apply to settings saves: **style/token definitions live in `document.settings` and creating/renaming one needs no dirtying pair.** The guard remains required only for tag-only writes (above). Desktop parity still to confirm (open question #4).

---

## ⚠️ CRITICAL: Performance rules

Measured on a 60-slide / 2,520-shape deck:

| Operation | Result |
|---|---|
| Scan 2,520 shapes + tags, batched (3 syncs) | ~2.0s ✅ |
| Write 420 shapes × 3 properties in ONE sync | **Hangs** ❌ |

**Rule 1 — always chunk writes.** Never queue an unbounded number of property writes into a single `ctx.sync()`. Chunk (25–100 shapes), sync per chunk, report progress, support cancellation at chunk boundaries.

**Rule 2 — always batch reads.** Queue all `load()` calls, then sync once. Never sync inside a per-slide loop.

```ts
// ✅ Batched read — 3 syncs total regardless of deck size
const slides = ctx.presentation.slides;
slides.load("items");
await ctx.sync();

const cols = slides.items.map(s => { s.shapes.load("items/name,items/type"); return s.shapes; });
await ctx.sync();

const all = cols.flatMap(c => c.items);
all.forEach(s => s.tags.load("items/key,items/value"));
await ctx.sync();
```

---

## ⚠️ CRITICAL: Linkage is deck-local

**Tags do NOT survive copy/paste between decks.** Confirmed by testing.

Do not attempt to find an identifier that survives paste. `creationId` is a *creation* ID — a pasted shape is a new shape. Nothing survives a clipboard round-trip reliably.

**The answer is fingerprint matching** (the adoption engine): compare a shape's actual formatting signature against style definitions. It works precisely because it depends on no identifier at all.

This also means **adoption is the product's onboarding path, not a repair tool.** Every pre-existing deck opens with 100% orphans.

---

## Architecture

### The hard boundary

```
src/
  core/          ← PURE. No Office JS. Fully unit-testable.
    schema/      ← Zod schemas, TS types, migrations
    tokens/      ← resolution, alias chains, cycle detection
    fingerprint/ ← signature computation, clustering, near-match
    styles/      ← style model operations
  office/        ← ALL Office JS lives here. Nowhere else.
    bridge.ts    ← interface that core/ depends on
    shapes.ts tags.ts adjustments.ts storage.ts sweep.ts
  ui/            ← React + Fluent UI v9
    views/ components/ state/
```

**Rule: `core/` must never import from `office/`.** `office/` implements an interface that `core/` defines. This keeps ~70% of the codebase testable without a PowerPoint host, and confines all the API weirdness above to one directory.

**Rule: no Office JS calls in React components.** Components call the store; the store calls the bridge.

---

## Data model

Two layers. Tokens hold values; styles reference tokens.

```jsonc
{
  "schemaVersion": 1,
  "tokenSets": [{ "id": "guid", "name": "Default", "origin": null, "tokens": [
    { "id": "guid", "name": "brand-primary", "type": "color",
      "value": { "kind": "themeColor", "slot": "accent1" } }
  ]}],
  "styles": [{
    "id": "guid", "name": "Takeaway", "origin": null, "basedOn": null,
    "layers": {
      "text":     { "fontSize": { "kind": "tokenRef", "tokenId": "guid" } },
      "shape":    { "fill":     { "kind": "tokenRef", "tokenId": "guid" } },
      "geometry": { "adjustments": [{ "kind": "tokenRef", "tokenId": "guid" }] },
      "table":    null
    }
  }]
}
```

### Value kinds — every property value is one of three
| Kind | Shape |
|---|---|
| `literal` | `{ kind: "literal", value: "#FFFFFF" }` |
| `tokenRef` | `{ kind: "tokenRef", tokenId: "guid" }` |
| `themeColor` | `{ kind: "themeColor", slot: "accent1" }` |

**Prefer `themeColor` for colour.** It cascades natively with the deck theme **and survives without the add-in installed.** Theme slots: `dk1 lt1 dk2 lt2 accent1–6 hlink folHlink`, plus major/minor fonts. Everything else (sizes, radii, weights, spacing) needs native tokens.

### Non-negotiable schema rules
- **IDs are GUIDs, never names.** Renaming must not sever links. Applies to styles and tokens.
- **`schemaVersion` is written on every save.** Migrations key off it.
- **Layers are individually nullable.** Enables partial styles and selective apply.
- **Alias chains are depth-limited with cycle detection** — reject `a → b → a` at write time.
- **Every style caches its last-resolved literal.** A missing/deleted token must degrade to the cached value, never blank a shape.
- **`origin` and `basedOn` are reserved** — always present, null in v1. Do not remove them.

---

## Stack

| Concern | Choice | Why |
|---|---|---|
| Language | TypeScript, `strict: true` | Schema-heavy product |
| UI | React 18 + **Fluent UI v9** (`@fluentui/react-components`) | Native Office look; automatic light/dark/high-contrast theme inheritance |
| Build | Vite | HTTPS dev server for sideloading; tree-shaking matters for panel cold start |
| State | Zustand | Crosses panel/modal boundary; ~1KB |
| Validation | Zod | Reading versioned JSON that may be older or hand-edited |
| Tests | Vitest | `core/` is pure and highly testable |
| Manifest | **XML add-in-only** | Directly supported on Mac; unified manifest is not. Proven in spike. |
| Hosting | Static (Vercel) | No backend in v1 |

**Do not hard-code colours.** Use Fluent design tokens (`colorNeutralBackground1`, `colorNeutralForeground1`, etc.) so the panel inherits the user's Office theme. See the UX spec §2.

---

## Coding conventions

- Named exports; no default exports except React pages
- `async/await` throughout; no `.then()` chains
- All Office JS calls wrapped in `PowerPoint.run()` inside `office/`
- Errors are typed results in `core/`, not thrown, except for programmer error
- Every bridge method takes an `AbortSignal` where it can run long
- Progress callbacks on anything that chunks
- No `any`. Use `unknown` plus a Zod parse at boundaries.
- Comments explain *why*, especially where the API is counter-intuitive — link back to this file

---

## UX invariants (from the UX spec — do not violate)

1. **Never change anything invisibly.** Every bulk action states its blast radius before firing: "47 shapes across 23 slides."
2. **Scope is always visible** — current slide / selected slides / whole deck, defaulting to last used.
3. **Progress for anything over ~1s**, with chunk-accurate counts and working cancellation.
4. **Token binding legible at a glance** — bound values show a 🔗 chip with the token name; literals show the raw value.
5. **Panel is 320px, single column.** No side-by-side layouts.
6. **Preserve local overrides** by default when applying.
7. **Native undo works** (confirmed) — surface "Ctrl+Z to undo" after bulk operations rather than building a custom undo stack.

---

## Open questions — spike before depending on them

**Phase 0 spike results (24 Jul 2026):**
1. ✅ **Group traversal — RESOLVED: children ARE enumerable.** `shape.group.shapes` returns the child collection (verified: two groups, 3 `GeometricShape` children each). Sweeps can descend into groups. Still confirm rotation/coordinate handling and nested groups when the sweep is built.
2. ✅ **Deck-level storage — RESOLVED.** `document.customXmlParts` does not exist in PowerPoint (Word/Excel only); storage uses `document.settings`. `settings.saveAsync` **self-persists across close/reopen without a dirtying change** (verified with and without a nudge). Style/token definitions therefore need no dirty guard; tag writes still do. See the dirty-flag section above. (Watch: `document.settings` payload-size limits — validate against a large token/style set in Phase 2.)
3. ✅ **Within-deck duplicate (Ctrl+D) — RESOLVED: tags SURVIVE.** A duplicated shape carries its tags (verified: 3 tagged → 6 tagged after Ctrl+D, values copied verbatim). Note this means duplicates inherit the *same* `STYLESMITH_ID` — linkage stays correct, but any per-shape unique identity cannot rely on tags alone.

**Still open:**
4. **Desktop parity** — re-verify all of the above on PowerPoint desktop (spikes above ran on web).
5. **`textFrame` depth** — especially bullet/list formatting, historically weak in Office JS.
6. **Table API granularity** — can header fill, banding, per-cell borders and padding actually be set?

---

## Things that will waste your time if you forget them

- `adjustments.get()` returns a **ClientResult** — sync, then `.value`. It has no `load()`.
- A write that doesn't throw **has not necessarily succeeded**. Verify by read-back.
- Metadata-only writes **vanish on web** unless the document is dirtied.
- Unbatched bulk writes **hang the host**. Always chunk.
- Tags **do not cross decks**. Don't design around them doing so.
- There is **no shadow API** and **no shape-type API**. Don't try.
