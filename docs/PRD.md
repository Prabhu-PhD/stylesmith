# StyleSmith — Product Requirements Document
**v1.0 · 23 July 2026 · Toss The Coin**
*Status: architecture validated by live testing. Ready for build planning.*
*Product name confirmed 23 Jul 2026. Trademark check pending on IP India / USPTO / EUIPO — note The Stylesmiths™ (interior design, UK), different class.*

---

## 1. Summary

**One-line pitch:** *Word's Styles panel, finally in PowerPoint — for text, shapes, geometry and tables.*

StyleSmith is a PowerPoint add-in that brings named, reusable styles to a tool that has never had them. A designer defines a style once, applies it across a deck, and when the style changes, every shape carrying it updates in one action.

Its defining capability — the one no competitor has — is that a StyleSmith style carries **shape geometry**: corner radius, arrow proportions, pie angles. Format Painter does not copy these values. That is precisely why designers eyeball corner radii today.

**Platform decision (validated):** Office JS → Windows, Mac, and web, distributed via AppSource.

---

## 2. Problem

PowerPoint has no style system. Word has had one for thirty years.

Today, propagating a formatting decision across a deck means either Format Painter (which misses geometry entirely and must be applied shape by shape) or manual slide-by-slide editing. A single decision — "make the takeaway boxes navy with tighter corners" — costs 30+ minutes on a 60-slide deck and still produces inconsistency, because eyeballing a corner radius 47 times produces 47 slightly different corner radii.

The cost compounds in agencies, where the same brand's decks are rebuilt continuously by different designers with no shared enforcement mechanism beyond a PDF style guide nobody opens.

**The gap in one sentence:** PowerPoint has a *theme* system (deck-wide colours and fonts) and a *master* system (layout inheritance), but nothing for the object-level formatting that designers actually spend their time on.

---

## 3. Users

### Primary — the agency presentation designer
Builds 20–80 slide decks under deadline, for clients with brand systems. Lives in PowerPoint 6+ hours a day. Uses Ctrl+Shift+C/V constantly. Highly sensitive to consistency because their work is judged on it. **Buys tools with their own money if the tool saves real hours.**

### Secondary — the in-house brand/marketing team
Maintains templates and polices deck quality across a company. Cares about enforcement and repeatability more than speed. **Buys per-seat, needs a procurement story.**

### Tertiary — the power-user consultant / analyst
Consulting, finance, corporate strategy. Builds dense decks, values precision and speed, comfortable with add-ins. **Buys individually, low support tolerance.**

### Explicit non-user
The occasional PowerPoint user making a quarterly update deck. They don't feel the pain and won't pay. Do not design for them.

---

## 4. Positioning & competition

| Tool | What it does | Where StyleSmith differs |
|---|---|---|
| **Format Painter** (native) | Copies fill, border, text formatting, one shape at a time | Named, persistent, reusable; deck-wide propagation; **carries geometry** |
| **PowerPoint Themes** (native) | Deck-wide colour/font sets | Object-level, not deck-level; arbitrary named styles, not a fixed slot system |
| **Empower, BrightSlide, think-cell** | Various productivity add-ins, some with brand libraries | Purpose-built styles engine; geometry layer; one-time price, not enterprise contract |
| **BrandGuard** (ours) | Compliance — detect and fix brand violations | Adjacent, deliberately: BrandGuard *enforces*, StyleSmith *creates*. Cross-sellable, non-overlapping |

**The headline nobody can copy quickly:** the geometry layer. It requires the `Shape.Adjustments` API, which is recent and undocumented enough that we had to reverse-engineer its interface through runtime introspection to find it.

---

## 5. Scope

### What StyleSmith is
A named-style engine for PowerPoint objects: create a style, apply it, modify it, propagate it deck-wide.

### What StyleSmith is not
- Not a template manager
- Not a compliance scanner (that's BrandGuard)
- Not a slide library or content manager
- Not a theme editor
- Not a design system documentation tool

---

## 6. The style model — four layers

A style definition contains up to four layers. **Each layer is independently optional** (schema-enforced), which enables partial styles and selective application.

| Layer | Properties | API | Status |
|---|---|---|---|
| **Text** | Font family, size, weight, colour, line spacing, space before/after, alignment | `shape.textFrame` → textRange | ✅ Available (bullet support to verify) |
| **Shape** | Fill (solid/none), border colour, border weight, border dash style | `shape.fill`, `shape.lineFormat` | ✅ Confirmed |
| **Geometry** ⭐ | Adjustment handle values — corner radius, arrow head proportions, pie/arc angles, callout tail | `shape.adjustments` | ✅ **Read + write confirmed** |
| **Table** | Header row format, banding, border rules, cell padding | `shape.getTable()`, Table API 1.8 | ✅ Available (granularity to verify) |

### Excluded from the model, with reason
| Excluded | Reason |
|---|---|
| Shadow, glow, reflection, 3-D effects | **No `shadow` or `effects` member exists on Shape in Office JS.** Confirmed by full member enumeration. |
| Shape type (rect → rounded rect) | **No `geometricShapeType` member exists.** Styles *tune* shapes; they cannot convert them. |
| Chart formatting | Native `.crtx` chart templates already solve this. Point users there. |

---

## 7. Feature specification

### 7.1 Core engine

| ID | Feature | Description | Priority |
|---|---|---|---|
| **S1** | Create style from selection | Format a shape → "Save as style" → name it. Captures all four layers. | P0 |
| **S2** | Apply style | Select one or many shapes → click style in panel. | P0 |
| **S3** | ⭐ Modify → Apply to all | Edit a style definition → every tagged shape in scope updates. **The hero feature.** | P0 |
| **S4** | Update style to match selection | Format one shape manually → push its formatting back into the style definition. | P0 |
| **S5** | Select all instances | Select every shape carrying a style. Doubles as navigation. | P0 |
| **S6** | Detach / clear | Remove a shape's style linkage. Option to keep or reset formatting. | P0 |
| **S7** | Rename / delete style | Rename is safe — linkage is keyed on GUID, not name. Delete prompts on in-use styles. | P0 |
| **S8** | Duplicate style | "New style based on this one" — the common authoring path. | P1 |

### 7.2 Safety & control
*Every feature in this group exists because of a measured finding, not a hunch.*

| ID | Feature | Driven by |
|---|---|---|
| **S9** | **Preview before apply** — "This will change 47 shapes across 23 slides. Preview / Apply / Cancel" | Bulk operations at scale need a confirmation step |
| **S10** | **Scope control** — current slide / selected slides / entire deck | Deck-wide must not be the only option |
| **S11** | **Chunked execution + progress indicator** | ⚠️ **Architectural.** Unbatched writes to 420 shapes hang the API. Chunking is mandatory, not an optimisation. |
| **S12** | **Selective layer application** — push colour without re-tuning geometry | Nullable layers make this nearly free |
| **S13** | **Preserve local overrides** on apply | Word-consistent; users expect manual tweaks to survive |
| **S14** | **Dirty-flag guard** on metadata-only writes | ⚠️ Tag-only writes are not autosaved on web and vanish silently. See §9.3. |

### 7.3 Adoption & management
*S15 is the product's true entry point — see §8.1.*

| ID | Feature | Description | Priority |
|---|---|---|---|
| **S15** | ⭐ **Adoption engine** | Fingerprint-match unlinked shapes against style definitions; offer to link. Exact match / near match with diff / bulk adopt / manual assign. | **P0** |
| **S16** | **Cross-deck style import** | Import style definitions from another .pptx, with conflict resolution (rename / overwrite / keep both). | P0 |
| **S17** | **Style inventory** | Every style, usage count, unused flagged. | P1 |
| **S18** | **Drift indicator** | Shapes whose formatting has diverged from their style. Natural BrandGuard cross-sell hook. | P1 |
| **S19** | **Panel search / grouping** | Needed past ~30 styles. | P1 |
| **S20** | **Keyboard shortcuts** | Word has Ctrl+Alt+1 for Heading 1. Designers are speed users. | P2 |

### 7.4 Design tokens ⭐
*Foundation decision taken 23 Jul 2026 — schema-level, see §10.*

| ID | Feature | Description | Priority |
|---|---|---|---|
| **S25** | **Token library** | Named values grouped by type: colour, type, geometry, stroke. Panel section of its own. | P0 |
| **S26** | **Bind any property to a token** | Style properties reference tokens instead of literals. Binding visible at a glance. | P0 |
| **S27** | ⭐ **Edit token → cascades to every style using it** | Change `brand-primary` once; every style bound to it updates, and every shape carrying those styles. | P0 |
| **S28** | **Theme colour binding** | Bind a colour token to a PowerPoint theme slot. Survives without the add-in. Preferred over literals. | P0 |
| **S29** | **Token usage count** | "Used by 4 styles, 112 shapes" — before you change anything. | P0 |
| **S30** | **Create token from a value** | Promote a literal in a style to a reusable token in one action. | P1 |
| **S31** | **Unbind to literal** | Break a token binding, keep the current value. | P1 |
| **S32** | **Token sets** — swap an entire palette | v2. Schema supports it from day one. | v2 |

### 7.5 Platform & compatibility

| ID | Feature |
|---|---|
| **S21** | Style definitions stored **inside the .pptx** (custom XML part) — styles travel with the file |
| **S22** | **Graceful degradation** — without the add-in, shapes retain formatting; linkage dormant, not broken |
| **S23** | **Windows + Mac + web** via AppSource |
| **S24** | **Native undo** — Ctrl+Z reverts add-in changes (confirmed); no custom undo engine required |

---

## 8. Key user flows

### 8.1 First run — the adoption flow ⭐
*This is the most important flow in the product. Every deck a user opens on day one contains zero linked shapes.*

1. User opens an existing 60-slide client deck and launches StyleSmith
2. Panel shows: *"No styles in this deck yet. Scan for candidates?"*
3. StyleSmith clusters shapes by formatting signature: *"14 shapes share one format. 9 share another. Create styles from these?"*
4. User names them — Takeaway, KPI Card — and StyleSmith creates the definitions **and links all matching shapes in one action**
5. The deck is now under management, in under a minute

**Without this flow, StyleSmith only works on decks built from scratch inside StyleSmith — which is no deck anyone has.**

### 8.2 The hero flow — takeaway box
1. Designer formats one takeaway box → "Save as style: Takeaway"
2. Applies it while building (or adopts existing ones via 8.1)
3. Client: *"tighter corners, navy, and bump the type"*
4. Designer edits one box → "Update style to match" → "Apply to all"
5. Preview: *"Will change 47 shapes across 23 slides"* → Apply
6. Every takeaway box updates deck-wide. **Corner radius included.** ~20 seconds.

### 8.3 The paste flow
1. Designer pastes 6 slides from another deck
2. StyleSmith detects unlinked shapes: *"These slides brought 11 unlinked shapes. 8 match your existing styles."*
3. Options: adopt into existing styles / import the source deck's styles / leave unlinked
4. Deck stays coherent instead of silently drifting

---

## 9. Technical architecture

### 9.1 Validated API contract
*Established by runtime introspection — the Adjustments interface does not match the documented collection pattern.*

```js
// Linkage
shape.tags.add("STYLESMITH_ID", "<guid>");
shape.tags.load("items/key,items/value");

// Geometry — NOT a standard collection
shape.adjustments.load("count");        // scalar property
shape.adjustments.get(i);               // returns ClientResult → sync → .value
shape.adjustments.set(i, value);        // write (confirmed: 0.16667 → 0.15)

// Shape + text + table
shape.fill.setSolidColor(hex);
shape.lineFormat.color / .weight / .dashStyle
shape.textFrame                          // → textRange formatting
shape.getTable()                         // Table API 1.8
```

**Host support measured:** PowerPointApi 1.1 through 1.10 (1.11+ not present).

### 9.2 Performance architecture
| Operation | Measured | Implication |
|---|---|---|
| Scan 2,520 shapes + tags, batched (3 syncs) | **~2.0 s** | Read path is a non-issue |
| Write 420 shapes × 3 properties, single sync | **Hangs** | ⚠️ Must chunk |

**Rule:** all writes execute in chunks with `ctx.sync()` per chunk and a progress callback. Chunk size to be tuned (25–100). All reads use batched nested loads with the minimum sync count.

### 9.3 ⭐ The dirty-flag rule
**Metadata-only writes do not mark the document modified, are therefore never autosaved on web, and are silently lost.** Verified: tags written and read back successfully in-session, gone after reopen — until a shape was nudged, after which they survived at 100%.

Any operation writing only metadata must be paired with a change that dirties the document. Normal use is safe (applying a style always changes fill/border/geometry). The exposure is in:
- Linking an already-correctly-formatted shape (tag only, no visual change)
- Creating or renaming a style definition without applying it
- Repair, migration, and adoption routines

**Open:** confirm whether writing to a custom XML part dirties the document. If not, style definitions themselves are exposed.

### 9.4 Linkage model
- **Deck-local by design.** Tags do not survive copy/paste between decks (confirmed by testing).
- Therefore: **no identifier is pursued that survives paste.** `creationId` is a creation ID; a pasted shape is a new shape.
- **Fingerprint matching (S15) is the answer** precisely because it depends on no identifier at all.

---

## 10. Data model

Stored as JSON inside a custom XML part in the .pptx. **Two-layer model: tokens hold values, styles reference tokens.**

### 10.1 Why tokens
Without a token layer, every style stores literal values. A rebrand means opening thirty styles and editing each one. With tokens, six values change and every style — and every shape carrying them — updates.

It is not only colour. A shared `radius-md` token means one edit re-tunes the corner radius of every card-like shape in the deck. Combined with the geometry layer, nothing on the market does this.

It also makes v2 nearly free: **a brand library is a token set.** Brand scoping stops being a new feature and becomes "swap the token values."

### 10.2 Schema

```jsonc
{
  "schemaVersion": 1,

  "tokenSets": [                     // ⭐ v1 has exactly one; v2 adds a switcher
    {
      "id": "guid",
      "name": "Default",
      "origin": null,                // reserved: brand library ID (v2)
      "tokens": [
        {
          "id": "guid",              // ⭐ reference key — NEVER the name
          "name": "brand-primary",
          "type": "color",           // color | fontFamily | fontSize | fontWeight
                                     // lineSpacing | radius | strokeWeight | dashStyle | spacing
          "value": { "kind": "themeColor", "slot": "accent1" },
          "description": "Primary brand blue"
        }
      ]
    }
  ],

  "styles": [
    {
      "id": "guid",                  // ⭐ linkage key — NEVER the name
      "name": "Takeaway",
      "origin": null,                // reserved: component library (v2)
      "basedOn": null,               // reserved: style inheritance (v2)
      "layers": {
        "text": {
          "fontFamily": { "kind": "tokenRef", "tokenId": "guid-font-body" },
          "fontSize":   { "kind": "tokenRef", "tokenId": "guid-size-body" },
          "color":      { "kind": "literal",  "value": "#FFFFFF" }
        },
        "shape": {
          "fill":         { "kind": "tokenRef", "tokenId": "guid-brand-primary" },
          "borderWeight": { "kind": "tokenRef", "tokenId": "guid-stroke-thin" }
        },
        "geometry": {
          "adjustments": [ { "kind": "tokenRef", "tokenId": "guid-radius-md" } ]
        },
        "table": null                // layers remain individually nullable
      }
    }
  ]
}
```

### 10.3 Value kinds
Every property value is one of three kinds — this union is the core of the model.

| Kind | Shape | Use |
|---|---|---|
| `literal` | `{ "kind": "literal", "value": "#FFFFFF" }` | One-off values not worth tokenising |
| `tokenRef` | `{ "kind": "tokenRef", "tokenId": "guid" }` | The default for anything reused |
| `themeColor` | `{ "kind": "themeColor", "slot": "accent1" }` | ⭐ **Preferred for colour** — see 10.4 |

### 10.4 ⭐ Theme colours are the preferred colour binding
PowerPoint already has a token system: twelve theme colour slots (`dk1`, `lt1`, `dk2`, `lt2`, `accent1`–`accent6`, `hlink`, `folHlink`) plus major/minor fonts.

Binding to a theme slot is **strictly better than a literal** because:
- The deck's own theme cascades natively — changing the theme updates everything, with or without StyleSmith
- **The binding survives without the add-in installed**
- It respects PowerPoint's colour variants and Designer

But theme slots cover only colour and two font families. Type scale, radii, stroke weights and spacing need StyleSmith's own tokens. **So the model is deliberately hybrid:** prefer theme references where they exist, use native tokens everywhere else.

### 10.5 Resolution rules
1. Resolve `tokenRef` → token → its value (which may itself be a `tokenRef`)
2. **Alias chains permitted, depth-limited, with cycle detection** — `a → b → a` must be caught at write time, not runtime
3. **Missing token → fall back to the last-resolved literal**, cached on the style. A deleted or unresolvable token must never break a style or blank a shape.
4. Precedence: token value → style layer value → local shape override (overrides always preserved, per S13)

### 10.6 Aliasing: schema now, UI later
The schema permits token→token references, which enables the standard two-tier design-system pattern (primitives like `blue-600`, semantics like `surface-emphasis` pointing at them).

**v1 UI exposes a single flat token list.** The capability exists structurally so v2 can reveal it without a migration. This is the same discipline as `basedOn` and `origin`.

### 10.7 Foundation decisions — cheap now, migrations later
| Decision | Rationale |
|---|---|
| **Style ID is a GUID, never the name** | Renaming must not sever every link |
| **Token ID is a GUID, never the name** | Same reason, one layer down |
| **`schemaVersion` from day one** | v2 cannot safely migrate v1 decks without it |
| **Layers individually nullable** | Enables partial styles and selective apply (S12) |
| **Value is a three-kind union** | Literal / token / theme binding without schema change |
| **`tokenSets` as an array from day one** | v2 brand switching becomes a dropdown, not a migration |
| **Aliasing permitted in schema** | Two-tier tokens later, no migration |
| **`origin` and `basedOn` reserved** | Component library and inheritance attach cleanly |
| **Last-resolved literal cached per style** | Guarantees styles never break on a missing token |

---

## 11. Known constraints

| Constraint | Consequence | Severity |
|---|---|---|
| No shadow/effects in Office JS | Not in the style model | Medium — fill and border cover most brand-critical formatting |
| No shape-type access | Styles tune shapes, can't convert them. Applying to a mismatched type must skip geometry gracefully and say so. | Low |
| Tags are deck-local | Adoption engine is mandatory (S15) | **Turned into a feature** |
| Dirty-flag dependency | Metadata-only writes need deliberate dirtying | Medium — must be systematic |
| Unbatched writes hang | Chunking is architectural | Medium — known and solvable |
| Group children not enumerated | Sweeps can't see inside groups | ⚠️ **Unresolved — see §15** |
| Push-based, not live-bound | Updates happen on user action | Low — Word behaves the same way |

---

## 12. Non-goals for v1
Component library · brand scoping · style sets · style inheritance · chart styles · cloud sync · collaboration features · styles on masters and layouts (pending decision) · shadow and effects · shape-type conversion

---

## 13. Roadmap

**v1 — MVP.** S1–S16, S21–S24. Must be sellable standing alone.
**v1.5.** S17–S20, style inspector, masters/layouts if testing supports it, bullet formatting if the API allows.
**v2 — Component Library.** The platform play: pre-built styled components (takeaway box, KPI card, quote block, callout) inserted already on-style; brand scoping so agencies keep one library per client; exportable brand libraries reusing S16's plumbing. Turns StyleSmith from a utility into a design system layer for PowerPoint.

---

## 14. Commercial model

| | Approach |
|---|---|
| **Base tool** | One-time purchase (Doc-Comp model — no subscription objection at entry) |
| **Ladder** | v2 component packs as paid downloads; brand-library licensing for agencies |
| **Licensing** | Direct sale with licence key (Polar.sh) likely simpler than AppSource in-app purchase |
| **Trial** | Mechanism affects where trial state lives — **decide before schema lock** |

---

## 15. Open items

### Blocking design decisions
1. **Group traversal** — can we reach children of a group? If not, sweeps silently miss grouped shapes, which is worse than failing loudly. Needs a probe and then a decision: recurse, or state the limitation prominently.
2. **Does writing a custom XML part dirty the document?** If not, style definitions face the same silent-loss risk as tags (§9.3).
3. **Within-deck duplicate** — do tags survive Ctrl+D inside one deck? Determines whether adoption fires constantly or occasionally, which changes its UI prominence.

### Verification before building the relevant layer
4. Desktop parity — same probes on PowerPoint desktop
5. `textFrame` depth, especially bullet formatting
6. Table API formatting granularity
7. Tag survival when a non-add-in user edits and saves the deck
8. Chunked write throughput and optimal chunk size

### Product decisions
9. Trial mechanics · 10. Styles on masters/layouts · 11. Multi-shape "new style from selection" sampling rule

---

## Appendix — Test evidence log (23 Jul 2026)

| Test | Result |
|---|---|
| Office JS host | PowerPoint / OfficeOnline, PowerPointApi 1.1–1.10 |
| Tag write + read, in session | ✅ 4/4 |
| Tag persistence across close/reopen | ✅ 100% (after document dirtied) |
| Tag persistence, metadata-only write | ❌ Lost — dirty-flag dependency identified |
| Tag survival, cross-deck paste | ❌ Not preserved — linkage is deck-local |
| `adjustments.get()` / `.set()` | ✅ 0.16667 → 0.15, verified by read-back |
| Native undo (Ctrl+Z) | ✅ Reverted add-in change to original colour |
| Scan 2,520 shapes, batched | ✅ ~2.0 s |
| Write 420 shapes × 3 props, one sync | ⚠️ Hang — chunking required |
| Shape member enumeration | 81 members; no `shadow`, no `effects`, no `geometricShapeType` |
| Group enumeration | Group returns as one shape; children not listed |
