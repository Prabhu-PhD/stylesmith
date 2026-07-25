# StyleSmith — UI/UX Specification
**v1.0 · 23 July 2026 · companion to StyleSmith PRD v1.0**

---

## 1. Design principles

1. **Native, not bolted on.** An Office add-in that invents its own visual language looks like a browser tab wearing a costume. Borrow Office's design language so it reads as part of PowerPoint.
2. **Never change anything invisibly.** Every bulk action states its blast radius before it fires. "47 shapes across 23 slides" is the product's core safety promise.
3. **Scope is always visible.** The user must never wonder what "apply" is about to touch.
4. **Progress for anything over a second.** Measured sweeps take real time; silence reads as a freeze.
5. **Binding is legible at a glance.** A designer should see instantly whether a value is bound to a token or hard-coded, without clicking.
6. **The panel is 320px.** One column. No side-by-side layouts, no wide tables, no horizontal scrolling.

---

## 2. Visual language — design tokens for the add-in itself

Use **Fluent UI 2** tokens so the panel inherits the user's Office theme (light, dark, high contrast) automatically. Do not hard-code colours.

### Colour — reference Fluent tokens, never hex
| Purpose | Fluent token |
|---|---|
| Panel background | `colorNeutralBackground1` |
| Card / row surface | `colorNeutralBackground2` |
| Hover surface | `colorNeutralBackground1Hover` |
| Primary text | `colorNeutralForeground1` |
| Secondary text | `colorNeutralForeground2` |
| Divider / border | `colorNeutralStroke2` |
| Accent (primary action) | `colorBrandBackground` |
| Warning (destructive preview) | `colorPaletteYellowBackground2` |
| Error | `colorPaletteRedForeground1` |

### Type ramp — Segoe UI Variable
| Role | Size / weight | Use |
|---|---|---|
| Section heading | 14px Semibold | "Styles", "Tokens" |
| Body / row label | 13px Regular | Style names, property labels |
| Secondary | 12px Regular | Usage counts, hints |
| Caption | 11px Regular | Empty-state help, timestamps |
| Numeric / mono | 12px Consolas | Token values, adjustment numbers |

### Spacing — 4px base grid
`4 · 8 · 12 · 16 · 24`. Row height 32px standard, 40px for style rows carrying a preview swatch. Panel padding 12px.

### Density
Compact. Task panes are narrow and users work in them all day — favour information density over generous whitespace, but never below 32px touch targets.

---

## 3. Information architecture

```
┌─ StyleSmith ──────────────────┐
│  [ Styles ] [ Tokens ]        │  ← segmented control, persistent
├───────────────────────────────┤
│  🔍 Search                     │
├───────────────────────────────┤
│  ⚠ 14 unlinked shapes  [Scan] │  ← contextual banner, dismissible
├───────────────────────────────┤
│                               │
│  ▸ style / token rows         │
│                               │
├───────────────────────────────┤
│  [ + New from selection ]     │  ← persistent footer action
└───────────────────────────────┘
```

Two top-level views, one modal layer (preview/confirm), one detail drill-down.

---

## 4. Styles view

### 4.1 Style row
```
┌──────────────────────────────────┐
│ ▉  Takeaway              47  ›   │
│ ▔  Aa 14 · ▢ navy · ⌒ 0.15       │
└──────────────────────────────────┘
```
- **Swatch (left):** a live mini-preview rendering the style's fill, border and corner radius. The corner radius in the swatch is deliberate — it advertises the geometry layer at a glance.
- **Name**
- **Usage count** — number of shapes carrying it in the deck
- **Layer summary line** — compact icons showing which of the four layers this style defines. A style with no text layer simply omits `Aa`.
- **Chevron** → detail

### 4.2 Style detail
Collapsible sections, one per layer. Undefined layers collapse to a single "Add text layer" row rather than showing empty fields.

```
‹ Takeaway                    ⋯

▾ TEXT
  Font        [🔗 font-body      ]
  Size        [🔗 size-body      ]
  Colour      [   #FFFFFF        ]
  Spacing     [   1.2            ]

▾ SHAPE
  Fill        [🔗 brand-primary  ]
  Border      [🔗 stroke-thin    ]

▾ GEOMETRY
  Corner      [🔗 radius-md  0.15]
  ⓘ Applies to rounded rectangles

▸ TABLE  (not defined)

──────────────────────────────
[ Apply to selection ]
[ Apply to all… ]
[ Update from selection ]
[ Select all instances ]
```

**The binding control** is the most important element in the product.
- **🔗 chip = token-bound.** Shows the token name, tinted with the accent colour.
- **Plain field = literal.** Shows the raw value.
- Clicking either opens a picker: choose a token, enter a literal, or **"Create token from this value"** (S30).
- Resolved value shown in muted text beside a token chip, so bound values are never opaque.

`⋯` menu: Rename · Duplicate · Delete · Export.

### 4.3 Empty state — first run
The single most important screen in the product, because every existing deck opens here.

```
      ▢▢▢
   No styles yet

  This deck has 312 shapes.
  StyleSmith can find groups
  that already share formatting.

  [ Scan this deck ]

  or [ Create from selection ]
```

---

## 5. Tokens view

### 5.1 Grouped list
```
COLOUR
  ▉ brand-primary    accent1 🎨   4 styles
  ▉ surface-muted    #F5F6F7      2 styles

TYPE
  Aa font-body       Body font 🎨  6 styles
  14 size-body       14pt          6 styles

GEOMETRY
  ⌒ radius-md        0.15          3 styles
  ⌒ radius-lg        0.28          1 style

STROKE
  ─ stroke-thin      1.5pt         4 styles
```

- **🎨 badge = bound to the deck theme** (S28). Communicates "this follows the deck, and survives without StyleSmith installed."
- **Usage count** is per style, not per shape, at list level — shape counts appear on edit.

### 5.2 Editing a token — the cascade preview
This is where the product's power becomes visible, so the confirmation must be generous rather than terse.

```
‹ brand-primary

Value   ( ) Literal      [#1A3A6B]
        (•) Theme colour [accent1 ▾]  🎨
        ( ) Another token [        ]

──────────────────────────────
Changing this affects:

  Takeaway         47 shapes
  KPI Card         28 shapes
  Section Header   12 shapes
  Quote Block       6 shapes
                   ───────────
                    93 shapes across 31 slides

[ Preview ]  [ Apply ]
```

Editing one token showing "93 shapes across 31 slides" **is the demo.** It should feel slightly powerful and slightly dangerous — which is exactly why Preview sits next to Apply.

---

## 6. Apply & preview modal

Triggered by any bulk action.

```
Apply "Takeaway"

SCOPE
  ( ) Current slide          3 shapes
  ( ) Selected slides       11 shapes
  (•) Entire deck           47 shapes

LAYERS
  [✓] Text     [✓] Shape
  [✓] Geometry [ ] Table

⚠ 4 shapes have local overrides.
  (•) Preserve them
  ( ) Reset to style

──────────────────────────────
Will change 47 shapes across 23 slides

[ Cancel ]              [ Apply ]
```

- Scope defaults to the **last used scope**, not always deck-wide
- Layer checkboxes implement selective application (S12)
- The override notice only appears when overrides exist — no permanent clutter
- The summary line is the safety promise, always immediately above the button

### 6.1 Progress state
Chunked execution is architectural (S11), so progress is a first-class state, not a spinner.

```
Applying "Takeaway"…
████████████░░░░░░░░  240 / 420

Slide 31 of 60          [ Cancel ]
```

Cancel must stop cleanly at a chunk boundary and report what was completed — a half-applied sweep the user knows about is recoverable; a silent one is not.

### 6.2 Completion
```
✓ Updated 47 shapes across 23 slides
  Ctrl+Z to undo
```
The undo hint appears because native undo works (confirmed) and most users won't expect it to.

---

## 7. The adoption flow (S15)

The product's real onboarding. Three steps.

**Step 1 — Scan**
```
Scanning 312 shapes…
```

**Step 2 — Clusters found**
```
Found 5 groups of matching shapes

┌────────────────────────────┐
│ ▉  14 shapes               │
│    Navy fill · 1.5pt · 0.15│
│    [ Name this style… ]    │
│    ☐ Include              │
└────────────────────────────┘
┌────────────────────────────┐
│ ▉  9 shapes                │
│    White fill · no border  │
│    [ Name this style… ]    │
│    ☐ Include              │
└────────────────────────────┘

    [ Create 2 styles ]
```

**Step 3 — Near matches**
```
3 shapes nearly match "Takeaway"

  Slide 12  ⌒ 0.18 (style: 0.15)   [Link] [Skip]
  Slide 24  fill #1A3A6C            [Link] [Skip]
  Slide 41  ⌒ 0.12, 13pt            [Link] [Skip]

  [ Link all and normalise ]
```

Showing the *specific* deviation for each near match is what makes this trustworthy rather than magic. It's also the same comparison BrandGuard performs — one engine, two products.

---

## 8. Contextual states

| State | Treatment |
|---|---|
| Unlinked shapes present | Dismissible banner with count and Scan action |
| Nothing selected, action needs selection | Action disabled with tooltip, never a modal error |
| Style applied to incompatible shape type | Inline notice after apply: "Geometry skipped on 3 shapes (not rounded rectangles)" |
| Token unresolvable | Token chip shows warning tint; style still renders via cached literal (§10.5) |
| Grouped shapes in scope | If group traversal proves impossible: "12 shapes inside groups were skipped" — must be stated, never silent |
| Deck has no styles | Empty state (§4.3) |
| Long operation | Progress with chunk-accurate counts (§6.1) |

---

## 9. Keyboard

| Shortcut | Action |
|---|---|
| `Ctrl+Alt+1…9` | Apply style 1–9 to selection (mirrors Word's Heading shortcuts) |
| `Ctrl+Shift+S` | Focus the panel search field |
| `Enter` on a style row | Apply to selection |
| `Esc` | Close detail / cancel modal |
| `Tab` order | Search → banner → list → footer, with a visible focus ring throughout |

---

## 10. Accessibility

- All Fluent tokens carry compliant contrast in light, dark and high-contrast themes — a further reason not to hard-code colours
- Never encode meaning in colour alone: the token chip uses the 🔗 icon plus tint; the theme badge uses 🎨 plus a text label
- Full keyboard operability, no mouse-only paths
- Screen reader labels on every swatch: "Takeaway style, navy fill, 47 shapes"
- Respect reduced-motion preferences on panel transitions

---

## 11. Open UI questions

1. **Style preview swatch fidelity** — how faithfully can a 24px swatch render fill, border and corner radius? A radius rendered wrong undersells the differentiator.
2. **Token picker density** — with 40+ tokens, does the picker need type-filtering or search?
3. **Does the token view need a "unused tokens" filter** in v1, or defer with style inventory (S17)?
4. **Near-match threshold** — how close is "nearly matches"? Needs a tunable tolerance per layer, and possibly a user-facing sensitivity control.
5. **Panel width** — task panes are resizable; does the layout need a two-column mode past ~450px, or stay single-column always?
