# StyleSmith — Manual QA Protocol
**v1.0 · Run before every release. Office JS cannot be meaningfully automated end-to-end, so this protocol is the release gate.**

---

## Environment matrix

Every suite runs on all three hosts. Behaviour genuinely differs between them — the API surface, autosave behaviour and undo semantics are not identical.

| Host | Notes |
|---|---|
| **PowerPoint web** (Edge) | Autosave. Where the dirty-flag failure was originally found. |
| **PowerPoint desktop, Windows** | Explicit save. Primary buyer platform. |
| **PowerPoint desktop, Mac** | Explicit save. WebKit webview — rendering and API differences possible. |

Also verify the panel in **light**, **dark** and **high-contrast** Office themes at least once per release.

**Fixture:** `fixtures/stylesmith-perf-test-deck.pptx` (60 slides, 2,520 shapes) plus at least one real-world deck.

---

## Suite 1 — Installation and load

| # | Step | Expected | ✓ |
|---|---|---|---|
| 1.1 | Sideload the manifest | Add-in appears in the ribbon | ☐ |
| 1.2 | Open the panel | Loads within 2s, no blocked/error icon | ☐ |
| 1.3 | Check the theme | Panel matches the current Office theme | ☐ |
| 1.4 | Switch Office to dark, reopen | Panel follows without a restart | ☐ |
| 1.5 | Switch to high contrast | All text and controls remain legible | ☐ |
| 1.6 | Resize the task pane narrow and wide | Layout holds; no horizontal scroll | ☐ |
| 1.7 | Close and reopen the panel | State restores sensibly | ☐ |

---

## Suite 2 — Style CRUD

| # | Step | Expected | ✓ |
|---|---|---|---|
| 2.1 | Format a rounded rectangle, create a style | All four layers captured | ☐ |
| 2.2 | Inspect the style detail | Geometry shows the actual corner radius | ☐ |
| 2.3 | Create a style from a plain rectangle | Geometry layer is null, not empty | ☐ |
| 2.4 | Create a style with a duplicate name | Conflict offered, no silent overwrite | ☐ |
| 2.5 | Rename a style used by shapes | All shapes remain linked | ☐ |
| 2.6 | Duplicate a style | Independent copy, new GUID | ☐ |
| 2.7 | Delete a style in use | Warning with usage count | ☐ |
| 2.8 | Confirm the deletion | Shapes keep formatting, become orphans | ☐ |
| 2.9 | Attempt create with nothing selected | Disabled with tooltip, not an error modal | ☐ |

---

## Suite 3 — Apply and scope

| # | Step | Expected | ✓ |
|---|---|---|---|
| 3.1 | Apply a style to one shape | All defined layers applied | ☐ |
| 3.2 | Apply to a multi-selection | All selected shapes updated | ☐ |
| 3.3 | Apply a geometry style to a plain rectangle | Text/shape applied, geometry skipped, **user informed** | ☐ |
| 3.4 | Apply style B to a shape carrying style A | Tag replaced; A's count decreases by one | ☐ |
| 3.5 | Open Apply to all | Counts stated and correct | ☐ |
| 3.6 | Set scope to current slide | Count reflects the slide only | ☐ |
| 3.7 | Set scope to selected slides | Count reflects that selection | ☐ |
| 3.8 | Cancel at the confirmation | **Zero** document changes | ☐ |
| 3.9 | Deselect the geometry layer, apply | Geometry untouched, others applied | ☐ |
| 3.10 | Deselect all layers | Apply disabled | ☐ |
| 3.11 | Apply with overrides present, preserve | Overrides survive, rest updates | ☐ |
| 3.12 | Apply with overrides present, reset | Overrides cleared | ☐ |
| 3.13 | Ctrl+Z after a bulk apply | Native undo reverts | ☐ |

---

## Suite 4 — Persistence ⚠️ CRITICAL

*This suite catches the failure mode that cost us a full debugging cycle. Run it in full, every release, on every host.*

| # | Step | Expected | ✓ |
|---|---|---|---|
| 4.1 | Create a style, save, close, reopen | Style present and intact | ☐ |
| 4.2 | Apply a style, save, close, reopen | Linkage intact; usage count correct | ☐ |
| 4.3 | **Rename a style only. Change nothing else.** Close, reopen | ⚠️ Rename persisted *(dirty-flag guard)* | ☐ |
| 4.4 | **Link a shape whose formatting already matches — no visual change.** Close, reopen | ⚠️ Linkage persisted *(dirty-flag guard)* | ☐ |
| 4.5 | **Create a token only. Change nothing else.** Close, reopen | ⚠️ Token persisted *(dirty-flag guard)* | ☐ |
| 4.6 | Repeat 4.3–4.5 **on PowerPoint web specifically** | All persist | ☐ |
| 4.7 | Duplicate a linked shape (Ctrl+D) | Behaviour matches documented expectation | ☐ |
| 4.8 | Copy a linked shape to another deck | Arrives unlinked; adoption banner appears | ☐ |
| 4.9 | Save on web, open on desktop | Styles and tokens intact | ☐ |
| 4.10 | Save on desktop, open on web | Styles and tokens intact | ☐ |
| 4.11 | Verify no unintended visual change from the dirty guard | Deck visually identical | ☐ |

---

## Suite 5 — Tokens and cascade

| # | Step | Expected | ✓ |
|---|---|---|---|
| 5.1 | Create a colour token | Appears grouped under Colour | ☐ |
| 5.2 | Bind a style fill to it | Chip shows name and resolved value | ☐ |
| 5.3 | Open the token picker for a radius property | Only geometry-type tokens offered | ☐ |
| 5.4 | Bind a token to a theme colour slot | Theme badge shown | ☐ |
| 5.5 | Change the deck theme | Theme-bound shapes update | ☐ |
| 5.6 | Open a token used by several styles | Cascade preview lists each with counts | ☐ |
| 5.7 | Verify cascade counts against reality | Exact match | ☐ |
| 5.8 | Change the token and apply | All affected shapes update | ☐ |
| 5.9 | Cancel a token edit | **Zero** document changes | ☐ |
| 5.10 | Attempt an alias cycle (a→b→a) | Rejected with explanation | ☐ |
| 5.11 | Delete a token in use, then apply the style | Falls back to cached literal; **no shape blanked**; style flagged | ☐ |
| 5.12 | Create a token from an existing literal value | Promotes correctly, binding replaces the literal | ☐ |
| 5.13 | Unbind a token to a literal | Value retained, binding removed | ☐ |

---

## Suite 6 — Adoption

| # | Step | Expected | ✓ |
|---|---|---|---|
| 6.1 | Open a real deck with no styles | Empty state offers a scan | ☐ |
| 6.2 | Run the scan | Clusters found with accurate counts | ☐ |
| 6.3 | Check a cluster preview | Representative and recognisable | ☐ |
| 6.4 | Name a cluster and create | Style created; **all** cluster shapes linked | ☐ |
| 6.5 | Review near matches | Each shows its **specific** deviation and value | ☐ |
| 6.6 | Link and normalise | Shape linked and updated | ☐ |
| 6.7 | Link as-is | Shape linked, deviation kept as an override | ☐ |
| 6.8 | Skip a near match | Remains unlinked | ☐ |
| 6.9 | Paste slides from another deck | Unlinked banner appears with a count | ☐ |
| 6.10 | Scan the 2,520-shape fixture | Completes without hanging; record the time | ☐ |
| 6.11 | Scan a **real** deck with natural drift | Clusters are sensible, not fragmented into dozens | ☐ |

---

## Suite 7 — Performance

Record the numbers; don't just tick the box. Trends across releases matter.

| # | Step | Target | Actual | ✓ |
|---|---|---|---|---|
| 7.1 | Panel cold start | < 2s | ______ | ☐ |
| 7.2 | Scan 2,520 shapes | < 5s | ______ | ☐ |
| 7.3 | Apply to 420 shapes | No hang | ______ | ☐ |
| 7.4 | Progress accuracy during 7.3 | Chunk-accurate throughout | | ☐ |
| 7.5 | Cancel mid-sweep | Stops cleanly, reports count | ______ | ☐ |
| 7.6 | Document state after cancel | Consistent, no partial mess | | ☐ |
| 7.7 | Panel memory after 30 min of use | No unbounded growth | ______ | ☐ |

---

## Suite 8 — Degradation without the add-in

| # | Step | Expected | ✓ |
|---|---|---|---|
| 8.1 | Open a styled deck with the add-in removed | Formatting renders correctly | ☐ |
| 8.2 | Check theme-bound shapes | Still follow the deck theme | ☐ |
| 8.3 | Edit and save without the add-in | No errors | ☐ |
| 8.4 | Reopen with the add-in | Styles and tokens intact *(or the limitation is documented)* | ☐ |
| 8.5 | Open a styled deck in Google Slides / Keynote | Degrades without corruption | ☐ |

---

## Suite 9 — Accessibility

| # | Step | Expected | ✓ |
|---|---|---|---|
| 9.1 | Navigate the whole panel by keyboard only | Every control reachable | ☐ |
| 9.2 | Check focus indicators | Visible on every focusable element | ☐ |
| 9.3 | Run the hero flow with keyboard only | Completable | ☐ |
| 9.4 | Test Ctrl+Alt+1…9 | Applies the correct styles | ☐ |
| 9.5 | Screen reader over style rows | Meaningful labels announced | ☐ |
| 9.6 | High-contrast theme | All state distinguishable | ☐ |
| 9.7 | Check colour-only signals | None — icons or text accompany every colour cue | ☐ |

---

## Suite 10 — Error handling

| # | Step | Expected | ✓ |
|---|---|---|---|
| 10.1 | Trigger an action with nothing selected | Disabled with tooltip, not an exception | ☐ |
| 10.2 | Apply a style to a picture or chart | Handled gracefully with a clear message | ☐ |
| 10.3 | Apply where the scope contains grouped shapes | Skipped count reported, never silent | ☐ |
| 10.4 | Corrupt the custom XML part by hand, reopen | Graceful failure, offer to reset, no crash | ☐ |
| 10.5 | Open a deck written by a newer schema version | Clear message, no data loss | ☐ |
| 10.6 | Force an Office JS error | Plain-language message, never raw exception text | ☐ |
| 10.7 | Disconnect the network mid-session | Panel remains functional (no backend in v1) | ☐ |

---

## Regression checklist — the four that must never break

These are the failures that would be unrecoverable for a user. Verify every release, no exceptions.

- ☐ **Metadata-only writes persist on web** (4.3, 4.4, 4.5, 4.6)
- ☐ **Bulk apply to 420+ shapes does not hang** (7.3)
- ☐ **A deleted token never blanks a shape** (5.11)
- ☐ **Cancel never leaves an undisclosed partial state** (7.5, 7.6)

---

## Sign-off

| Host | Tester | Date | Suites passed | Notes |
|---|---|---|---|---|
| Web (Edge) | | | | |
| Windows desktop | | | | |
| Mac desktop | | | | |

**Release blocked unless:** all four regression items pass on all three hosts, and Suite 4 passes in full.
