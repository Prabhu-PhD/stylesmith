# StyleSmith — Acceptance Criteria (v1 / P0)
**v1.0 · 23 July 2026**

Testable criteria for every P0 feature. A feature is done when all of its criteria pass on **PowerPoint web, Windows desktop and Mac desktop**.

Format: **Given** / **When** / **Then**.

---

## S1 — Create style from selection

**AC1.1** Given a formatted shape is selected, when the user chooses "New from selection" and enters a name, then a style is created capturing all four layers, and the shape is linked to it.

**AC1.2** Given a shape with no adjustment handles (e.g. a plain rectangle), when a style is created from it, then the geometry layer is `null` rather than an empty array.

**AC1.3** Given a style name that already exists, when the user attempts to save, then a conflict is surfaced with rename/replace options — never a silent overwrite.

**AC1.4** Given a style is created, when the document is closed and reopened, then the style persists with all layers intact. *(Verifies the dirty-flag guard.)*

**AC1.5** Given nothing is selected, when the user opens the create action, then it is disabled with an explanatory tooltip — not an error modal.

---

## S2 — Apply style

**AC2.1** Given a style and one selected shape, when applied, then every defined layer is written to the shape and the shape is tagged with the style's GUID.

**AC2.2** Given multiple shapes are selected, when a style is applied, then all selected shapes receive it in a single operation.

**AC2.3** Given a style with a geometry layer and a selected shape with **zero** adjustments, when applied, then text and shape layers apply, geometry is skipped, and the user is told: "Geometry skipped on N shapes (no adjustable handles)."

**AC2.4** Given a style with a geometry layer defining more adjustments than the target shape has, when applied, then only the available indices are written and no error is thrown.

**AC2.5** Given a shape already carrying style A, when style B is applied, then the tag is replaced (not duplicated) and A's usage count decreases by one.

---

## S3 — Modify → Apply to all ⭐

**AC3.1** Given a style used by N shapes across M slides, when the user opens "Apply to all", then the confirmation states exactly "Will change N shapes across M slides" with correct counts.

**AC3.2** Given the user confirms, when the sweep runs, then all N shapes update and progress reports chunk-accurate counts throughout.

**AC3.3** Given a sweep is in progress, when the user cancels, then it stops at the next chunk boundary and reports precisely how many shapes were changed.

**AC3.4** Given 420+ tagged shapes on the 60-slide test deck, when applied deck-wide, then the operation **completes without hanging the host**.

**AC3.5** Given shapes with local overrides and "Preserve overrides" selected, when the sweep runs, then overridden properties are left untouched and all others update.

**AC3.6** Given a completed sweep, when the user presses Ctrl+Z, then the change is reverted by PowerPoint's native undo.

**AC3.7** Given scope is "Current slide", when applied, then only shapes on the active slide change — the count in the confirmation reflects that scope, not the deck.

---

## S4 — Update style to match selection

**AC4.1** Given a shape linked to a style and manually reformatted, when the user chooses "Update style to match", then the style definition absorbs the shape's current formatting.

**AC4.2** Given a property in the style was bound to a token, when the style is updated from a shape whose value differs, then the user is asked whether to unbind to a literal or update the token itself — never silently broken.

**AC4.3** Given the style is updated, when other shapes carry it, then they are **not** changed until "Apply to all" is explicitly run.

---

## S5 — Select all instances

**AC5.1** Given a style used by N shapes, when "Select all instances" is chosen, then all N are selected in PowerPoint.

**AC5.2** Given instances span multiple slides, when selected, then the behaviour is defined and communicated (PowerPoint cannot select across slides — navigate to the first and report "N shapes on M slides; showing slide X").

---

## S6 — Detach / clear

**AC6.1** Given a linked shape, when "Detach" is chosen with "keep formatting", then the tag is removed and visual formatting is unchanged.

**AC6.2** Given a linked shape, when "Detach" is chosen with "reset formatting", then the tag is removed and formatting returns to PowerPoint defaults.

**AC6.3** Given a detach operation writes only metadata, when the document is closed and reopened, then the detachment persisted. *(Dirty-flag guard.)*

---

## S7 — Rename / delete style

**AC7.1** Given a style used by N shapes, when renamed, then all N remain linked. *(Verifies GUID keying.)*

**AC7.2** Given a style used by N shapes, when deletion is attempted, then a warning states the usage count and offers detach-and-delete or cancel.

**AC7.3** Given a style is deleted, when shapes previously carried it, then those shapes retain their formatting and become unlinked orphans — never blanked.

---

## S9 — Preview before apply

**AC9.1** Given any bulk action, when triggered, then a confirmation appears stating affected shape and slide counts before anything changes.

**AC9.2** Given the confirmation is shown, when the user cancels, then no document change occurs whatsoever.

**AC9.3** Given counts are displayed, when compared against actual results, then they match exactly.

---

## S10 — Scope control

**AC10.1** Given the apply modal, when opened, then scope options show live counts for current slide, selected slides and entire deck.

**AC10.2** Given the user selected a scope previously, when the modal reopens, then that scope is pre-selected.

**AC10.3** Given "Selected slides" with no multi-selection, when the modal opens, then that option is disabled with an explanation.

---

## S11 — Chunked execution and progress

**AC11.1** Given an operation affecting more than one chunk, when it runs, then progress updates at least once per chunk with accurate counts.

**AC11.2** Given an operation exceeding ~1 second, when it runs, then a progress indicator is visible for its entire duration.

**AC11.3** Given cancellation mid-operation, when it stops, then the document is left in a consistent state and the user is told what was completed.

---

## S12 — Selective layer application

**AC12.1** Given a style with four layers, when the user deselects Geometry and applies, then text, shape and table apply and geometry does not.

**AC12.2** Given all layers are deselected, when the user attempts to apply, then Apply is disabled.

---

## S13 — Preserve local overrides

**AC13.1** Given a linked shape with a manually changed fill, when the style is applied with preserve enabled, then the fill remains overridden and all other properties update.

**AC13.2** Given the same shape with "Reset to style", when applied, then the override is cleared and the style value takes effect.

**AC13.3** Given overrides exist in scope, when the apply modal opens, then their count is shown with the preserve/reset choice — and the notice is absent when no overrides exist.

---

## S14 — Dirty-flag guard ⚠️

**AC14.1** Given an operation writing **only** metadata (tag-only link, style rename, token rename), when the document is closed and reopened **without any other edit**, then the change persisted.

**AC14.2** Given AC14.1 on PowerPoint web specifically, then it passes there too. *(This is where the original failure was observed.)*

**AC14.3** Given the dirty guard runs, when the user inspects the deck, then no visible or unintended change was introduced to any shape.

---

## S15 — Adoption engine ⭐

**AC15.1** Given a deck with no styles, when opened, then the empty state offers to scan.

**AC15.2** Given a scan on a deck with groups of identically formatted shapes, when it completes, then those groups are clustered with accurate counts and a representative preview.

**AC15.3** Given clusters are shown, when the user names one and confirms, then a style is created and **all** shapes in the cluster are linked in one operation.

**AC15.4** Given shapes that nearly match an existing style, when surfaced, then each shows the **specific** deviating property and value — e.g. `⌒ 0.18 (style: 0.15)`.

**AC15.5** Given a near match, when "Link and normalise" is chosen, then the shape is linked and updated to the style's values.

**AC15.6** Given a near match, when "Link as-is" is chosen, then the shape is linked and its deviation is retained as a local override.

**AC15.7** Given shapes pasted from another deck, when detected, then the unlinked banner appears with a count and a scan action.

**AC15.8** Given a scan on the 2,520-shape test deck, when it runs, then it completes without hanging and within an acceptable time (measure and record).

---

## S16 — Cross-deck style import

**AC16.1** Given another .pptx containing StyleSmith styles, when imported, then its styles and tokens are added to the current deck.

**AC16.2** Given an imported style name collides with an existing one, when the conflict is surfaced, then rename / overwrite / keep-both are offered.

**AC16.3** Given imported styles reference tokens, when imported, then those tokens are imported too and references resolve correctly.

**AC16.4** Given a file that is not a StyleSmith deck, when import is attempted, then it fails gracefully with a clear message.

---

## S25 — Token library

**AC25.1** Given tokens exist, when the Tokens view opens, then they are grouped by type with name, resolved value and usage count.

**AC25.2** Given a token bound to a theme colour, when displayed, then it carries a visible theme badge.

**AC25.3** Given a token is created, when the document is closed and reopened, then it persists.

---

## S26 — Bind property to token

**AC26.1** Given a style property, when the user opens its binding control, then existing tokens of the matching type are offered, plus literal entry and "create token from this value".

**AC26.2** Given a bound property, when displayed, then it shows a token chip with the token's name **and** its resolved value.

**AC26.3** Given a token of the wrong type (e.g. a colour token for a radius), when the picker opens, then it is not offered.

---

## S27 — Edit token → cascade ⭐

**AC27.1** Given a token used by multiple styles, when opened for edit, then every affected style is listed with its shape count and a correct total.

**AC27.2** Given the token value is changed and applied, when the sweep completes, then every shape carrying every affected style reflects the new value.

**AC27.3** Given a change is previewed, when cancelled, then no document change occurs.

**AC27.4** Given a token referenced in an alias chain, when a cycle would be created, then the edit is rejected with an explanation.

**AC27.5** Given a token is deleted while styles reference it, when those styles are next applied, then they fall back to the cached literal and the affected styles are flagged — no shape is blanked.

---

## S28 — Theme colour binding

**AC28.1** Given a colour token, when bound to a theme slot, then it resolves to that slot's current colour.

**AC28.2** Given a theme-bound token, when the deck's theme changes, then shapes update **without StyleSmith running**.

**AC28.3** Given a theme-bound style is applied, when the deck is opened on a machine **without** the add-in, then the formatting is correct and theme-linked.

---

## S29 — Token usage count

**AC29.1** Given a token, when displayed in the list, then its style usage count is accurate.

**AC29.2** Given a token opened for edit, when the cascade preview renders, then both style and shape counts are accurate.

---

## Cross-cutting

### Persistence
**ACX.1** Every create, edit and delete of a style or token survives close and reopen on web, Windows and Mac.

### Degradation
**ACX.2** Given a deck opened without the add-in installed, then all formatting renders correctly and no error is shown.
**ACX.3** Given that deck is edited and saved without the add-in, when reopened with it, then styles and tokens are intact. *(If this fails, it must be documented as a known limitation.)*

### Performance
**ACX.4** Deck scan of 2,520 shapes completes in under 5 seconds.
**ACX.5** Sweep of 420 shapes completes without hanging, with a recorded throughput figure.

### Accessibility
**ACX.6** Every interactive element is reachable and operable by keyboard with a visible focus indicator.
**ACX.7** The panel renders correctly in light, dark and high-contrast Office themes.
**ACX.8** No information is conveyed by colour alone.

### Error handling
**ACX.9** No operation leaves the document in a partially applied state without telling the user.
**ACX.10** Office JS errors are surfaced in plain language, never as raw exception text.
