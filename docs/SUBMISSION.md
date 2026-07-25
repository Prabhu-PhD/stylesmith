# StyleSmith — AppSource submission checklist

Phase 7 deliverable. Tracks what is code-complete vs. what needs real values, a
host, or a human before AppSource submission.

## Code-complete (this repo)

- [x] XML add-in manifest with **add-in commands** (ribbon button → task pane),
      `manifest/manifest.xml` (prod) and `manifest/manifest-localhost.xml` (dev).
- [x] Privacy policy and support pages — `public/privacy.html`, `public/support.html`.
- [x] Ribbon-button command file — `public/commands.html`.
- [x] Logo source — `public/assets/logo.svg`.
- [x] Accessibility: Fluent tokens (light/dark/high-contrast), aria labels on
      swatches/rows/controls, no meaning by colour alone, reduced-motion, keyboard
      shortcuts (Ctrl+Alt+1–9, Ctrl+Shift+S, Esc), error boundary.
- [x] Bundle: rare surfaces code-split (`React.lazy`); Fluent v9 is the size floor
      (~200 kB gzip main chunk).
- [x] Vercel static hosting config — `vercel.json`.

## Hosting — DONE

- [x] **Hosted on GitHub Pages**: <https://prabhu-phd.github.io/stylesmith/>.
      Repo <https://github.com/Prabhu-PhD/stylesmith>, auto-deployed by
      `.github/workflows/deploy.yml` on every push to `main`. Vite `base` is
      `/stylesmith/` for builds; the prod manifest points at this origin.
      (Verified live: taskpane/support/commands/icons all return 200 and the
      panel mounts from the Pages URL.)
- [x] **PNG icons** at `public/assets/icon-16|32|64|80.png` (generated from
      `logo.svg`), served under `/stylesmith/assets/`.

## Needs real values before submission (business)

- [ ] **Production Id.** Confirm/replace the `<Id>` GUID as the canonical AppSource id.
- [ ] **Contact address** in `privacy.html` / `support.html` (currently a placeholder).
- [ ] **AppSource listing assets**: screenshots (1366×768 recommended), a 300×300
      store logo, category, and long description.
- [ ] Optional: move to a **custom domain** (then update the manifest URLs and set
      Vite `base` back to `/`).

## Needs a host / validation run (cannot be done headlessly)

- [ ] `npx office-addin-manifest validate manifest/manifest.xml` passes.
- [ ] Sideload and run the full manual protocol in `docs/QA-PROTOCOL.md` on
      **PowerPoint web, Windows desktop, and Mac desktop** (the phase-7 gate).
- [ ] Confirm every earlier phase's host-pending gate on desktop + web:
      - Phase 0 spikes (group traversal ✓, settings persistence ✓, Ctrl+D ✓) re-verified on **desktop**.
      - Phase 2 storage round-trip (write → reopen → identical).
      - Phase 3 apply sweep: 400+ shapes without hanging; record throughput; tune chunk size.
      - Phase 4 hero loop end-to-end (create → apply → modify → apply-to-all).
      - Phase 5 token cascade: edit token → accurate counts → apply → shapes update.
      - Phase 6 adoption: open the 60-slide fixture cold → scan → clusters match → create → link.
- [ ] Native undo (Ctrl+Z) reverts bulk applies (AC3.6).

## Known limitations to disclose in the listing

- **Theme-colour cascade (AC28.2/28.3)** is not delivered in v1: PowerPoint JS 1.x
  exposes no deck-theme palette read and no theme-slot fill write, so `themeColor`
  bindings apply as a resolved hex and do not cascade without the add-in. Revisit
  when the API allows.
- **Deep text** (line spacing, space before/after, alignment, bullets) is reported
  as deferred, not applied — the "textFrame depth" open question.
- **Table layer** is schema-ready but not yet editable (Phase-5-style follow-up).
- Adoption fingerprints top-level, unlinked, non-group shapes; group descent is a
  known follow-up (now feasible per the Phase 0 spike).
