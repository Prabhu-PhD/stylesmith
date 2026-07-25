/**
 * Office theme reading. Office JS is confined to office/, so the UI theme hook
 * (ui/theme/useOfficeTheme) asks us for the host's light/dark mode rather than
 * touching the Office globals itself.
 *
 * The host does not expose a "high contrast" flag here; the hook detects that
 * from the `forced-colors` CSS media query. We only classify light vs dark,
 * from the luminance of the task pane body background the host reports.
 */
import type { ThemeSlot } from "../core/schema/types";

export type OfficeThemeMode = "light" | "dark" | "highContrast";

/* ── Deck theme palette (for resolving `themeColor` bindings to a hex) ─────── */

/**
 * Resolve a deck theme slot (accent1, dk1, …) to its current hex colour.
 *
 * ⚠️ OPEN QUESTION (flagged for Phase 3): PowerPoint Office JS 1.1–1.10 exposes
 * no documented API for reading the deck's *theme palette*. `Office.context
 * .officeTheme` is the task-pane (Office UI) theme, not the presentation theme.
 * Until a read path is found, this returns undefined, and `resolveValue` falls
 * back to the binding's cached literal for preview. This does NOT affect
 * correctness of applying a theme colour to a shape (a separate Phase 3 concern)
 * — only the panel's ability to preview the live slot colour.
 */
export function readDeckThemeColor(_slot: ThemeSlot): string | undefined {
  return undefined;
}

/** A theme resolver suitable for core's ResolveContext. Best-effort (see above). */
export function createDeckThemeResolver(): (slot: ThemeSlot) => string | undefined {
  return readDeckThemeColor;
}

/** Relative luminance (0 = black, 1 = white) of an #rgb / #rrggbb colour. */
function luminance(hex: string): number | null {
  const m = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(hex.trim());
  if (!m || !m[1]) return null;
  // Expand 3-digit shorthand (#abc -> #aabbcc), then parse as one integer.
  const h = m[1].length === 3 ? m[1].replace(/(.)/g, "$1$1") : m[1];
  const int = parseInt(h, 16);
  const r = ((int >> 16) & 0xff) / 255;
  const g = ((int >> 8) & 0xff) / 255;
  const b = (int & 0xff) / 255;
  // Perceptual weighting; good enough to split light from dark task-pane themes.
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/**
 * Light or dark, per the host's reported task-pane background — or null when the
 * host theme is unavailable (e.g. running in a plain browser during dev).
 */
export function readOfficeThemeMode(): "light" | "dark" | null {
  if (typeof Office === "undefined" || !Office.context) return null;
  const bg = Office.context.officeTheme?.bodyBackgroundColor;
  if (!bg) return null;
  const lum = luminance(bg);
  if (lum === null) return null;
  return lum < 0.5 ? "dark" : "light";
}
