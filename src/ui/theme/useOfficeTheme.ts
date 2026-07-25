import { useEffect, useState } from "react";
import {
  webLightTheme,
  webDarkTheme,
  teamsHighContrastTheme,
  type Theme,
} from "@fluentui/react-components";
import { readOfficeThemeMode, type OfficeThemeMode } from "../../office/theme";

const THEMES: Record<OfficeThemeMode, Theme> = {
  light: webLightTheme,
  dark: webDarkTheme,
  highContrast: teamsHighContrastTheme,
};

const FORCED_COLORS = "(forced-colors: active)";
const PREFERS_DARK = "(prefers-color-scheme: dark)";

function matches(query: string): boolean {
  return typeof window !== "undefined" && !!window.matchMedia && window.matchMedia(query).matches;
}

/**
 * Resolve the effective mode. Precedence:
 *   1. Windows High Contrast / forced-colors  -> highContrast
 *   2. The host's reported Office theme        -> light | dark
 *   3. OS colour-scheme preference (dev tab)   -> light | dark
 */
function detect(): OfficeThemeMode {
  if (matches(FORCED_COLORS)) return "highContrast";
  const office = readOfficeThemeMode();
  if (office) return office;
  return matches(PREFERS_DARK) ? "dark" : "light";
}

/**
 * Follows the user's Office theme (light / dark / high-contrast) and re-resolves
 * on live changes. The Office web host propagates theme switches to the task-pane
 * iframe as CSS media-query changes, so those double as our live signal; a
 * visibilitychange re-check covers switches made while the pane was hidden.
 */
export function useOfficeTheme(): { mode: OfficeThemeMode; theme: Theme } {
  const [mode, setMode] = useState<OfficeThemeMode>(detect);

  useEffect(() => {
    if (typeof window === "undefined" || !window.matchMedia) return;
    const update = (): void => setMode(detect());
    const mqls = [window.matchMedia(FORCED_COLORS), window.matchMedia(PREFERS_DARK)];
    mqls.forEach((mql) => mql.addEventListener("change", update));
    document.addEventListener("visibilitychange", update);
    return () => {
      mqls.forEach((mql) => mql.removeEventListener("change", update));
      document.removeEventListener("visibilitychange", update);
    };
  }, []);

  return { mode, theme: THEMES[mode] };
}
