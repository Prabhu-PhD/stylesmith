import { useEffect, useRef } from "react";
import { useStore } from "./state/store";

/**
 * Global keyboard shortcuts (UX spec §9). Reads the store imperatively so it
 * doesn't re-subscribe on every state change.
 *
 *   Ctrl/⌘+Alt+1–9   apply style N to the current selection
 *   Ctrl/⌘+Shift+S   focus the panel search
 *   Esc              close the top-most surface (modal → adoption → detail)
 */
export function useKeyboardShortcuts(focusSearch: () => void): void {
  const focusRef = useRef(focusSearch);
  focusRef.current = focusSearch;

  useEffect(() => {
    const handler = (e: KeyboardEvent): void => {
      const st = useStore.getState();

      if (e.key === "Escape") {
        if (st.applyStyleId) st.closeApply();
        else if (st.adoptionOpen) st.cancelAdoption();
        else if (st.operation?.phase === "done") st.dismissOperation();
        else if (st.selectedTokenId) st.selectToken(null);
        else if (st.selectedStyleId) st.selectStyle(null);
        return;
      }

      const mod = e.ctrlKey || e.metaKey;

      if (mod && e.shiftKey && (e.key === "s" || e.key === "S")) {
        e.preventDefault();
        st.setView("styles");
        st.selectStyle(null);
        focusRef.current();
        return;
      }

      if (mod && e.altKey && /^[1-9]$/.test(e.key)) {
        const style = st.doc?.styles[Number(e.key) - 1];
        if (style) {
          e.preventDefault();
          void st.applyToSelection(style.id);
        }
      }
    };

    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, []);
}
