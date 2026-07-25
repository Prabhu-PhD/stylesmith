import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./ui/App";

function mount(): void {
  const el = document.getElementById("root");
  if (!el) throw new Error("Root element #root not found");
  createRoot(el).render(
    <StrictMode>
      <App />
    </StrictMode>,
  );
}

// Inside PowerPoint, wait for Office to be ready before touching Office APIs.
// In a plain browser tab, office.js loads but Office.onReady() never resolves
// (no host handshake), so race it with a short timeout to mount either way.
// Promise.race resolves once, so mount runs exactly once.
const ready =
  typeof Office !== "undefined" && Office.onReady ? Office.onReady() : Promise.resolve();
const fallback = new Promise<void>((resolve) => setTimeout(resolve, 1500));
void Promise.race([ready, fallback]).then(mount);
