/**
 * Deck-level storage for the StyleSmith document, plus the central dirty-flag
 * guard.
 *
 * Storage medium (Phase 0 spike): Office.context.document.settings. PowerPoint
 * has NO document-level customXmlParts — that Common API is Word/Excel only.
 * Crucially, `settings.saveAsync` self-persists across close/reopen WITHOUT a
 * dirtying change, so the document save itself needs no dirty guard.
 *
 * The dirty guard exists ONLY for tag-only writes (linking a shape with no
 * visual change), which are otherwise dropped on web. Implemented once here,
 * never ad hoc at call sites (CLAUDE.md).
 *
 * On load, raw settings go through the migration harness — the document may have
 * been written by an older build or hand-edited.
 */
import { migrateToLatest, type MigrationError } from "../core/schema/migrations";
import { ok, err, type Result, type StyleSmithDocument } from "../core/schema/types";
import { officeErrorMessage } from "./context";

/** The settings key holding the serialised StyleSmith document JSON. */
export const DOCUMENT_SETTING_KEY = "stylesmith:document";

export type StorageError =
  | { readonly code: "no-host" }
  | { readonly code: "office-error"; readonly message: string }
  | { readonly code: "invalid-document"; readonly issues: string[] };

function settings(): Office.Settings | null {
  if (typeof Office === "undefined") return null;
  return Office.context?.document?.settings ?? null;
}

function saveSettings(store: Office.Settings): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    store.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(result.error);
    });
  });
}

/**
 * Load the document from the deck, or `null` if none is stored yet. Invalid or
 * unreadable data is a typed error, never a throw — a corrupt blob must not
 * crash the panel.
 */
export async function loadDocument(): Promise<Result<StyleSmithDocument | null, StorageError>> {
  const store = settings();
  if (!store) return err({ code: "no-host" });

  const raw: unknown = store.get(DOCUMENT_SETTING_KEY);
  if (raw === undefined || raw === null) return ok(null);

  let parsed: unknown;
  try {
    parsed = typeof raw === "string" ? JSON.parse(raw) : raw;
  } catch {
    return err({ code: "invalid-document", issues: ["stored value is not valid JSON"] });
  }

  const migrated = migrateToLatest(parsed);
  if (!migrated.ok) {
    return err({ code: "invalid-document", issues: describeMigrationError(migrated.error) });
  }
  return ok(migrated.value);
}

/**
 * Persist the document. `settings.saveAsync` self-persists, so this needs no
 * dirtying pair (verified in Phase 0).
 */
export async function saveDocument(
  doc: StyleSmithDocument,
): Promise<Result<void, StorageError>> {
  const store = settings();
  if (!store) return err({ code: "no-host" });

  store.set(DOCUMENT_SETTING_KEY, JSON.stringify(doc));
  try {
    await saveSettings(store);
    return ok(undefined);
  } catch (e) {
    return err({ code: "office-error", message: officeErrorMessage(e) });
  }
}

/** Remove the stored document (used by tests/reset). */
export async function clearDocument(): Promise<Result<void, StorageError>> {
  const store = settings();
  if (!store) return err({ code: "no-host" });
  store.remove(DOCUMENT_SETTING_KEY);
  try {
    await saveSettings(store);
    return ok(undefined);
  } catch (e) {
    return err({ code: "office-error", message: officeErrorMessage(e) });
  }
}

/* ── The dirty-flag guard (tag-only writes) ──────────────────────────────── */

/**
 * Mark the document modified without any visible change, so a metadata-only
 * (tag-only) write is autosaved on web. Re-sets a shape's `left` to its own
 * value. Call within the same PowerPoint.run as the tag write, then sync.
 *
 * ⚠️ This is the ONE place this trick lives. Never inline it at call sites.
 */
export async function ensureDocumentDirty(ctx: PowerPoint.RequestContext): Promise<void> {
  const shape = await firstShape(ctx);
  if (!shape) return; // empty deck — nothing to nudge (and nothing tagged either)
  shape.load("left");
  await ctx.sync();
  shape.left = shape.left; // no-op value change; dirties the document
  await ctx.sync();
}

async function firstShape(ctx: PowerPoint.RequestContext): Promise<PowerPoint.Shape | null> {
  const slides = ctx.presentation.slides;
  slides.load("items");
  await ctx.sync();
  const first = slides.items[0];
  if (!first) return null;
  first.shapes.load("items/id");
  await ctx.sync();
  return first.shapes.items[0] ?? null;
}

function describeMigrationError(error: MigrationError): string[] {
  switch (error.code) {
    case "missing-version":
      return ["document has no schemaVersion"];
    case "from-newer":
      return [`document is from a newer version (v${error.version}) — update the add-in`];
    case "no-migration":
      return [`no migration path from v${error.version}`];
    case "invalid":
      return error.issues;
  }
}
