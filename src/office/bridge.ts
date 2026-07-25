/**
 * The Office bridge — the single surface the UI/store calls to reach the host.
 * It deals in core/ data types (StyleSmithDocument, …) and confines all Office
 * JS to office/. The UI never touches Office globals; components call the store,
 * the store calls this bridge.
 *
 * Phase 2 scope: document load/save, shape scan, selection, linkage. The apply
 * engine (writing fill/line/text/geometry + chunked sweeps) extends this in
 * Phase 3.
 *
 * Every method returns a typed Result and never throws across the boundary.
 */
import { ok, err, type Result, type StyleSmithDocument } from "../core/schema/types";
import {
  isHostAvailable,
  runPowerPoint,
  officeErrorMessage,
  OperationCancelled,
  throwIfAborted,
} from "./context";
import { scanShapes, getSelectedShapeIds, type ScannedShape } from "./shapes";
import { queueTagWrite, readStyleId } from "./tags";
import { readChunkCurrent } from "./formatting";
import type { CurrentFormatting } from "../core/styles/diff";

/** An unlinked shape with its formatting, for fingerprint-based adoption. */
export interface AdoptionShape {
  readonly id: string;
  readonly name: string;
  readonly slideIndex: number;
  readonly formatting: CurrentFormatting;
}
import {
  loadDocument,
  saveDocument,
  clearDocument,
  ensureDocumentDirty,
  type StorageError,
} from "./storage";
import { readOfficeThemeMode, createDeckThemeResolver, type OfficeThemeMode } from "./theme";
import { runSweep, type SweepScope, type SweepOptions, type SweepResult } from "./sweep";
import { indexTokens } from "../core/tokens/resolve";
import { resolveStyleToTargets } from "../core/styles/diff";

export type { ScannedShape } from "./shapes";
export type { SweepScope, SweepResult } from "./sweep";

export type BridgeError =
  | { readonly code: "no-host" }
  | { readonly code: "cancelled" }
  | { readonly code: "office-error"; readonly message: string }
  | { readonly code: "invalid-document"; readonly issues: string[] };

function fromStorageError(e: StorageError): BridgeError {
  return e;
}

function fromThrown(e: unknown): BridgeError {
  if (e instanceof OperationCancelled) return { code: "cancelled" };
  return { code: "office-error", message: officeErrorMessage(e) };
}

export interface StyleSmithBridge {
  readonly isHostAvailable: boolean;

  /** Load the deck's StyleSmith document, or null if none stored yet. */
  loadDocument(): Promise<Result<StyleSmithDocument | null, BridgeError>>;
  /** Persist the document (settings.saveAsync self-persists; no dirty pair needed). */
  saveDocument(doc: StyleSmithDocument): Promise<Result<void, BridgeError>>;
  /** Remove the stored document. */
  clearDocument(): Promise<Result<void, BridgeError>>;

  /** Enumerate every top-level shape with its linkage tag (batched, cancellable). */
  scanShapes(opts?: { signal?: AbortSignal }): Promise<Result<ScannedShape[], BridgeError>>;
  /** Ids of the currently selected shapes. */
  getSelectedShapeIds(): Promise<Result<string[], BridgeError>>;

  /**
   * Capture the primary selected shape's current formatting (for "New from
   * selection", S1). Null when nothing is selected.
   */
  captureSelection(): Promise<Result<CurrentFormatting | null, BridgeError>>;

  /**
   * Tag the current selection with a style GUID and dirty the document so the
   * tag-only write survives on web. Returns the number of shapes linked.
   */
  linkSelectionToStyle(styleId: string): Promise<Result<number, BridgeError>>;

  /**
   * Apply a style to every shape carrying it within scope — the chunked sweep
   * (S3). Resolves the style's bound values against the stored document's tokens,
   * then writes in cancellable chunks with progress.
   */
  applyStyle(
    styleId: string,
    scope: SweepScope,
    options: SweepOptions,
  ): Promise<Result<SweepResult, BridgeError>>;

  /** The host's light/dark task-pane theme mode, or null outside a host. */
  readThemeMode(): OfficeThemeMode | null;

  /** Subscribe to PowerPoint selection changes. Returns an unsubscribe fn. */
  onSelectionChanged(handler: () => void): () => void;

  /** Read full current formatting for every UNLINKED, non-group shape (adoption). */
  readShapesFormatting(opts?: { signal?: AbortSignal }): Promise<Result<AdoptionShape[], BridgeError>>;

  /** Tag the given shapes (by id) with a style GUID; dirties the doc. Returns the count linked. */
  linkShapes(styleId: string, shapeIds: string[]): Promise<Result<number, BridgeError>>;
}

export const officeBridge: StyleSmithBridge = {
  get isHostAvailable() {
    return isHostAvailable();
  },

  async loadDocument() {
    const result = await loadDocument();
    return result.ok ? ok(result.value) : err(fromStorageError(result.error));
  },

  async saveDocument(doc) {
    const result = await saveDocument(doc);
    return result.ok ? ok(undefined) : err(fromStorageError(result.error));
  },

  async clearDocument() {
    const result = await clearDocument();
    return result.ok ? ok(undefined) : err(fromStorageError(result.error));
  },

  async scanShapes(opts) {
    try {
      return ok(await scanShapes(opts?.signal));
    } catch (e) {
      return err(fromThrown(e));
    }
  },

  async getSelectedShapeIds() {
    try {
      return ok(await getSelectedShapeIds());
    } catch (e) {
      return err(fromThrown(e));
    }
  },

  async captureSelection() {
    try {
      const captured = await runPowerPoint(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load("items/id");
        await ctx.sync();
        const shape = sel.items[0];
        if (!shape) return null;
        const [current] = await readChunkCurrent(ctx, [shape], {
          fill: true,
          line: true,
          text: true,
          geometry: true,
        });
        return current ?? null;
      });
      return ok(captured);
    } catch (e) {
      return err(fromThrown(e));
    }
  },

  async linkSelectionToStyle(styleId) {
    try {
      const count = await runPowerPoint(async (ctx) => {
        const sel = ctx.presentation.getSelectedShapes();
        sel.load("items");
        await ctx.sync();
        if (sel.items.length === 0) return 0;

        sel.items.forEach((shape) => queueTagWrite(shape, styleId));
        await ctx.sync();

        // Tag-only write → pair with a dirtying change so web autosaves it.
        await ensureDocumentDirty(ctx);
        return sel.items.length;
      });
      return ok(count);
    } catch (e) {
      return err(fromThrown(e));
    }
  },

  async applyStyle(styleId, scope, options) {
    try {
      const loaded = await loadDocument();
      if (!loaded.ok) return err(fromStorageError(loaded.error));
      const doc = loaded.value;
      if (!doc) return err({ code: "invalid-document", issues: ["no StyleSmith document in this deck"] });

      const style = doc.styles.find((s) => s.id === styleId);
      if (!style) return err({ code: "invalid-document", issues: [`style ${styleId} not found`] });

      const ctx = { tokens: indexTokens(doc), theme: createDeckThemeResolver() };
      const resolved = resolveStyleToTargets(style, ctx);

      return ok(await runSweep(styleId, resolved, scope, options));
    } catch (e) {
      return err(fromThrown(e));
    }
  },

  readThemeMode() {
    if (typeof window !== "undefined" && window.matchMedia?.("(forced-colors: active)").matches) {
      return "highContrast";
    }
    return readOfficeThemeMode();
  },

  onSelectionChanged(handler) {
    if (typeof Office === "undefined" || !Office.context?.document?.addHandlerAsync) {
      return () => {};
    }
    const doc = Office.context.document;
    doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);
    return () => {
      doc.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler });
    };
  },

  async readShapesFormatting(opts) {
    try {
      const result = await runPowerPoint(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();

        const cols = slides.items.map((slide) => {
          slide.shapes.load("items/id,items/name,items/type");
          return slide.shapes;
        });
        await ctx.sync();

        const flat: { shape: PowerPoint.Shape; slideIndex: number }[] = [];
        cols.forEach((col, slideIndex) => col.items.forEach((shape) => flat.push({ shape, slideIndex })));

        flat.forEach(({ shape }) => shape.tags.load("items/key,items/value"));
        await ctx.sync();

        // Adoption targets: unlinked, non-group shapes.
        const targets = flat.filter(
          ({ shape }) => readStyleId(shape) === null && shape.type !== PowerPoint.ShapeType.group,
        );

        const out: AdoptionShape[] = [];
        const CHUNK = 50;
        for (let i = 0; i < targets.length; i += CHUNK) {
          throwIfAborted(opts?.signal);
          const chunk = targets.slice(i, i + CHUNK);
          const formatting = await readChunkCurrent(
            ctx,
            chunk.map((c) => c.shape),
            { fill: true, line: true, text: true, geometry: true },
          );
          chunk.forEach((c, idx) =>
            out.push({ id: c.shape.id, name: c.shape.name, slideIndex: c.slideIndex, formatting: formatting[idx] ?? {} }),
          );
        }
        return out;
      });
      return ok(result);
    } catch (e) {
      return err(fromThrown(e));
    }
  },

  async linkShapes(styleId, shapeIds) {
    try {
      const linked = await runPowerPoint(async (ctx) => {
        const idSet = new Set(shapeIds);
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        const cols = slides.items.map((slide) => {
          slide.shapes.load("items/id");
          return slide.shapes;
        });
        await ctx.sync();

        const shapes = cols.flatMap((c) => c.items).filter((sh) => idSet.has(sh.id));
        let count = 0;
        const CHUNK = 100;
        for (let i = 0; i < shapes.length; i += CHUNK) {
          shapes.slice(i, i + CHUNK).forEach((sh) => queueTagWrite(sh, styleId));
          await ctx.sync();
          count += Math.min(CHUNK, shapes.length - i);
        }
        // Tag-only writes → dirty once so they survive on web.
        if (shapes.length > 0) await ensureDocumentDirty(ctx);
        return count;
      });
      return ok(linked);
    } catch (e) {
      return err(fromThrown(e));
    }
  },
};
