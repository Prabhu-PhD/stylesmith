/**
 * Batched shape enumeration. Measured: 2,520 shapes across 60 slides in ~2.0s
 * with exactly 3 syncs regardless of deck size (CLAUDE.md). NEVER sync per slide.
 *
 * Returns PLAIN, serialisable data — not live proxies (proxies don't survive
 * across PowerPoint.run). core/ (adoption, usage) consumes this; write paths
 * re-enumerate their own proxies inside their own run.
 *
 * NOTE: this enumerates TOP-LEVEL shapes. Group children are reachable
 * (Phase 0 spike: shape.group.shapes), but descending into groups is a sweep
 * concern deferred to Phase 3 — recorded so it is added deliberately, not missed.
 */
import { runPowerPoint, throwIfAborted } from "./context";
import { readStyleId } from "./tags";

export interface ScannedShape {
  readonly id: string;
  readonly name: string;
  /** PowerPoint.ShapeType as a string (e.g. "geometricShape", "group"). */
  readonly type: string;
  readonly slideIndex: number;
  /** The linked style GUID from the STYLESMITH_ID tag, or null if unlinked. */
  readonly styleId: string | null;
  readonly isGroup: boolean;
}

/** Enumerate every top-level shape in the deck with its linkage tag (3 syncs). */
export async function scanShapes(signal?: AbortSignal): Promise<ScannedShape[]> {
  return runPowerPoint(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items");
    await ctx.sync(); // 1
    throwIfAborted(signal);

    const shapeCols = slides.items.map((slide) => {
      slide.shapes.load("items/id,items/name,items/type");
      return slide.shapes;
    });
    await ctx.sync(); // 2
    throwIfAborted(signal);

    const flat: { shape: PowerPoint.Shape; slideIndex: number }[] = [];
    shapeCols.forEach((col, slideIndex) =>
      col.items.forEach((shape) => flat.push({ shape, slideIndex })),
    );

    flat.forEach(({ shape }) => shape.tags.load("items/key,items/value"));
    await ctx.sync(); // 3
    throwIfAborted(signal);

    return flat.map(({ shape, slideIndex }) => ({
      id: shape.id,
      name: shape.name,
      type: String(shape.type),
      slideIndex,
      styleId: readStyleId(shape),
      isGroup: shape.type === PowerPoint.ShapeType.group,
    }));
  });
}

/** Ids of the currently selected shapes (for scope resolution / linking). */
export async function getSelectedShapeIds(): Promise<string[]> {
  return runPowerPoint(async (ctx) => {
    const sel = ctx.presentation.getSelectedShapes();
    sel.load("items/id");
    await ctx.sync();
    return sel.items.map((s) => s.id);
  });
}
