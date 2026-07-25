/**
 * The chunked apply sweep — the performance-critical heart of Phase 3.
 *
 * ⚠️ Unbatched bulk writes hang the host (CLAUDE.md): 420 shapes × 3 props in one
 * sync never returns. So writes are chunked, one ctx.sync() per chunk, with a
 * progress callback and cancellation at chunk boundaries. Reads are batched.
 *
 * Flow (one PowerPoint.run):
 *   1. enumerate every shape + tags (3 syncs), filter to the style + scope
 *   2. per chunk: read apply gates (hasText / adjustment count); if preserving
 *      overrides, read current formatting; build a plan per shape (core); apply
 *      supported+gated targets; queue the snapshot tag; sync
 *   3. dirty the document once so snapshot (tag-only) writes survive on web
 *
 * The style is resolved to concrete targets by core (resolveStyleToTargets); this
 * module only reads/writes shapes and orchestrates chunks.
 */
import { runPowerPoint, OperationCancelled } from "./context";
import { readStyleId, queueTagWrite } from "./tags";
import { ensureDocumentDirty } from "./storage";
import {
  isSupportedTarget,
  applyTargetToShape,
  readApplyGates,
  readChunkCurrent,
  readSnapshot,
  queueSnapshotWrite,
} from "./formatting";
import { planApply, type ResolvedStyle, type CurrentFormatting } from "../core/styles/diff";
import type { LayerName } from "../core/schema/types";

export type SweepScope =
  | { readonly kind: "deck" }
  | { readonly kind: "slides"; readonly slideIndices: number[] }
  | { readonly kind: "shapes"; readonly shapeIds: string[] };

export interface SweepOptions {
  readonly chunkSize?: number; // default 50 (tune per gate)
  readonly preserveOverrides: boolean;
  /** If set, apply only these layers (selective application, S12). */
  readonly layers?: readonly LayerName[];
  readonly signal?: AbortSignal;
  readonly onProgress?: (done: number, total: number) => void;
}

export interface SweepResult {
  /** Shapes carrying the style within scope. */
  readonly matched: number;
  /** Shapes we wrote at least one property to. */
  readonly applied: number;
  /** Total properties preserved as local overrides across all shapes. */
  readonly preservedProps: number;
  /** Shapes where a geometry layer was skipped (no adjustable handles — AC2.3). */
  readonly skippedGeometryShapes: number;
  /** Style property paths that could not be resolved (e.g. unreadable theme colour). */
  readonly unresolvedPaths: string[];
  /** Supported-but-deferred paths (deep text — Phase 4). */
  readonly deferredPaths: string[];
  readonly cancelled: boolean;
  readonly chunkSize: number;
  readonly elapsedMs: number;
  /** Shapes applied per second — the throughput figure to record at the gate. */
  readonly throughput: number;
}

const DEFAULT_CHUNK_SIZE = 50;

function inScope(shapeId: string, slideIndex: number, scope: SweepScope): boolean {
  switch (scope.kind) {
    case "deck":
      return true;
    case "slides":
      return scope.slideIndices.includes(slideIndex);
    case "shapes":
      return scope.shapeIds.includes(shapeId);
  }
}

/**
 * Apply a resolved style to every shape carrying `styleId` within `scope`.
 * `resolved` is the full resolution; supported targets are selected here.
 */
export async function runSweep(
  styleId: string,
  resolved: ResolvedStyle,
  scope: SweepScope,
  options: SweepOptions,
): Promise<SweepResult> {
  const chunkSize = options.chunkSize ?? DEFAULT_CHUNK_SIZE;
  const { preserveOverrides, signal, onProgress } = options;

  const layerFilter = options.layers ? new Set<LayerName>(options.layers) : null;
  const supportedTargets = resolved.targets
    .filter(isSupportedTarget)
    .filter((t) => !layerFilter || layerFilter.has(t.layer));
  const supported: ResolvedStyle = { targets: supportedTargets, unresolved: resolved.unresolved };
  const deferredPaths = resolved.targets.filter((t) => !isSupportedTarget(t)).map((t) => t.path);
  const unresolvedPaths = resolved.unresolved.map((u) => u.path);

  const needText = supportedTargets.some((t) => t.layer === "text");
  const needGeometry = supportedTargets.some((t) => t.layer === "geometry");
  const needFill = supportedTargets.some((t) => t.layer === "shape" && t.property === "fill");
  const needLine = supportedTargets.some(
    (t) => t.layer === "shape" && t.property !== "fill",
  );

  const start = performance.now();
  let applied = 0;
  let preservedProps = 0;
  let skippedGeometryShapes = 0;
  let cancelled = false;

  const matched = await runPowerPoint(async (ctx) => {
    // 1 — enumerate every shape with id + tags, filter to style + scope.
    const slides = ctx.presentation.slides;
    slides.load("items");
    await ctx.sync();

    const cols = slides.items.map((slide) => {
      slide.shapes.load("items/id");
      return slide.shapes;
    });
    await ctx.sync();

    const all: { shape: PowerPoint.Shape; slideIndex: number }[] = [];
    cols.forEach((col, slideIndex) =>
      col.items.forEach((shape) => all.push({ shape, slideIndex })),
    );

    all.forEach(({ shape }) => shape.tags.load("items/key,items/value"));
    await ctx.sync();

    // "shapes" scope (Apply to selection, near-match link+normalise) targets the
    // given shapes REGARDLESS of their current tag — they get linked below. This
    // is what applies a style to newly-selected, not-yet-linked shapes (AC2.1).
    // Deck/slides scope re-applies to shapes already carrying the style.
    const explicitIds = scope.kind === "shapes" ? new Set(scope.shapeIds) : null;
    const targets = all.filter(({ shape, slideIndex }) =>
      explicitIds
        ? explicitIds.has(shape.id)
        : readStyleId(shape) === styleId && inScope(shape.id, slideIndex, scope),
    );

    // 2 — chunked apply.
    let anyWrite = false;
    for (let i = 0; i < targets.length; i += chunkSize) {
      if (signal?.aborted) {
        cancelled = true;
        break;
      }
      const chunk = targets.slice(i, i + chunkSize).map((t) => t.shape);

      const gates = await readApplyGates(ctx, chunk, needText, needGeometry);
      let current: CurrentFormatting[] = [];
      if (preserveOverrides) {
        current = await readChunkCurrent(ctx, chunk, {
          fill: needFill,
          line: needLine,
          text: needText,
          geometry: needGeometry,
        });
      }

      chunk.forEach((shape, idx) => {
        const gate = gates[idx] ?? { hasText: false, adjustmentCount: 0 };
        const snapshot = readSnapshot(shape);
        const plan = planApply(supported, snapshot, current[idx] ?? {}, preserveOverrides);
        preservedProps += plan.preserved.length;

        let wroteShape = false;
        let geometrySkipped = false;
        for (const target of plan.toApply) {
          if (target.layer === "text" && !gate.hasText) continue;
          if (target.layer === "geometry") {
            if (gate.adjustmentCount === 0) {
              geometrySkipped = true;
              continue;
            }
            if (target.index !== undefined && target.index >= gate.adjustmentCount) continue;
          }
          applyTargetToShape(shape, target);
          wroteShape = true;
        }
        if (geometrySkipped) skippedGeometryShapes += 1;
        if (wroteShape) applied += 1;

        // Link the shape to this style. Idempotent on re-apply; replaces a
        // different style's tag (AC2.5). This is what makes "Apply to selection"
        // link new shapes instead of silently skipping them.
        queueTagWrite(shape, styleId);
        // Record what we (would have) set so later applies can detect overrides.
        queueSnapshotWrite(shape, plan.nextSnapshot);
        anyWrite = true;
      });

      await ctx.sync();
      onProgress?.(Math.min(i + chunkSize, targets.length), targets.length);
    }

    // 3 — snapshot tags are metadata; dirty once so they survive on web.
    if (anyWrite) {
      try {
        await ensureDocumentDirty(ctx);
      } catch (e) {
        if (e instanceof OperationCancelled) cancelled = true;
        // a failed dirty must not lose the applied formatting — swallow otherwise
      }
    }

    return targets.length;
  });

  const elapsedMs = performance.now() - start;
  return {
    matched,
    applied,
    preservedProps,
    skippedGeometryShapes,
    unresolvedPaths,
    deferredPaths,
    cancelled,
    chunkSize,
    elapsedMs,
    throughput: elapsedMs > 0 ? (applied / elapsedMs) * 1000 : 0,
  };
}
