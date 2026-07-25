/**
 * The geometry layer — adjustment handle values.
 *
 * ⚠️ Adjustments is NOT a standard Office JS collection (CLAUDE.md). Members:
 * count, get, set, load, toJSON. `get(i)` returns a ClientResult<number> — you
 * sync, then read `.value`; it has NO load(). A write that doesn't throw has not
 * necessarily succeeded — verify by read-back. All of this is confined here.
 */

/**
 * Read all adjustment values for a shape. Returns [] for shapes with no
 * adjustable handles (e.g. a plain rectangle) — which is also why applying a
 * geometry layer to such a shape must skip gracefully (AC2.3).
 */
export async function readAdjustments(
  ctx: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
): Promise<number[]> {
  shape.adjustments.load("count");
  await ctx.sync();

  const count = shape.adjustments.count;
  if (!count) return [];

  // get() returns ClientResult — queue all, sync once, then read .value.
  const results = Array.from({ length: count }, (_, i) => shape.adjustments.get(i));
  await ctx.sync();
  return results.map((r) => r.value);
}

/**
 * Queue a single adjustment write. The caller syncs at the chunk boundary
 * (never queue an unbounded number of writes into one sync — it hangs the host).
 */
export function queueAdjustmentWrite(shape: PowerPoint.Shape, index: number, value: number): void {
  shape.adjustments.set(index, value);
}

/** Read back an adjustment to verify a write actually took (no false positives). */
export async function readAdjustment(
  ctx: PowerPoint.RequestContext,
  shape: PowerPoint.Shape,
  index: number,
): Promise<number> {
  const result = shape.adjustments.get(index);
  await ctx.sync();
  return result.value;
}
