/**
 * VERIFIED OFFICE JS PATTERNS — StyleSmith
 *
 * Every pattern below was confirmed working against a live PowerPoint host
 * (PowerPoint web / OfficeOnline, PowerPointApi 1.10) on 23 July 2026.
 *
 * This file is REFERENCE, not production code. It exists because several of
 * these APIs do not behave the way the documented/obvious patterns suggest,
 * and rediscovering that costs hours.
 *
 * See CLAUDE.md for the full rules.
 */

/* ────────────────────────────────────────────────────────────────────────
   1. ADJUSTMENTS — the geometry layer
   ────────────────────────────────────────────────────────────────────────
   Adjustments is NOT a standard Office JS collection.

   Actual members: count, get, set, load, toJSON
   _scalarPropertyNames === ["count"]

   There is NO items array, NO getItemAt, NO getCount.
   get(i) returns a ClientResult<number> — you sync, then read .value.
   It has no load() method.
   ──────────────────────────────────────────────────────────────────────── */

async function readAdjustments(
  ctx: PowerPoint.RequestContext,
  shape: PowerPoint.Shape
): Promise<number[]> {
  shape.adjustments.load("count");
  await ctx.sync();

  const count = shape.adjustments.count;
  if (!count) return [];

  // get() returns ClientResult — queue them all, sync once, then read .value
  const results = Array.from({ length: count }, (_, i) => shape.adjustments.get(i));
  await ctx.sync();

  return results.map(r => r.value);
}

function queueAdjustmentWrite(shape: PowerPoint.Shape, index: number, value: number): void {
  // Verified: 0.16667 -> 0.15, confirmed by read-back.
  // Caller syncs at the chunk boundary.
  shape.adjustments.set(index, value);
}

/* ❌ NONE OF THESE WORK — do not generate them:
     shape.adjustments.getCount()          // not a function
     shape.adjustments.getItemAt(0)        // not a function
     shape.adjustments.items[0]            // items is undefined
     shape.adjustments.load("items")       // succeeds, yields nothing
     shape.adjustments.setItemAt(0, v)     // not a function
     shape.adjustments.set({ items: [v] }) // throws InvalidArgument
     clientResult.load("value")            // ClientResult has no load()

   ⚠️ A write that does not throw has NOT necessarily succeeded.
      Always verify by reading the value back.
*/

/* ────────────────────────────────────────────────────────────────────────
   2. BATCHED ENUMERATION — always 3 syncs, regardless of deck size
   ────────────────────────────────────────────────────────────────────────
   Measured: 2,520 shapes across 60 slides in ~2.0s.
   Syncing per slide instead is dramatically slower.
   ──────────────────────────────────────────────────────────────────────── */

async function collectAllShapes(
  ctx: PowerPoint.RequestContext,
  withTags: boolean
): Promise<PowerPoint.Shape[]> {
  const slides = ctx.presentation.slides;
  slides.load("items");
  await ctx.sync();                                    // sync 1

  const collections = slides.items.map(slide => {
    slide.shapes.load("items/name,items/id,items/type");
    return slide.shapes;
  });
  await ctx.sync();                                    // sync 2

  const all = collections.flatMap(c => c.items);

  if (withTags) {
    all.forEach(s => s.tags.load("items/key,items/value"));
    await ctx.sync();                                  // sync 3
  }

  return all;
}

/* ────────────────────────────────────────────────────────────────────────
   3. TAGS — the linkage layer
   ────────────────────────────────────────────────────────────────────────
   Write + read confirmed. Persistence confirmed at 100% ACROSS close/reopen
   — but ONLY when the document is dirtied. See section 5.

   ⚠️ Tags do NOT survive copy/paste between decks. Confirmed.
      Do not design around them doing so.
   ──────────────────────────────────────────────────────────────────────── */

const TAG_KEY = "STYLESMITH_ID";

function queueTagWrite(shape: PowerPoint.Shape, styleId: string): void {
  shape.tags.add(TAG_KEY, styleId);
}

function readStyleId(shape: PowerPoint.Shape): string | null {
  try {
    const tag = shape.tags.items.find(t => t.key === TAG_KEY);
    return tag ? tag.value : null;
  } catch {
    return null; // tags not loaded
  }
}

/* ────────────────────────────────────────────────────────────────────────
   4. CHUNKED WRITES — mandatory, not an optimisation
   ────────────────────────────────────────────────────────────────────────
   Measured: 420 shapes x 3 properties queued into ONE ctx.sync() HANGS
   the host. No error, no completion — it simply never returns.
   ──────────────────────────────────────────────────────────────────────── */

interface SweepOptions {
  chunkSize: number;                                     // start at 50, tune
  onProgress: (done: number, total: number) => void;
  signal: AbortSignal;
}

async function chunkedSweep(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
  applyToShape: (s: PowerPoint.Shape) => void,
  opts: SweepOptions
): Promise<{ completed: number; cancelled: boolean }> {
  let done = 0;

  for (let i = 0; i < shapes.length; i += opts.chunkSize) {
    if (opts.signal.aborted) {
      return { completed: done, cancelled: true };      // clean boundary stop
    }

    const chunk = shapes.slice(i, i + opts.chunkSize);
    chunk.forEach(applyToShape);
    await ctx.sync();                                    // one sync per chunk

    done += chunk.length;
    opts.onProgress(done, shapes.length);
  }

  return { completed: done, cancelled: false };
}

/* ────────────────────────────────────────────────────────────────────────
   5. ⚠️ THE DIRTY-FLAG RULE
   ────────────────────────────────────────────────────────────────────────
   Metadata-only writes are SILENTLY LOST on PowerPoint web.

   Evidence: tags written, read back 4/4 in-session, document closed and
   reopened -> 0 tags. Same test with a shape nudged first -> 100% survival.

   A write that changes only tags or only a custom XML part does not mark
   the document modified, so autosave never fires.

   Implement ONCE, centrally. Never ad hoc at call sites.

   OPEN: not yet verified whether writing a custom XML part dirties the
   document. Assume it does NOT until proven. Spike in Phase 0.
   ──────────────────────────────────────────────────────────────────────── */

async function ensureDocumentDirty(
  ctx: PowerPoint.RequestContext,
  shape: PowerPoint.Shape
): Promise<void> {
  // Re-set a property to its existing value so the host marks the document
  // modified without introducing any visible change.
  // Candidate approach — verify no visual side effect before shipping.
  shape.load("left");
  await ctx.sync();
  shape.left = shape.left;
  await ctx.sync();
}

/* ────────────────────────────────────────────────────────────────────────
   6. SHAPE MEMBERS — what exists and what does not
   ────────────────────────────────────────────────────────────────────────
   Confirmed by enumerating all 81 members of PowerPoint.Shape.

   ✅ EXISTS:
      adjustments, fill, lineFormat, textFrame, tags, getTable,
      customXmlParts, creationId, getParentSlide, group, parentGroup,
      id, name, type, left, top, width, height, rotation,
      delete, duplicate, setZOrder, setHyperlink, altTextDescription,
      isDecorative, level, placeholderFormat, visible, zOrderPosition

   ❌ DOES NOT EXIST — do not attempt:
      shadow                — no shadow/glow/reflection/3-D in the style model
      effects               — same
      geometricShapeType    — styles TUNE shapes; they cannot convert types

   ⚠️ Groups return as a single shape. Children are NOT enumerated by
      slide.shapes.items. Whether children are reachable at all is an
      OPEN QUESTION — spike in Phase 0. If unreachable, sweeps must report
      "N shapes inside groups were skipped" rather than silently missing them.
   ──────────────────────────────────────────────────────────────────────── */

/* ────────────────────────────────────────────────────────────────────────
   7. HOST CAPABILITY DETECTION
   ────────────────────────────────────────────────────────────────────────
   Measured on PowerPoint web: 1.1 through 1.10 supported, 1.11+ absent.
   Manifest minimum should be 1.1; feature-detect anything above.
   ──────────────────────────────────────────────────────────────────────── */

function supportsApi(version: string): boolean {
  return Office.context.requirements.isSetSupported("PowerPointApi", version);
}

/* ────────────────────────────────────────────────────────────────────────
   8. SELECTION
   ──────────────────────────────────────────────────────────────────────── */

async function getSelectedShapes(
  ctx: PowerPoint.RequestContext
): Promise<PowerPoint.Shape[]> {
  const sel = ctx.presentation.getSelectedShapes();
  sel.load("items/name,items/id,items/type");
  await ctx.sync();
  return sel.items;
}

export {
  readAdjustments,
  queueAdjustmentWrite,
  collectAllShapes,
  queueTagWrite,
  readStyleId,
  chunkedSweep,
  ensureDocumentDirty,
  supportsApi,
  getSelectedShapes,
  TAG_KEY,
};
