/**
 * Reading and writing a shape's formatting for the apply engine.
 *
 * Phase 3 applies the layers on verified APIs: shape fill/border, geometry
 * adjustments, and basic text (font name/size/colour/bold). Deeper text
 * (line spacing, space before/after, alignment, bullets) is the documented
 * "textFrame depth" open question, deferred to Phase 4 — such targets are
 * reported as deferred rather than applied.
 *
 * ⚠️ API ASSUMPTIONS (to confirm on the host harness): reading current fill via
 * `fill.type`/`fill.foregroundColor`, and reading `lineFormat`/`font` scalar
 * properties. Writes (setSolidColor, lineFormat.*, font.*, adjustments.set) are
 * from the validated contract (CLAUDE.md / PRD §9.1). The preserve-off apply
 * path does no current-formatting read, so the gate does not depend on the read
 * assumptions.
 *
 * ⚠️ themeColor applies as a resolved hex (decision: "hex now, accept the gap").
 * It does NOT cascade without the add-in (AC28.2/28.3) — a known v1 limitation.
 */
import type { CurrentFormatting, Snapshot, TargetValue } from "../core/styles/diff";

const SNAPSHOT_TAG_KEY = "STYLESMITH_SNAPSHOT";

/** Whether the apply engine can write this target in Phase 3. */
export function isSupportedTarget(target: TargetValue): boolean {
  if (target.layer === "geometry") return target.property === "adjustments";
  if (target.layer === "text") {
    return ["fontFamily", "fontSize", "color", "fontWeight"].includes(target.property);
  }
  if (target.layer === "shape") {
    return ["fill", "borderColor", "borderWeight", "borderDashStyle"].includes(target.property);
  }
  return false; // table deferred (Phase 5)
}

function isNone(value: string | number): boolean {
  return typeof value === "string" && value.trim().toLowerCase() === "none";
}

/**
 * Queue the writes for a single target on a shape. Caller has already gated
 * text targets to text-bearing shapes and geometry indices to available handles.
 * No sync here — the sweep syncs per chunk.
 */
export function applyTargetToShape(shape: PowerPoint.Shape, target: TargetValue): void {
  switch (target.layer) {
    case "shape":
      applyShapeTarget(shape, target);
      return;
    case "text":
      applyTextTarget(shape, target);
      return;
    case "geometry":
      if (target.property === "adjustments" && target.index !== undefined) {
        shape.adjustments.set(target.index, Number(target.value));
      }
      return;
    case "table":
      return; // deferred
  }
}

function applyShapeTarget(shape: PowerPoint.Shape, target: TargetValue): void {
  switch (target.property) {
    case "fill":
      if (isNone(target.value)) shape.fill.clear();
      else shape.fill.setSolidColor(String(target.value));
      return;
    case "borderColor":
      shape.lineFormat.color = String(target.value);
      return;
    case "borderWeight":
      shape.lineFormat.weight = Number(target.value);
      return;
    case "borderDashStyle":
      shape.lineFormat.dashStyle = String(target.value) as PowerPoint.ShapeLineDashStyle;
      return;
  }
}

function applyTextTarget(shape: PowerPoint.Shape, target: TargetValue): void {
  const font = shape.textFrame.textRange.font;
  switch (target.property) {
    case "fontFamily":
      font.name = String(target.value);
      return;
    case "fontSize":
      font.size = Number(target.value);
      return;
    case "color":
      font.color = String(target.value);
      return;
    case "fontWeight":
      font.bold = typeof target.value === "number" ? target.value >= 600 : String(target.value).toLowerCase() === "bold";
      return;
  }
}

/* ── Apply gates — cheap reads that make applying SAFE ───────────────────────
   Applying a text target to a text-less shape, or a geometry index a shape does
   not have, throws and would fail the whole chunk sync. We read hasText and the
   adjustment count first so the sweep can skip those safely (AC2.3/AC2.4). */

export interface ApplyGate {
  readonly hasText: boolean;
  readonly adjustmentCount: number;
}

/** Load + read per-shape apply gates for a chunk (one sync). */
export async function readApplyGates(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
  needText: boolean,
  needGeometry: boolean,
): Promise<ApplyGate[]> {
  shapes.forEach((shape) => {
    if (needText) shape.textFrame.load("hasText");
    if (needGeometry) shape.adjustments.load("count");
  });
  await ctx.sync();

  return shapes.map((shape) => ({
    hasText: needText ? safeBool(() => shape.textFrame.hasText) : false,
    adjustmentCount: needGeometry ? safeNum(() => shape.adjustments.count) : 0,
  }));
}

/* ── Current formatting read (preserve-overrides path only) ─────────────────
   ⚠️ Contains the read assumptions noted at the top of this file. */

export interface CurrentReadOptions {
  readonly fill: boolean;
  readonly line: boolean;
  readonly text: boolean;
  readonly geometry: boolean;
}

export async function readChunkCurrent(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
  opts: CurrentReadOptions,
): Promise<CurrentFormatting[]> {
  shapes.forEach((shape) => {
    if (opts.fill) shape.fill.load("type,foregroundColor");
    if (opts.line) shape.lineFormat.load("color,weight,dashStyle");
    if (opts.text) {
      shape.textFrame.load("hasText");
      shape.textFrame.textRange.font.load("name,size,color,bold");
    }
    if (opts.geometry) shape.adjustments.load("count");
  });
  await ctx.sync();

  // Adjustments use the ClientResult get() pattern — a second sync.
  const adjResults = opts.geometry
    ? shapes.map((shape) => {
        const n = safeNum(() => shape.adjustments.count);
        return Array.from({ length: n }, (_, i) => shape.adjustments.get(i));
      })
    : [];
  if (opts.geometry) await ctx.sync();

  return shapes.map((shape, idx) => {
    const cur: CurrentFormatting = {};
    const setIf = (key: string, value: string | number | null | undefined): void => {
      if (value !== null && value !== undefined) cur[key] = value;
    };
    if (opts.fill) {
      try {
        const type = String(shape.fill.type).toLowerCase();
        if (type.includes("nofill")) cur["shape.fill"] = "none";
        else setIf("shape.fill", shape.fill.foregroundColor);
      } catch {
        /* fill unreadable — leave absent (override detection falls back to apply) */
      }
    }
    if (opts.line) {
      try {
        setIf("shape.borderColor", shape.lineFormat.color);
        setIf("shape.borderWeight", shape.lineFormat.weight);
        setIf("shape.borderDashStyle", shape.lineFormat.dashStyle == null ? null : String(shape.lineFormat.dashStyle));
      } catch {
        /* line unreadable */
      }
    }
    if (opts.text) {
      try {
        if (shape.textFrame.hasText) {
          const font = shape.textFrame.textRange.font;
          setIf("text.fontFamily", font.name);
          setIf("text.fontSize", font.size);
          setIf("text.color", font.color);
          if (font.bold !== null && font.bold !== undefined) {
            cur["text.fontWeight"] = font.bold ? 700 : 400;
          }
        }
      } catch {
        /* text unreadable */
      }
    }
    if (opts.geometry) {
      (adjResults[idx] ?? []).forEach((r, i) => {
        cur[`geometry.adjustments[${i}]`] = r.value;
      });
    }
    return cur;
  });
}

/* ── Snapshot tag (per-shape last-applied values) ───────────────────────────
   Enables override detection (AC13.1). Tag-only writes need the dirty guard —
   the sweep pairs snapshot writes with ensureDocumentDirty. */

export function readSnapshot(shape: PowerPoint.Shape): Snapshot | null {
  try {
    const tag = shape.tags.items.find((t) => t.key === SNAPSHOT_TAG_KEY);
    if (!tag) return null;
    return JSON.parse(tag.value) as Snapshot;
  } catch {
    return null;
  }
}

export function queueSnapshotWrite(shape: PowerPoint.Shape, snapshot: Snapshot): void {
  shape.tags.add(SNAPSHOT_TAG_KEY, JSON.stringify(snapshot));
}

function safeBool(read: () => boolean): boolean {
  try {
    return read();
  } catch {
    return false;
  }
}

function safeNum(read: () => number): number {
  try {
    return read() || 0;
  } catch {
    return 0;
  }
}
